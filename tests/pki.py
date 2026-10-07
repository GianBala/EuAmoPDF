"""AC e assinantes descartáveis, para gerar PDFs assinados sem usar documentos de ninguém. Usado pelos
testes e pelo teste de fumaça do build (packaging/build.py), que confere o pyHanko dentro do executável."""
import io
import types
from datetime import datetime, timedelta, timezone

from asn1crypto import crl, keys
from asn1crypto import x509 as asn1_x509
from cryptography import x509
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives.serialization import Encoding, NoEncryption, PrivateFormat
from cryptography.x509.oid import NameOID, ObjectIdentifier
from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
from pyhanko.sign import signers
from pyhanko.sign.fields import MDPPerm, SigFieldSpec
from pyhanko_certvalidator.registry import SimpleCertificateStore

GOOD_CPF = "52998224725"  # CPF de exemplo com dígitos verificadores válidos
BIRTH = "15031990"


def _issue(issuer_name, issuer_key, name, key, *, ca, valid=(-30, 335), personal=None, policy=True, org=None):
    """Certificado como o do gov.br: só "assinatura digital" no uso da chave e, na pessoa, o OID
    2.16.76.1.3.1 com a data de nascimento e o CPF. `personal` é o valor do OID (bytes, em OCTET STRING)."""
    now = datetime.now(timezone.utc)
    subject = x509.Name(([x509.NameAttribute(NameOID.ORGANIZATION_NAME, org)] if org else [])
                        + [x509.NameAttribute(NameOID.COMMON_NAME, name)])
    builder = (x509.CertificateBuilder().subject_name(subject).issuer_name(issuer_name or subject)
               .public_key(key.public_key()).serial_number(x509.random_serial_number())
               .not_valid_before(now + timedelta(days=valid[0])).not_valid_after(now + timedelta(days=valid[1]))
               .add_extension(x509.BasicConstraints(ca=ca, path_length=None), critical=True))
    if not ca:
        builder = builder.add_extension(x509.KeyUsage(
            digital_signature=True, content_commitment=False, key_encipherment=False, data_encipherment=False,
            key_agreement=False, key_cert_sign=False, crl_sign=False, encipher_only=False, decipher_only=False),
            critical=True)
        if policy:  # a política do certificado de assinatura de pessoa do gov.br
            builder = builder.add_extension(x509.CertificatePolicies(
                [x509.PolicyInformation(ObjectIdentifier("2.16.76.3.2.1.1"), None)]), critical=False)
        if personal is not None:
            other = x509.OtherName(ObjectIdentifier("2.16.76.1.3.1"), bytes([0x04, len(personal)]) + personal)
            builder = builder.add_extension(x509.SubjectAlternativeName([other]), critical=False)
    return builder.sign(issuer_key or key, hashes.SHA256())


def _pair(key, cert):
    return types.SimpleNamespace(key=key, cert=cert, asn1=asn1_x509.Certificate.load(cert.public_bytes(Encoding.DER)))


def new_authority(name, org=None):
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    return _pair(key, _issue(None, None, name, key, ca=True, org=org))


def issue_person(authority, name="MARIA DE TESTE", **options):
    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    options.setdefault("personal", (BIRTH + GOOD_CPF + "0" * 26).encode())
    return _pair(key, _issue(authority.cert.subject, authority.key, name, key, ca=False, **options))


def sign_pdf(person, pdf, field_name="Signature1", box=None, certify=None):
    """`box` cria o campo visível na página 1 (como o do gov.br); `certify` (um MDPPerm) faz uma assinatura de
    certificação, a primeira de um documento que depois recebe outras."""
    signer = signers.SimpleSigner(
        signing_cert=person.asn1, cert_registry=SimpleCertificateStore(),
        signing_key=keys.PrivateKeyInfo.load(person.key.private_bytes(Encoding.DER, PrivateFormat.PKCS8, NoEncryption())))
    writer = IncrementalPdfFileWriter(io.BytesIO(pdf))
    meta = signers.PdfSignatureMetadata(field_name=field_name, md_algorithm="sha256", certify=certify is not None,
                                        docmdp_permissions=certify or MDPPerm.FILL_FORMS)
    visible = SigFieldSpec(field_name, box=box, on_page=0) if box else None
    return signers.sign_pdf(writer, meta, signer=signer, new_field_spec=visible).getvalue()


def crl_of(authority, revoked=(), when=None):
    now = datetime.now(timezone.utc)
    builder = (x509.CertificateRevocationListBuilder().issuer_name(authority.cert.subject)
               .last_update(now - timedelta(days=1)).next_update(now + timedelta(days=1)))
    for serial in revoked:
        builder = builder.add_revoked_certificate(
            x509.RevokedCertificateBuilder().serial_number(serial).revocation_date(when or now - timedelta(hours=1)).build())
    return crl.CertificateList.load(builder.sign(authority.key, hashes.SHA256()).public_bytes(Encoding.DER))


def timestamp_pdf(authority, pdf):
    """Acrescenta ao PDF um carimbo de tempo do documento, de uma TSA de teste."""
    from pyhanko.sign.timestamps.dummy_client import DummyTimeStamper
    tsa = issue_person(authority, name="TSA DE TESTE")
    stamper = DummyTimeStamper(
        tsa_cert=tsa.asn1, certs_to_embed=[authority.asn1],
        tsa_key=keys.PrivateKeyInfo.load(tsa.key.private_bytes(Encoding.DER, PrivateFormat.PKCS8, NoEncryption())))
    return signers.PdfTimeStamper(stamper).timestamp_pdf(IncrementalPdfFileWriter(io.BytesIO(pdf)), "sha256").getvalue()


def append_update(pdf, change):
    """Uma atualização incremental depois da assinatura ("page", "title" ou "field"), escrita pelo próprio pyHanko: o
    saveIncr do PyMuPDF 1.25 não consegue regravar um PDF que ele assinou ("cannot find object in xref")."""
    from pyhanko.pdf_utils import generic
    from pyhanko.pdf_utils.writer import PageObject
    from pyhanko.sign.fields import SigFieldSpec, append_signature_field
    writer = IncrementalPdfFileWriter(io.BytesIO(pdf))
    if change == "page":
        writer.insert_page(PageObject(contents=writer.add_object(generic.StreamObject(stream_data=b"")), media_box=(0, 0, 595, 842)))
    elif change == "title":
        writer.set_info(generic.DictionaryObject({generic.NameObject("/Title"): generic.TextStringObject("outro")}))
    else:
        append_signature_field(writer, SigFieldSpec("Signature2", box=(72, 700, 272, 760), on_page=0))
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()
