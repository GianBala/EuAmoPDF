import io
import tempfile
import zipfile
from pathlib import Path

import pymupdf
import pytest
from PIL import Image

import app as euamopdf


@pytest.fixture
def client():
    return euamopdf.app.test_client()


def make_pdf(pages=3, **save_options):
    doc = pymupdf.open()
    for i in range(pages):
        doc.new_page().insert_text((72, 72), f"Pagina {i + 1}")
    return doc.tobytes(**save_options)


def make_image(mode="RGB", size=(200, 100), color=(0, 0, 255), fmt="PNG", **save_options):
    buf = io.BytesIO()
    Image.new(mode, size, color).save(buf, fmt, **save_options)
    return buf.getvalue()


def post(client, action, *files, headers=None, **form):
    """files: pares (nome, conteúdo)."""
    data = {"action": action, "file": [(io.BytesIO(content), name) for name, content in files], **form}
    return client.post("/convert", data=data, content_type="multipart/form-data", headers=headers or {})


def texts(pdf):
    return [page.get_text().strip() for page in pymupdf.open(stream=pdf, filetype="pdf")]


def unzip(data):
    z = zipfile.ZipFile(io.BytesIO(data))
    return {name: z.read(name) for name in z.namelist()}


# --- Organizar ---

def test_merge_keeps_upload_order(client):
    r = post(client, "merge-pdf", ("b.pdf", make_pdf(2)), ("a.pdf", make_pdf(1)))
    assert r.status_code == 200
    assert texts(r.data) == ["Pagina 1", "Pagina 2", "Pagina 1"]


def test_split_every_page_into_zip(client):
    r = post(client, "split-pdf", ("doc.pdf", make_pdf(3)))
    files = unzip(r.data)
    assert sorted(files) == ["pag_1.pdf", "pag_2.pdf", "pag_3.pdf"]
    assert texts(files["pag_2.pdf"]) == ["Pagina 2"]


@pytest.mark.parametrize("spec, expected", [
    ("2-3", ["Pagina 2", "Pagina 3"]),
    ("3, 1", ["Pagina 3", "Pagina 1"]),
    ("1-2,3", ["Pagina 1", "Pagina 2", "Pagina 3"]),
])
def test_split_page_list(client, spec, expected):
    r = post(client, "split-pdf", ("doc.pdf", make_pdf(3)), pages=spec)
    assert r.status_code == 200
    assert texts(r.data) == expected


@pytest.mark.parametrize("spec", ["0-1", "3-1", "1-9", "4", "-2", "abc", " , "])
def test_split_rejects_invalid_pages(client, spec):
    r = post(client, "split-pdf", ("doc.pdf", make_pdf(3)), pages=spec)
    assert r.status_code == 400


# --- Converter de PDF ---

def test_pdf_to_jpg_converts_every_page(client):
    r = post(client, "pdf-to-jpg", ("doc.pdf", make_pdf(3)), dpi="72")
    files = unzip(r.data)
    assert sorted(files) == ["doc_pag_1.jpg", "doc_pag_2.jpg", "doc_pag_3.jpg"]
    assert Image.open(io.BytesIO(files["doc_pag_1.jpg"])).size == (595, 842)


def test_pdf_to_jpg_single_page_returns_jpg(client):
    r = post(client, "pdf-to-jpg", ("doc.pdf", make_pdf(1)))
    assert Image.open(io.BytesIO(r.data)).format == "JPEG"


def test_pdf_to_word(client):
    r = post(client, "pdf-to-word", ("doc.pdf", make_pdf(2)))
    assert r.status_code == 200
    assert b"Pagina 2" in unzip(r.data)["word/document.xml"]


# --- Converter para PDF ---

def test_images_to_pdf_one_a4_page_per_image(client):
    r = post(client, "jpg-to-pdf", ("a.png", make_image()), ("b.jpg", make_image(fmt="JPEG")))
    doc = pymupdf.open(stream=r.data, filetype="pdf")
    assert doc.page_count == 2
    assert [round(x) for x in doc[0].rect] == [0, 0, 842, 595]  # imagem deitada, A4 deitada


def test_images_to_pdf_transparency_becomes_white(client):
    r = post(client, "jpg-to-pdf", ("t.png", make_image("RGBA", color=(0, 0, 0, 0))))
    page = pymupdf.open(stream=r.data, filetype="pdf")[0]
    assert page.get_pixmap(dpi=10).pixel(10, 10) == (255, 255, 255)


def test_images_to_pdf_follows_exif_orientation(client):
    img = Image.new("RGB", (400, 300))
    exif = img.getexif()
    exif[0x0112] = 6  # foto de celular: gravada deitada, deve ser exibida em pé
    buf = io.BytesIO()
    img.save(buf, "JPEG", exif=exif)
    r = post(client, "jpg-to-pdf", ("foto.jpg", buf.getvalue()))
    page = pymupdf.open(stream=r.data, filetype="pdf")[0]
    assert page.rect.height > page.rect.width


needs_libreoffice = pytest.mark.skipif(not euamopdf.find_soffice(), reason="LibreOffice não instalado")


@needs_libreoffice
def test_word_to_pdf(client):
    docx = pytest.importorskip("docx")
    document = docx.Document()
    document.add_paragraph("Olá, documento")
    buf = io.BytesIO()
    document.save(buf)
    r = post(client, "word-to-pdf", ("carta.docx", buf.getvalue()))
    assert r.status_code == 200
    assert "Olá, documento" in texts(r.data)[0]


@needs_libreoffice
def test_corrupted_office_file_is_reported(client):
    r = post(client, "word-to-pdf", ("quebrado.docx", b"PK\x03\x04lixo"))
    assert r.status_code == 400


# --- Validação ---

@pytest.mark.parametrize("action, files, message", [
    ("nao-existe", [("a.pdf", make_pdf())], "desconhecida"),
    ("merge-pdf", [], "Escolha um arquivo"),
    ("split-pdf", [("a.jpg", make_image())], "não é do tipo aceito"),
    ("split-pdf", [("a.pdf", make_pdf()), ("b.pdf", make_pdf())], "um arquivo por vez"),
    ("split-pdf", [("a.pdf", b"")], "vazio"),
    ("split-pdf", [("a.pdf", b"isto nao e um pdf")], "não é um PDF válido"),
    ("jpg-to-pdf", [("a.png", b"isto nao e uma imagem")], "não é uma imagem válida"),
    ("split-pdf", [("a.pdf", make_pdf(encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="1", owner_pw="1"))],
     "protegido por senha"),
])
def test_rejects_bad_input_with_message(client, action, files, message):
    r = post(client, action, *files)
    assert r.status_code == 400
    assert message in r.get_data(as_text=True)


def test_rejects_upload_over_limit(client, monkeypatch):
    monkeypatch.setitem(euamopdf.app.config, "MAX_CONTENT_LENGTH", 100)
    assert post(client, "merge-pdf", ("a.pdf", make_pdf())).status_code == 413


def test_parse_pages():
    assert euamopdf.parse_pages("1-3, 5", 5) == [0, 1, 2, 4]
    with pytest.raises(euamopdf.UserError):
        euamopdf.parse_pages("6", 5)


# --- Segurança e isolamento ---

def test_upload_names_cannot_escape_temp_folder(client, tmp_path):
    outside = tmp_path / "fora.pdf"
    relative = Path(tempfile.gettempdir(), "euamopdf_escapou.pdf")
    for name in (str(outside), "../euamopdf_escapou.pdf", "..\\euamopdf_escapou.pdf"):
        assert post(client, "split-pdf", (name, make_pdf())).status_code == 200
    assert not outside.exists()
    assert not relative.exists()


@pytest.mark.parametrize("headers, base_url", [
    ({"Origin": "https://site-malicioso.com"}, "http://127.0.0.1:5000"),
    ({"Origin": "null"}, "http://127.0.0.1:5000"),
    ({}, "http://site-malicioso.com"),  # DNS rebinding
])
def test_rejects_requests_from_other_sites(client, headers, base_url):
    r = client.post("/convert", base_url=base_url, headers=headers, content_type="multipart/form-data",
                    data={"action": "merge-pdf", "file": (io.BytesIO(make_pdf()), "a.pdf")})
    assert r.status_code == 403


def test_accepts_requests_from_the_app_itself(client):
    r = post(client, "merge-pdf", ("a.pdf", make_pdf()), headers={"Origin": "http://127.0.0.1:5000"})
    assert r.status_code == 200


def test_leaves_no_files_behind(client, tmp_path, monkeypatch):
    monkeypatch.setattr(tempfile, "tempdir", str(tmp_path))
    post(client, "split-pdf", ("a.pdf", make_pdf()))
    post(client, "pdf-to-jpg", ("a.pdf", make_pdf()))
    post(client, "split-pdf", ("a.pdf", b"invalido"))
    assert list(tmp_path.iterdir()) == []


def test_same_file_name_never_returns_previous_result(client):
    first = post(client, "split-pdf", ("mesmo.pdf", make_pdf(3)))
    second = post(client, "split-pdf", ("mesmo.pdf", make_pdf(2)))
    assert len(unzip(first.data)) == 3
    assert len(unzip(second.data)) == 2


def test_free_port_skips_busy_port():
    import socket
    with socket.socket() as busy:
        busy.bind(("127.0.0.1", 0))
        busy.listen()
        port = busy.getsockname()[1]
        assert euamopdf.free_port(port) != port
