import hashlib
import io
import tempfile
import types
import zipfile
from pathlib import Path
from urllib.parse import unquote

import numpy as np
import pymupdf
import pytest
from PIL import Image, ImageDraw, ImageFilter

import app as euamopdf


@pytest.fixture
def client():
    return euamopdf.app.test_client()


def noise(size, sigma, seed=0):
    """Ruído gaussiano como o do Image.effect_noise, mas sempre o mesmo: teste reproduzível."""
    values = np.random.default_rng(seed).normal(128, sigma, (size[1], size[0]))
    return Image.fromarray(values.clip(0, 255).astype(np.uint8))


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


def make_photo_pdf():
    """PDF com uma foto de alta resolução numa área pequena da página."""
    buf = io.BytesIO()
    noise((1200, 1200), 60).convert("RGB").save(buf, "PNG")
    doc = pymupdf.open()
    doc.new_page().insert_image(pymupdf.Rect(72, 72, 272, 272), stream=buf.getvalue())
    return doc.tobytes()


@pytest.mark.parametrize("level", ["recomendada", "forte"])
def test_compress_reduces_images(client, level):
    original = make_photo_pdf()
    r = post(client, "compress-pdf", ("foto.pdf", original), level=level)
    assert r.status_code == 200
    assert len(r.data) < len(original) / 2
    assert "Reduzido de" in unquote(r.headers["X-Mensagem"])


def make_scan_pdf(dpi):
    """Página A4 escaneada: uma foto ocupando a página inteira na resolução dada."""
    size = (round(595 / 72 * dpi), round(842 / 72 * dpi))
    doc = pymupdf.open()
    doc.new_page().insert_image(pymupdf.paper_rect("a4"), stream=make_photo(size, quality=90))
    return doc.tobytes()


@pytest.mark.parametrize("level, width", [("recomendada", 1240), ("forte", 794)])
def test_compress_pdf_reaches_the_target_resolution(client, level, width):
    # 200 DPI -> 150 ou 96 DPI: fatores que não são potência de 2, que o MuPDF não reduzia
    r = post(client, "compress-pdf", ("scan.pdf", make_scan_pdf(200)), level=level)
    [image] = pymupdf.open(stream=r.data, filetype="pdf")[0].get_image_info()
    assert abs(image["width"] - width) <= 2
    assert image["bbox"][2] == pytest.approx(595, abs=1)  # continua ocupando a página


def test_compress_pdf_keeps_transparent_and_small_lossless_images(client):
    doc = pymupdf.open()
    page = doc.new_page()
    page.insert_image(pymupdf.Rect(0, 0, 200, 200), stream=make_photo((1000, 1000), fmt="PNG", mode="RGBA"))
    page.insert_image(pymupdf.Rect(0, 300, 300, 600), stream=make_photo((300, 300), fmt="PNG"))  # 72 DPI
    r = post(client, "compress-pdf", ("doc.pdf", doc.tobytes()))
    page = pymupdf.open(stream=r.data, filetype="pdf")[0]
    assert sorted(info["width"] for info in page.get_image_info()) == [300, 1000]
    assert [item[1] > 0 for item in page.get_images(full=True)].count(True) == 1  # a transparência continua


def test_compress_never_makes_file_bigger(client):
    original = make_pdf(1, garbage=4, deflate=True, clean=True, use_objstms=1)
    r = post(client, "compress-pdf", ("texto.pdf", original))
    assert len(r.data) <= len(original)
    assert "já está otimizado" in unquote(r.headers["X-Mensagem"])


def rotations(pdf):
    return [page.rotation for page in pymupdf.open(stream=pdf, filetype="pdf")]


def test_rotate_all_pages(client):
    r = post(client, "rotate-pdf", ("doc.pdf", make_pdf(3)), angle="90")
    assert rotations(r.data) == [90, 90, 90]


def test_rotate_selected_pages_adds_to_current_rotation(client):
    once = post(client, "rotate-pdf", ("doc.pdf", make_pdf(3)), angle="270", pages="1, 3").data
    twice = post(client, "rotate-pdf", ("doc.pdf", once), angle="180", pages="1").data
    assert rotations(twice) == [90, 0, 270]


def test_rotate_requires_valid_angle(client):
    assert post(client, "rotate-pdf", ("doc.pdf", make_pdf()), angle="45").status_code == 400


def test_protect_then_unlock(client):
    r = post(client, "protect-pdf", ("doc.pdf", make_pdf(2)), password="ação123", password2="ação123")
    locked = pymupdf.open(stream=r.data, filetype="pdf")
    assert locked.needs_pass
    assert locked.authenticate("ação123")

    assert post(client, "unlock-pdf", ("doc.pdf", r.data), password="errada").get_data(as_text=True) == "Senha incorreta."
    r = post(client, "unlock-pdf", ("doc.pdf", r.data), password="ação123")
    assert not pymupdf.open(stream=r.data, filetype="pdf").needs_pass
    assert texts(r.data) == ["Pagina 1", "Pagina 2"]


def test_protect_requires_matching_passwords(client):
    assert post(client, "protect-pdf", ("doc.pdf", make_pdf()), password="", password2="").status_code == 400
    assert post(client, "protect-pdf", ("doc.pdf", make_pdf()), password="a", password2="b").status_code == 400


def test_unlock_removes_restrictions_without_password(client):
    restricted = make_pdf(encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw="dono",
                          permissions=pymupdf.PDF_PERM_ACCESSIBILITY)
    r = post(client, "unlock-pdf", ("doc.pdf", restricted))
    assert r.status_code == 200
    assert pymupdf.open(stream=r.data, filetype="pdf").metadata["encryption"] is None


def test_unlock_plain_pdf_is_reported(client):
    r = post(client, "unlock-pdf", ("doc.pdf", make_pdf()))
    assert r.status_code == 400


def visual_lines(page):
    """Linhas de texto com direção e posição como o usuário vê a página (já girada)."""
    m = page.rotation_matrix
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            direction = pymupdf.Point(line["dir"]) * m - pymupdf.Point(0, 0) * m
            text = "".join(span["text"] for span in line["spans"])
            yield text, (round(direction.x, 2), round(direction.y, 2)), pymupdf.Rect(line["bbox"]) * m


@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_watermark_rises_diagonally_on_any_rotation(client, rotation):
    doc = pymupdf.open()
    doc.new_page().set_rotation(rotation)
    r = post(client, "watermark-pdf", ("doc.pdf", doc.tobytes()), text="CONFIDENCIAL")
    page = pymupdf.open(stream=r.data, filetype="pdf")[0]
    [(text, direction, bbox)] = visual_lines(page)
    assert text == "CONFIDENCIAL"
    assert direction == (0.71, -0.71)  # sobe da esquerda para a direita
    assert abs(bbox.x0 + bbox.x1 - page.rect.width) < 20  # centralizada


@pytest.mark.parametrize("text", ["", "   ", "Aprovado ✅"])
def test_watermark_rejects_invalid_text(client, text):
    assert post(client, "watermark-pdf", ("doc.pdf", make_pdf()), text=text).status_code == 400


@pytest.mark.parametrize("fmt, expected", [("n", "2"), ("n-de-t", "2 / 3"), ("pagina", "Página 2 de 3")])
def test_number_pages_formats(client, fmt, expected):
    r = post(client, "number-pages", ("doc.pdf", make_pdf(3)), format=fmt)
    page = pymupdf.open(stream=r.data, filetype="pdf")[1]
    assert expected in [text for text, _, _ in visual_lines(page)]


@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_number_pages_at_bottom_center_on_any_rotation(client, rotation):
    doc = pymupdf.open()
    doc.new_page().set_rotation(rotation)
    r = post(client, "number-pages", ("doc.pdf", doc.tobytes()))
    page = pymupdf.open(stream=r.data, filetype="pdf")[0]
    [(text, direction, bbox)] = visual_lines(page)
    assert (text, direction) == ("1", (1.0, 0.0))  # horizontal, como o usuário lê
    assert bbox.y1 > page.rect.height - 30
    assert abs(bbox.x0 + bbox.x1 - page.rect.width) < 5


def make_scanned_pdf(text, rotation=0):
    """PDF só com a imagem do texto, como um escaneamento. Com rotação, a imagem é gravada
    girada e a página a exibe em pé, como fazem muitos scanners."""
    source = pymupdf.open()
    source.new_page().insert_text((72, 100), text, fontsize=18)
    pix = source[0].get_pixmap(dpi=150)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples).rotate(rotation, expand=True)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    doc = pymupdf.open()
    page = doc.new_page(width=img.width * 72 / 150, height=img.height * 72 / 150)
    page.insert_image(page.rect, stream=buf.getvalue())
    page.set_rotation(rotation)
    return doc.tobytes()


needs_tesseract = pytest.mark.skipif(
    not (euamopdf.find_tessdata() and Path(euamopdf.find_tessdata(), "por.traineddata").exists()),
    reason="Tesseract com português não instalado")


@needs_tesseract
@pytest.mark.parametrize("rotation", [0, 90, 180, 270])
def test_ocr_makes_scanned_text_searchable(client, rotation):
    scanned = make_scanned_pdf("Contrato de locação número 42", rotation)
    assert texts(scanned) == [""]
    r = post(client, "ocr-pdf", ("scan.pdf", scanned))
    assert r.status_code == 200
    page = pymupdf.open(stream=r.data, filetype="pdf")[0]
    assert " ".join(page.get_text().split()) == "Contrato de locação número 42"
    assert page.search_for("locação")  # dá para buscar


def test_invisible_text_ignores_degenerate_ocr_boxes():
    page = pymupdf.open().new_page()
    euamopdf.add_invisible_text(page, [(10, 10, 10, 20, "vazia", 0, 0, 0), (10, 20, 50, 20, "achatada", 0, 1, 0),
                                       (10, 30, 60, 45, "certa", 0, 2, 0)])
    assert page.get_text().split() == ["certa"]


@needs_tesseract
def test_ocr_skips_pages_that_already_have_text(client):
    r = post(client, "ocr-pdf", ("doc.pdf", make_pdf()))
    assert r.status_code == 400
    assert "já têm texto" in r.get_data(as_text=True)


def test_ocr_without_tesseract_explains_what_to_install(client, monkeypatch):
    monkeypatch.setattr(euamopdf, "find_tessdata", lambda: None)
    r = post(client, "ocr-pdf", ("scan.pdf", make_pdf()))
    assert r.status_code == 400
    assert "Tesseract" in r.get_data(as_text=True)


def test_organize_reorders_rotates_duplicates_and_deletes(client):
    order = '[{"pagina": 3, "giro": 90}, {"pagina": 1}, {"pagina": 1, "giro": -90}]'
    r = post(client, "organize-pdf", ("doc.pdf", make_pdf(3)), order=order)
    assert texts(r.data) == ["Pagina 3", "Pagina 1", "Pagina 1"]
    assert rotations(r.data) == [90, 0, 270]


@pytest.mark.parametrize("order", ["", "[]", "nao e json", "[1, 2]", '[{"pagina": 4}]', '[{"pagina": 0}]',
                                   '[{"pagina": 1, "giro": 45}]', '{"pagina": 1}', '[{"page": 1}]'])
def test_organize_rejects_invalid_order(client, order):
    assert post(client, "organize-pdf", ("doc.pdf", make_pdf(3)), order=order).status_code == 400


def test_page_thumbnails(client):
    r = client.post("/pages", data={"file": (io.BytesIO(make_pdf(2)), "a.pdf"), "miniaturas": "1"},
                    content_type="multipart/form-data")
    thumbs = r.get_json()["miniaturas"]
    assert len(thumbs) == 2
    assert all(t.startswith("data:image/jpeg;base64,") for t in thumbs)


def make_photo(size=(1600, 1200), fmt="JPEG", quality=95, exif=None, mode="RGB"):
    """Imagem com gradientes e ruído, que se comprime como uma foto de verdade."""
    w, h = size
    img = Image.merge("RGB", [Image.linear_gradient("L").resize(size), Image.radial_gradient("L").resize(size),
                              noise(size, 30)])
    if mode == "RGBA":
        img.putalpha(Image.linear_gradient("L").resize(size))  # topo transparente, base opaca
    buf = io.BytesIO()
    options = {"quality": quality} if fmt in ("JPEG", "WEBP") else {}
    img.save(buf, fmt, **options, **({"exif": exif} if exif else {}))
    return buf.getvalue()


def opened(data):
    return Image.open(io.BytesIO(data))


def compress_image(client, *files, **form):
    return post(client, "compress-image", *files, **form)


def test_compress_image_levels_get_progressively_smaller(client):
    photo = make_photo()
    sizes = [len(compress_image(client, ("foto.jpg", photo), image_level=level).data)
             for level in ("leve", "recomendada", "forte", "extrema")]
    assert sizes == sorted(sizes, reverse=True)
    assert sizes[0] < len(photo)
    r = compress_image(client, ("foto.jpg", photo), image_level="recomendada")
    assert r.headers["Content-Disposition"].endswith("foto_comprimida.jpg")
    assert opened(r.data).format == "JPEG"
    assert "Reduzida de" in unquote(r.headers["X-Mensagem"])


def test_compress_image_resizes_only_down(client):
    r = compress_image(client, ("foto.jpg", make_photo((3000, 2000))), max_size="1280")
    assert opened(r.data).size == (1280, 853)
    r = compress_image(client, ("pequena.jpg", make_photo((600, 400))), max_size="1280")
    assert opened(r.data).size == (600, 400)


@pytest.mark.parametrize("target, fmt, ext", [("jpeg", "JPEG", ".jpg"), ("webp", "WEBP", ".webp"), ("png", "PNG", ".png")])
def test_compress_image_converts_format(client, target, fmt, ext):
    r = compress_image(client, ("foto.png", make_photo((600, 400), fmt="PNG")), image_format=target)
    assert opened(r.data).format == fmt
    assert r.headers["Content-Disposition"].endswith(f"foto_comprimida{ext}")


def test_compress_image_png_to_jpg_puts_transparency_on_white(client):
    png = make_photo((200, 200), fmt="PNG", mode="RGBA")
    r = compress_image(client, ("logo.png", png), image_format="jpeg")
    assert opened(r.data).getpixel((100, 0))[0] > 240  # topo transparente vira branco


def test_compress_image_png_keeps_transparency_and_reduces_colors(client):
    png = make_photo((400, 300), fmt="PNG", mode="RGBA")
    r = compress_image(client, ("logo.png", png), image_level="forte")
    img = opened(r.data)
    assert img.format == "PNG" and img.mode == "P"
    assert len(img.getcolors(256)) <= 128
    assert "transparency" in img.info or img.convert("RGBA").getextrema()[3][0] < 255
    assert len(r.data) < len(png)


def exif_with_location():
    exif = Image.Exif()
    exif[0x010F] = "Camera do celular"   # fabricante
    exif[0x0112] = 6                      # orientação: gravada deitada
    exif[0x8825] = {2: (23.0, 32.0, 0.0)}   # GPS: latitude (via get_ifd, o Pillow < 11.1 não grava)
    return exif


def test_compress_image_removes_photo_data_by_default_in_the_ui(client):
    photo = make_photo((400, 300), exif=exif_with_location())
    r = compress_image(client, ("foto.jpg", photo), strip_metadata="1")
    img = opened(r.data)
    assert not img.getexif()
    assert img.size == (300, 400)  # a orientação foi aplicada antes de tirar os dados


def test_compress_image_can_keep_photo_data(client):
    photo = make_photo((400, 300), exif=exif_with_location())
    r = compress_image(client, ("foto.jpg", photo))
    exif = opened(r.data).getexif()
    assert exif[0x010F] == "Camera do celular"
    assert 0x0112 not in exif  # orientação já aplicada: não pode ser girada de novo
    assert exif.get_ifd(0x8825)


def test_compress_image_keeps_already_compressed_original(client):
    small = make_photo((400, 300), quality=30)
    r = compress_image(client, ("leve.jpg", small), image_level="leve")
    assert r.data == small
    assert "já estava bem comprimida" in unquote(r.headers["X-Mensagem"])


def test_compress_many_images_into_zip_without_name_clashes(client):
    r = compress_image(client, ("foto.jpg", make_photo((400, 300))), ("foto.jpg", make_photo((300, 200))),
                       ("tela.png", make_photo((300, 200), fmt="PNG")))
    assert r.headers["Content-Disposition"].endswith("Imagens_comprimidas.zip")
    assert sorted(unzip(r.data)) == ["foto_comprimida.jpg", "foto_comprimida_2.jpg", "tela_comprimida.png"]
    assert unquote(r.headers["X-Mensagem"]).startswith("3 imagens: de")


@pytest.mark.parametrize("name, content", [("anim.gif", make_image(fmt="GIF")), ("foto.jpg", b"nao e imagem")])
def test_compress_image_rejects_other_files(client, name, content):
    assert compress_image(client, (name, content)).status_code == 400


def estimate(client, action, *files, **form):
    r = client.post("/estimate", data={"action": action, "file": [(io.BytesIO(c), n) for n, c in files], **form},
                    content_type="multipart/form-data")
    return r


@pytest.mark.parametrize("form", [
    {"image_level": "recomendada"},
    {"image_level": "extrema", "image_format": "webp"},
    {"image_level": "forte", "image_format": "png", "max_size": "1280"},
    {"image_level": "leve", "max_size": "800"},
])
def test_image_estimate_is_close_to_the_real_result(client, form):
    photos = [("a.jpg", make_photo((3000, 2000))), ("b.jpg", make_photo((1600, 1200), quality=80))]  # a primeira passa pelo mosaico
    r = estimate(client, "compress-image", *photos, **form)
    assert r.status_code == 200
    body = r.get_json()
    assert body["antes"] == sum(len(content) for _, content in photos)
    real = sum(len(data) for data in unzip(compress_image(client, *photos, **form).data).values())
    assert body["depois"] == pytest.approx(real, rel=0.2)


def test_image_estimate_samples_at_least_three_images(client, monkeypatch):
    monkeypatch.setattr(euamopdf, "ESTIMATE_SECONDS", 0)  # prazo esgotado já na primeira imagem
    files = [("foto.jpg", make_photo((1600, 1200))), ("logo.png", make_photo((600, 600), fmt="PNG", mode="RGBA"))]
    form = {"image_level": "forte", "image_format": "webp"}
    body = estimate(client, "compress-image", *files, **form).get_json()
    real = sum(len(data) for data in unzip(compress_image(client, *files, **form).data).values())
    assert body["depois"] == pytest.approx(real, rel=0.15)


def test_image_estimate_never_exceeds_an_original_that_stays(client):
    small = make_photo((400, 300), quality=30)
    body = estimate(client, "compress-image", ("leve.jpg", small), image_level="leve").get_json()
    assert body["depois"] == len(small)


@pytest.mark.parametrize("level", ["recomendada", "forte"])
def test_pdf_estimate_is_close_to_the_real_result(client, level):
    doc = pymupdf.open()
    for dpi in (200, 250, 150):  # três páginas escaneadas, cada uma numa resolução
        doc.insert_pdf(pymupdf.open(stream=make_scan_pdf(dpi), filetype="pdf"))
    pdf = doc.tobytes()
    body = estimate(client, "compress-pdf", ("scan.pdf", pdf), level=level).get_json()
    real = len(post(client, "compress-pdf", ("scan.pdf", pdf), level=level).data)
    assert body["antes"] == len(pdf)
    assert body["depois"] == pytest.approx(real, rel=0.1)


def test_pdf_estimate_with_uncompressed_image(client):
    # Imagem gravada sem compressão: ao salvar, o PDF a comprime, e a estimativa precisa contar com isso
    doc = pymupdf.open()
    page = doc.new_page()
    xref = page.insert_image(page.rect, stream=make_photo((1240, 1754), fmt="PNG"))
    doc.update_stream(xref, pymupdf.Pixmap(doc, xref).samples, compress=False)
    doc.xref_set_key(xref, "Filter", "null")
    doc.xref_set_key(xref, "DecodeParms", "null")
    pdf = doc.tobytes()
    body = estimate(client, "compress-pdf", ("scan.pdf", pdf), level="forte").get_json()
    real = len(post(client, "compress-pdf", ("scan.pdf", pdf), level="forte").data)
    assert 0 < body["depois"] < body["antes"]
    assert body["depois"] == pytest.approx(real, rel=0.2)


def test_pdf_estimate_never_exceeds_the_original(client):
    pdf = make_pdf(1, garbage=4, deflate=True, clean=True, use_objstms=1)
    body = estimate(client, "compress-pdf", ("texto.pdf", pdf)).get_json()
    assert body["depois"] <= body["antes"]


@pytest.mark.parametrize("action, files, status", [
    ("merge-pdf", [("a.pdf", make_pdf())], 400),                       # ferramenta sem estimativa
    ("compress-pdf", [("a.pdf", b"lixo")], 400),                       # arquivo inválido
    ("compress-pdf", [("a.pdf", make_pdf(encryption=pymupdf.PDF_ENCRYPT_AES_256, user_pw="1", owner_pw="1"))], 400),
    ("compress-image", [("a.gif", make_image(fmt="GIF"))], 400),       # tipo não aceito
])
def test_estimate_rejects_what_the_tool_would_reject(client, action, files, status):
    assert estimate(client, action, *files).status_code == status


def test_spread_picks_items_across_the_list():
    assert euamopdf.spread(list(range(10)), 3) == [0, 3, 6]
    assert euamopdf.spread([1, 2], 3) == [1, 2]


# --- Remover fundo ---

class FakeModel:
    """Faz as vezes do ISNet: a máscara é um disco no centro, qualquer que seja a imagem."""
    def __init__(self, found=True):
        yy, xx = np.mgrid[:1024, :1024]
        self.mask = ((yy - 512) ** 2 + (xx - 512) ** 2 < 300 ** 2).astype(np.float32) * found

    def get_inputs(self):
        return [types.SimpleNamespace(name="entrada")]

    def run(self, outputs, feeds):
        x = feeds["entrada"]
        assert x.shape == (1, 3, 1024, 1024) and x.dtype == np.float32 and -0.51 < x.min() and x.max() < 0.51
        return [self.mask[None, None]]


@pytest.fixture
def fake_model(monkeypatch):
    monkeypatch.setattr(euamopdf, "_bg_session", FakeModel())


def remove_bg(client, *files, **form):
    return post(client, "remove-background", *files, **form)


def test_remove_background_keeps_original_pixels_and_size(client, fake_model):
    photo = make_photo((1600, 1200), fmt="PNG")
    r = remove_bg(client, ("foto.png", photo))
    assert r.headers["Content-Disposition"].endswith("foto_sem_fundo.png")
    out, src = opened(r.data), opened(photo).convert("RGB")
    assert out.format == "PNG" and out.mode == "RGBA" and out.size == src.size
    alpha = np.asarray(out)[..., 3]
    assert alpha[600, 800] == 255 and alpha[0, 0] == 0
    inside = alpha == 255  # no objeto, os pixels são exatamente os da fonte
    assert np.array_equal(np.asarray(out)[..., :3][inside], np.asarray(src)[inside])


def test_remove_background_erases_what_was_behind_transparent_pixels(client, fake_model):
    r = remove_bg(client, ("foto.png", make_photo((400, 300), fmt="PNG")))
    pixels = np.asarray(opened(r.data))
    assert not pixels[pixels[..., 3] == 0].any()  # o fundo não fica escondido atrás da transparência


def test_remove_background_follows_photo_orientation_and_keeps_color_profile(client, fake_model):
    img = Image.open(io.BytesIO(make_photo((600, 400))))
    exif = Image.Exif()
    exif[0x0112] = 6
    buf = io.BytesIO()
    img.save(buf, "JPEG", exif=exif, icc_profile=b"perfil de cor de teste")
    out = opened(remove_bg(client, ("foto.jpg", buf.getvalue())).data)
    assert out.size == (400, 600)
    assert out.info.get("icc_profile") == b"perfil de cor de teste"


@pytest.mark.parametrize("form, color", [
    ({"background": "branco"}, (255, 255, 255)),
    ({"background": "preto"}, (0, 0, 0)),
    ({"background": "cor", "background_color": "#2F80ED"}, (47, 128, 237)),
])
def test_remove_background_paints_new_background(client, fake_model, form, color):
    photo = make_photo((400, 300), fmt="PNG")
    out = opened(remove_bg(client, ("foto.png", photo), **form).data)
    assert out.mode == "RGB"
    assert out.getpixel((0, 0)) == color
    assert out.getpixel((200, 150)) == opened(photo).convert("RGB").getpixel((200, 150))


def test_remove_background_webp_is_lossless(client, fake_model):
    photo = make_photo((400, 300), fmt="PNG")
    r = remove_bg(client, ("foto.png", photo), bg_format="webp")
    out = opened(r.data)
    assert out.format == "WEBP" and r.headers["Content-Disposition"].endswith(".webp")
    assert out.convert("RGB").getpixel((200, 150)) == opened(photo).convert("RGB").getpixel((200, 150))


def test_remove_background_keeps_existing_transparency(client, fake_model):
    png = make_photo((400, 400), fmt="PNG", mode="RGBA")  # transparência em degradê: topo transparente
    alpha = np.asarray(opened(remove_bg(client, ("logo.png", png)).data))[..., 3].astype(int)
    source = np.asarray(opened(png))[..., 3].astype(int)
    for y in (120, 200, 300):  # dentro do disco (raio ~117 px), vale a transparência original
        assert abs(alpha[y, 200] - source[y, 200]) <= 1
    assert alpha[200, 10] == 0  # fora do disco, transparente


def test_remove_background_many_images_and_no_subject(client, monkeypatch):
    monkeypatch.setattr(euamopdf, "_bg_session", FakeModel(found=False))
    r = remove_bg(client, ("a.jpg", make_photo((300, 200))), ("a.jpg", make_photo((200, 300))))
    assert sorted(unzip(r.data)) == ["a_sem_fundo.png", "a_sem_fundo_2.png"]
    assert "Nenhum objeto em destaque" in unquote(r.headers["X-Mensagem"])


def test_remove_background_rejects_invalid_color(client, fake_model):
    r = remove_bg(client, ("a.jpg", make_photo((300, 200))), background="cor", background_color="vermelho")
    assert r.status_code == 400


# --- Prévia e correção do remover fundo ---

def png_bytes(img):
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def scene():
    """Fundo azul com textura e dois retângulos amarelos; o 'modelo' só achou o da esquerda."""
    size = (400, 300)
    channel = lambda low, seed: noise(size, 12, seed).point(lambda v: low + v // 4)
    img = Image.merge("RGB", [channel(10, 1), channel(80, 2), channel(160, 3)])
    draw = ImageDraw.Draw(img)
    draw.rectangle((40, 60, 160, 240), fill=(240, 200, 30))
    draw.rectangle((240, 60, 360, 240), fill=(235, 195, 40))
    mask = Image.new("L", size, 0)
    ImageDraw.Draw(mask).rectangle((40, 60, 160, 240), fill=255)
    return img, mask


def dot(size, center, radius=8, scale=1):
    """Traço como a prévia manda: RGBA colorido e semitransparente, no tamanho da tela."""
    layer = Image.new("RGBA", (round(size[0] * scale), round(size[1] * scale)))
    x, y, r = center[0] * scale, center[1] * scale, radius * scale
    ImageDraw.Draw(layer).ellipse((x - r, y - r, x + r, y + r), fill=(40, 200, 80, 120))
    return layer


def refine(client, img, mask, strokes, mode):
    return client.post("/background/refine", content_type="multipart/form-data", data={
        "modo": mode, "file": (io.BytesIO(png_bytes(img)), "cena.png"),
        "mascara": (io.BytesIO(png_bytes(mask)), "mascara.png"),
        "traco": (io.BytesIO(png_bytes(strokes)), "traco.png")})


def region(mask, box):
    x0, y0, x1, y1 = box
    return np.asarray(mask)[y0:y1 + 1, x0:x1 + 1]


def test_background_mask_route_returns_the_model_mask(client, fake_model):
    r = client.post("/background/mask", data={"action": "remove-background", "file": (io.BytesIO(make_photo((600, 400))), "f.jpg")},
                    content_type="multipart/form-data")
    mask = opened(r.data)
    assert r.mimetype == "image/png" and mask.mode == "L" and mask.size == (600, 400)
    assert mask.getpixel((300, 200)) == 255 and mask.getpixel((0, 0)) == 0


def background_restored(new):
    """Fração do fundo azul (fora dos dois retângulos) que foi restaurada por engano."""
    outside = np.asarray(new).copy()
    outside[58:243, 38:163] = 0
    outside[58:243, 238:363] = 0
    return (outside > 127).mean()


def test_refine_restores_the_rest_of_a_part_the_model_cut(client):
    img, mask = scene()
    ImageDraw.Draw(mask).rectangle((115, 100, 160, 200), fill=0)  # o modelo cortou um pedaço do retângulo
    # um clique só, numa prévia com metade do tamanho da imagem
    new = opened(refine(client, img, mask, dot(img.size, (138, 150), radius=10, scale=0.5), "restaurar").data)
    assert (region(new, (117, 102, 158, 198)) > 127).mean() > 0.95
    assert background_restored(new) < 0.02
    assert (region(new, (242, 62, 358, 238)) == 0).all()  # o outro retângulo não foi tocado


def test_refine_erases_leftover_background_but_keeps_the_object(client):
    img, mask = scene()
    ImageDraw.Draw(mask).rectangle((170, 100, 225, 200), fill=255)  # pedaço de fundo que ficou
    new = opened(refine(client, img, mask, dot(img.size, (198, 150)), "apagar").data)
    assert (region(new, (172, 102, 223, 198)) < 128).mean() > 0.95
    assert (region(new, (42, 62, 158, 238)) == 255).all()


def test_refine_restores_a_whole_object_the_model_missed_painted_over(client):
    img, _ = scene()
    nothing = Image.new("L", img.size, 0)  # o modelo não achou nenhum dos dois
    stroke = Image.new("RGBA", img.size)
    ImageDraw.Draw(stroke).line([(70, 90), (130, 210)], fill=(40, 200, 80, 255), width=16)
    new = opened(refine(client, img, nothing, stroke, "restaurar").data)
    assert (region(new, (42, 62, 158, 238)) > 127).mean() > 0.95
    assert (region(new, (242, 62, 358, 238)) == 0).all()  # o outro retângulo, que não foi pintado, não muda
    assert background_restored(new) < 0.02


def test_refine_makes_a_half_transparent_part_solid(client):
    img, mask = scene()
    ImageDraw.Draw(mask).rectangle((240, 60, 360, 240), fill=150)  # o modelo ficou em dúvida no da direita
    stroke = Image.new("RGBA", img.size)
    ImageDraw.Draw(stroke).line([(260, 90), (340, 210)], fill=(40, 200, 80, 255), width=16)  # pintado por dentro dele
    new = opened(refine(client, img, mask, stroke, "restaurar").data)
    assert (region(new, (244, 64, 356, 236)) >= 250).mean() > 0.95


def test_refine_erases_a_half_transparent_leftover(client):
    img, mask = scene()
    ImageDraw.Draw(mask).rectangle((170, 100, 225, 200), fill=110)  # fundo que ficou meio visível
    new = opened(refine(client, img, mask, dot(img.size, (198, 150)), "apagar").data)
    assert (region(new, (174, 104, 221, 196)) <= 5).mean() > 0.95
    assert (region(new, (42, 62, 158, 238)) == 255).all()


def test_refine_does_not_follow_a_stroke_that_spills_over(client):
    img, mask = scene()
    ImageDraw.Draw(mask).rectangle((115, 100, 160, 200), fill=0)
    # traço grosso que começa no pedaço que faltou e passa um raio de pincel para o fundo azul
    stroke = Image.new("RGBA", img.size)
    ImageDraw.Draw(stroke).line([(130, 150), (170, 150)], fill=(40, 200, 80, 255), width=16)
    new = np.asarray(opened(refine(client, img, mask, stroke, "restaurar").data))
    assert (new[102:199, 117:159] > 127).mean() > 0.95
    assert (new[143:158, 168:186] > 127).mean() < 0.5  # o fundo por onde o traço vazou não volta


def test_refine_without_smart_detection_changes_exactly_what_was_painted(client):
    img, mask = scene()
    stroke = Image.new("RGBA", img.size)
    ImageDraw.Draw(stroke).rectangle((180, 100, 220, 140), fill=(40, 200, 80, 255))  # só fundo azul
    r = client.post("/background/refine", content_type="multipart/form-data", data={
        "modo": "restaurar", "inteligente": "0", "file": (io.BytesIO(png_bytes(img)), "cena.png"),
        "mascara": (io.BytesIO(png_bytes(mask)), "mascara.png"), "traco": (io.BytesIO(png_bytes(stroke)), "traco.png")})
    new = np.asarray(opened(r.data))
    assert (new[100:141, 180:221] == 255).all()        # o que foi pintado voltou, mesmo sendo fundo
    changed = new != np.asarray(mask)
    changed[100:141, 180:221] = False
    assert not changed.any()                           # e nada além disso mudou

@pytest.mark.parametrize("change, message", [
    ({"modo": "pintar"}, "restaurar e apagar"),
    ({"mascara": None}, "Falta a máscara"),
    ({"mascara": Image.new("L", (10, 10))}, "tamanho da imagem"),
    ({"traco": Image.new("RGBA", (400, 300))}, "Pinte por cima"),
])
def test_refine_rejects_bad_requests(client, change, message):
    img, mask = scene()
    data = {"modo": "restaurar", "file": (io.BytesIO(png_bytes(img)), "cena.png"),
            "mascara": mask, "traco": dot(img.size, (300, 150))}
    data.update(change)
    for field in ("mascara", "traco"):
        if data[field] is None:
            del data[field]
        elif isinstance(data[field], Image.Image):
            data[field] = (io.BytesIO(png_bytes(data[field])), f"{field}.png")
    r = client.post("/background/refine", data=data, content_type="multipart/form-data")
    assert r.status_code == 400
    assert message in r.get_data(as_text=True)


def test_download_uses_the_mask_edited_in_the_preview(client, fake_model):
    img, mask = scene()
    r = client.post("/convert", content_type="multipart/form-data", data={
        "action": "remove-background", "file": (io.BytesIO(png_bytes(img)), "cena.png"),
        "mascara": (io.BytesIO(png_bytes(mask)), "mascara.png")})
    alpha = np.asarray(opened(r.data))[..., 3]
    assert np.array_equal(alpha, np.asarray(mask))  # e não o disco do modelo falso


def test_download_needs_one_mask_per_image(client, fake_model):
    img, mask = scene()
    r = client.post("/convert", content_type="multipart/form-data", data={
        "action": "remove-background",
        "file": [(io.BytesIO(png_bytes(img)), "a.png"), (io.BytesIO(png_bytes(img)), "b.png")],
        "mascara": (io.BytesIO(png_bytes(mask)), "mascara.png")})
    assert r.status_code == 400


@pytest.fixture
def model_source(tmp_path, monkeypatch):
    """Um 'modelo' servido de um arquivo local, com o SHA-256 dele."""
    source = tmp_path / "modelo.onnx"
    source.write_bytes(b"modelo de teste" * 1000)
    monkeypatch.setattr(euamopdf, "BG_MODEL_URL", source.as_uri())
    monkeypatch.setattr(euamopdf, "BG_MODEL_SHA256", hashlib.sha256(source.read_bytes()).hexdigest())
    return source


def test_model_download_puts_the_file_in_place(tmp_path, model_source):
    target = tmp_path / "cache" / "modelo.onnx"
    euamopdf.download_bg_model(target)
    assert target.read_bytes() == model_source.read_bytes()
    assert list(target.parent.iterdir()) == [target]  # nada de .part sobrando


def test_model_download_rejects_a_corrupted_file(tmp_path, model_source, monkeypatch):
    monkeypatch.setattr(euamopdf, "BG_MODEL_SHA256", "0" * 64)
    target = tmp_path / "cache" / "modelo.onnx"
    with pytest.raises(euamopdf.UserError, match="corrompido"):
        euamopdf.download_bg_model(target)
    assert list(target.parent.iterdir()) == []


def test_model_download_without_internet(tmp_path, monkeypatch):
    monkeypatch.setattr(euamopdf, "BG_MODEL_URL", (tmp_path / "nao-existe.onnx").as_uri())
    with pytest.raises(euamopdf.UserError, match="internet"):
        euamopdf.download_bg_model(tmp_path / "cache" / "modelo.onnx")


def test_model_goes_to_the_system_cache(monkeypatch, tmp_path):
    monkeypatch.delenv("LOCALAPPDATA", raising=False)
    monkeypatch.setenv("XDG_CACHE_HOME", str(tmp_path))
    assert euamopdf.bg_model_path() == tmp_path / "euamopdf" / "isnet-general-use.onnx"


def test_first_use_warns_about_the_model_download(client, monkeypatch, tmp_path):
    monkeypatch.setattr(euamopdf, "bg_model_path", lambda: tmp_path / "nao-baixado.onnx")
    assert "baixa o modelo de IA" in client.get("/").get_data(as_text=True)
    monkeypatch.setattr(euamopdf, "bg_model_path", lambda: tmp_path)  # "já baixado"
    assert "baixa o modelo de IA" not in client.get("/").get_data(as_text=True)


@pytest.mark.skipif(not euamopdf.bg_model_path().exists(), reason="modelo de remover fundo não baixado")
def test_remove_background_with_the_real_model(client, monkeypatch):
    monkeypatch.setattr(euamopdf, "_bg_session", None)
    scene = Image.open(io.BytesIO(make_photo((1200, 800)))).filter(ImageFilter.GaussianBlur(25))
    subject = Image.new("L", scene.size, 0)
    ImageDraw.Draw(subject).ellipse((400, 200, 800, 600), fill=255)
    scene.paste(Image.new("RGB", scene.size, (230, 60, 40)), mask=subject)  # disco vermelho em destaque
    buf = io.BytesIO()
    scene.save(buf, "PNG")
    alpha = np.asarray(opened(remove_bg(client, ("cena.png", buf.getvalue())).data))[..., 3] > 127
    truth = np.asarray(subject) > 127
    assert (alpha & truth).sum() / (alpha | truth).sum() > 0.9


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


def make_table_pdf(rows):
    """PDF com uma tabela de bordas desenhadas, como as geradas por planilhas."""
    doc = pymupdf.open()
    page = doc.new_page()
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = pymupdf.Rect(72 + c * 120, 72 + r * 24, 192 + c * 120, 96 + r * 24)
            page.draw_rect(cell, color=(0, 0, 0), width=0.5)
            page.insert_text(cell.bl + (4, -7), value)
    return doc.tobytes()


def test_pdf_to_excel_extracts_tables(client):
    openpyxl = pytest.importorskip("openpyxl")
    rows = [["Produto", "Preço"], ["Café", "12,50"], ["Pão", "0,75"]]
    r = post(client, "pdf-to-excel", ("tabela.pdf", make_table_pdf(rows)))
    assert r.status_code == 200
    sheet = openpyxl.load_workbook(io.BytesIO(r.data)).worksheets[0]
    assert [[cell.value for cell in row] for row in sheet.iter_rows()] == rows


def test_pdf_to_excel_without_tables(client):
    r = post(client, "pdf-to-excel", ("texto.pdf", make_pdf()))
    assert r.status_code == 400
    assert "Nenhuma tabela" in r.get_data(as_text=True)


def test_pdf_to_ppt_one_slide_per_page(client):
    pptx = pytest.importorskip("pptx")
    doc = pymupdf.open()
    doc.new_page(width=842, height=595)  # A4 deitada
    doc.new_page(width=595, height=842)  # A4 em pé: deve caber sem distorcer
    r = post(client, "pdf-to-ppt", ("slides.pdf", doc.tobytes()))
    assert r.status_code == 200
    prs = pptx.Presentation(io.BytesIO(r.data))
    assert len(prs.slides) == 2
    assert prs.slide_width > prs.slide_height
    picture = prs.slides[1].shapes[0]
    assert picture.height == prs.slide_height
    assert abs(picture.width / picture.height - 595 / 842) < 0.01


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


def test_unexpected_error_gives_generic_message(client, monkeypatch):
    def broken(files, form, tmp):
        raise RuntimeError("/caminho/interno/que/nao/deve/vazar")
    monkeypatch.setitem(euamopdf.ACTIONS, "merge-pdf", (broken, (".pdf",), True))
    r = post(client, "merge-pdf", ("a.pdf", make_pdf()))
    assert r.status_code == 500
    assert "Não foi possível processar" in r.get_data(as_text=True)
    assert "caminho" not in r.get_data(as_text=True)


def test_unknown_route_is_still_404(client):
    assert client.get("/nao-existe").status_code == 404


def test_rejects_upload_over_limit(client, monkeypatch):
    monkeypatch.setitem(euamopdf.app.config, "MAX_CONTENT_LENGTH", 100)
    assert post(client, "merge-pdf", ("a.pdf", make_pdf())).status_code == 413


def test_parse_pages():
    assert euamopdf.parse_pages("1-3, 5", 5) == [0, 1, 2, 4]
    with pytest.raises(euamopdf.UserError):
        euamopdf.parse_pages("6", 5)


# --- Interface ---

def test_index_has_a_button_for_every_tool(client):
    page = client.get("/").get_data(as_text=True)
    for action in euamopdf.ACTIONS:
        assert f'data-action="{action}"' in page


def test_page_count(client):
    r = client.post("/pages", data={"file": (io.BytesIO(make_pdf(4)), "a.pdf")}, content_type="multipart/form-data")
    assert r.get_json() == {"paginas": 4}
    r = client.post("/pages", data={"file": (io.BytesIO(b"lixo"), "a.pdf")}, content_type="multipart/form-data")
    assert r.status_code == 400


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
