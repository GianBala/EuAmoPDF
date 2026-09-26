import io
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import unquote

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


def make_photo_pdf():
    """PDF com uma foto de alta resolução numa área pequena da página."""
    buf = io.BytesIO()
    Image.effect_noise((1200, 1200), 60).convert("RGB").save(buf, "PNG")
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
