from flask import Flask, render_template, request, send_file
import io
import os
import re
import sys
import shutil
import socket
import subprocess
import tempfile
import webbrowser # Biblioteca para abrir o navegador
from pathlib import Path
from threading import Timer # Para atrasar a abertura em 1 segundo
from urllib.parse import urlparse

import zipfile
import pymupdf
from PIL import Image, ImageOps


app = Flask(__name__)

# --- CONFIGURAÇÕES ---
LOCAL_HOSTS = {'127.0.0.1', 'localhost'}
MAX_UPLOAD_MB = 500
LIBREOFFICE_TIMEOUT = 180  # segundos
app.config['MAX_CONTENT_LENGTH'] = MAX_UPLOAD_MB * 1024 * 1024


class UserError(Exception):
    """Erro causado pela entrada do usuário; a mensagem é mostrada como está."""


@app.before_request
def only_local_requests():
    """Recusa requisições de outros sites abertos no navegador (CSRF e DNS rebinding)."""
    origin = request.headers.get('Origin')
    if urlparse('//' + request.host).hostname not in LOCAL_HOSTS or \
            (origin and urlparse(origin).hostname not in LOCAL_HOSTS):
        return "Acesso permitido só a partir deste computador.", 403


def find_soffice():
    found = shutil.which('soffice') or shutil.which('libreoffice')
    if found or sys.platform != 'win32':
        return found
    # No Windows o instalador do LibreOffice não põe o soffice no PATH
    for base in (os.environ.get('PROGRAMFILES'), os.environ.get('PROGRAMFILES(X86)')):
        if base and (exe := Path(base, 'LibreOffice', 'program', 'soffice.exe')).exists():
            return str(exe)
    return None


def msoffice_to_pdf(src, out, app_name):
    """Converte com o Microsoft Office numa instância própria, sem tocar no Office que o usuário tem aberto."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    office = None
    try:
        office = win32com.client.DispatchEx(app_name)
        if app_name == "Word.Application":
            office.DisplayAlerts = 0  # wdAlertsNone
            doc = office.Documents.Open(str(src), ReadOnly=True)
            try: doc.SaveAs(str(out), FileFormat=17)  # wdFormatPDF
            finally: doc.Close(False)
        elif app_name == "Excel.Application":
            office.DisplayAlerts = False
            wb = office.Workbooks.Open(str(src), ReadOnly=True)
            try: wb.ExportAsFixedFormat(0, str(out))  # xlTypePDF
            finally: wb.Close(False)
        else:
            pres = office.Presentations.Open(str(src), ReadOnly=True, WithWindow=False)
            try: pres.SaveAs(str(out), 32)  # ppSaveAsPDF
            finally: pres.Close()
    finally:
        if office is not None:
            office.Quit()
        pythoncom.CoUninitialize()


def libreoffice_to_pdf(src, out_dir):
    soffice = find_soffice()
    if not soffice:
        raise UserError("Para converter arquivos do Office é preciso ter o Microsoft Office (Windows) ou o LibreOffice instalado.")
    # Perfil próprio por conversão: não conflita com outro LibreOffice aberto nem com conversões simultâneas
    profile = (out_dir / 'perfil').as_uri()
    try:
        subprocess.run([soffice, '--headless', '--norestore', f'-env:UserInstallation={profile}',
                        '--convert-to', 'pdf', '--outdir', str(out_dir), str(src)],
                       capture_output=True, timeout=LIBREOFFICE_TIMEOUT)
    except subprocess.TimeoutExpired:
        raise UserError("O LibreOffice demorou demais para converter este arquivo.")
    # O LibreOffice sai com código 0 mesmo quando falha: o que vale é o PDF existir
    out = out_dir / f"{src.stem}.pdf"
    if not out.exists():
        raise UserError("O LibreOffice não conseguiu converter este arquivo. Veja se ele abre normalmente.")
    return out


def office_to_pdf(src, tmp, app_name):
    if sys.platform == 'win32':
        out = tmp / 'saida.pdf'
        try:
            msoffice_to_pdf(src, out, app_name)
            return out
        except Exception:
            app.logger.info("Microsoft Office indisponível; tentando o LibreOffice", exc_info=True)
    out_dir = tmp / 'libreoffice'
    out_dir.mkdir()
    return libreoffice_to_pdf(src, out_dir)


def open_pdf(path):
    try:
        doc = pymupdf.open(path, filetype='pdf')
    except Exception:
        raise UserError("O arquivo não é um PDF válido ou está corrompido.")
    if doc.needs_pass:
        raise UserError("Este PDF está protegido por senha.")
    return doc


def pdf_bytes(doc):
    return doc.tobytes(garbage=3, deflate=True)


def parse_pages(spec, count):
    """Converte '1-3, 5' em índices [0, 1, 2, 4], na ordem em que foram escritos."""
    pages = []
    for part in spec.replace(' ', '').split(','):
        if not part:
            continue
        m = re.fullmatch(r'(\d+)(?:-(\d+))?', part)
        if not m:
            raise UserError(f"Não entendi \"{part}\". Escreva as páginas assim: 1-3, 5, 8.")
        start, end = int(m[1]), int(m[2] or m[1])
        if start > end:
            raise UserError(f"Intervalo invertido: \"{part}\". Use o menor número primeiro.")
        if start < 1 or end > count:
            raise UserError(f"Página fora do documento: \"{part}\". O PDF tem {count} página(s).")
        pages.extend(range(start - 1, end))
    if not pages:
        raise UserError("Informe ao menos uma página.")
    return pages


# --- AÇÕES ---
# Cada ação recebe a lista de (caminho salvo, nome original sem extensão),
# o formulário e a pasta temporária da requisição, e devolve (bytes, nome do download).

def merge_pdf(files, form, tmp):
    out = pymupdf.open()
    for path, _ in files:
        out.insert_pdf(open_pdf(path))
    return pdf_bytes(out), "PDF_Unido.pdf"


def split_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    spec = form.get('pages', '').strip()

    if spec:
        doc.select(parse_pages(spec, doc.page_count))
        return pdf_bytes(doc), f"{base}_paginas.pdf"

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for i in range(doc.page_count):
            page = pymupdf.open()
            page.insert_pdf(doc, from_page=i, to_page=i)
            z.writestr(f"pag_{i+1}.pdf", pdf_bytes(page))
    return buf.getvalue(), f"{base}_dividido.zip"


def office_action(app_name):
    def convert(files, form, tmp):
        path, base = files[0]
        return office_to_pdf(path, tmp, app_name).read_bytes(), f"{base}.pdf"
    return convert


A4 = pymupdf.paper_rect('a4')


def load_image(path):
    """Abre a imagem já na orientação da foto e com a transparência sobre fundo branco."""
    try:
        img = Image.open(path)
        fmt = img.format
        img = ImageOps.exif_transpose(img)
    except Exception:
        raise UserError("O arquivo não é uma imagem válida.")
    if img.mode in ('RGBA', 'LA', 'PA') or 'transparency' in img.info:
        img = img.convert('RGBA')
        bg = Image.new('RGB', img.size, 'white')
        bg.paste(img, mask=img.getchannel('A'))
        img = bg
    buf = io.BytesIO()
    if fmt == 'JPEG':  # foto continua JPEG; o resto vira PNG para não borrar texto
        img.convert('RGB').save(buf, 'JPEG', quality=95)
    else:
        img.convert('RGB').save(buf, 'PNG')
    return buf.getvalue(), img.size


def images_to_pdf(files, form, tmp):
    out = pymupdf.open()
    for path, _ in files:
        data, (w, h) = load_image(path)
        # Página A4 em pé ou deitada, conforme a imagem, com a imagem inteira centralizada
        size = (A4.width, A4.height) if h >= w else (A4.height, A4.width)
        page = out.new_page(width=size[0], height=size[1])
        page.insert_image(page.rect, stream=data)
    name = f"{files[0][1]}.pdf" if len(files) == 1 else "Imagens.pdf"
    return pdf_bytes(out), name


def pdf_to_word(files, form, tmp):
    path, base = files[0]
    open_pdf(path)
    from pdf2docx import Converter  # importação lenta: só quando usada
    out = tmp / 'saida.docx'
    cv = Converter(str(path)); cv.convert(str(out)); cv.close()
    return out.read_bytes(), f"{base}.docx"


def pdf_to_jpg(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    dpi = {'72': 72, '300': 300}.get(form.get('dpi'), 150)

    def jpg(page):
        return page.get_pixmap(dpi=dpi).tobytes('jpg', jpg_quality=90)

    if doc.page_count == 1:
        return jpg(doc[0]), f"{base}.jpg"
    digits = len(str(doc.page_count))  # pag_01, pag_02... para ordenar certo
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as z:  # JPEG já vem comprimido
        for page in doc:
            z.writestr(f"{base}_pag_{page.number + 1:0{digits}d}.jpg", jpg(page))
    return buf.getvalue(), f"{base}_imagens.zip"


def pdf_to_excel(files, form, tmp):
    from openpyxl import Workbook
    from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
    path, base = files[0]
    doc = open_pdf(path)
    wb = Workbook()
    wb.remove(wb.active)
    for page in doc:
        for n, table in enumerate(page.find_tables().tables, 1):
            ws = wb.create_sheet(f"Pág {page.number + 1} - Tabela {n}")
            for row in table.extract():
                ws.append([ILLEGAL_CHARACTERS_RE.sub('', cell) if cell else None for cell in row])
    if not wb.sheetnames:
        raise UserError("Nenhuma tabela encontrada neste PDF. Se ele for escaneado, passe o OCR antes.")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), f"{base}.xlsx"


PDF = ('.pdf',)
WORD = ('.doc', '.docx', '.odt', '.rtf')
EXCEL = ('.xls', '.xlsx', '.ods', '.csv')
POWERPOINT = ('.ppt', '.pptx', '.odp')
IMAGES = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.tif', '.tiff')

# ação: (função, extensões aceitas, aceita vários arquivos)
ACTIONS = {
    "merge-pdf": (merge_pdf, PDF, True),
    "split-pdf": (split_pdf, PDF, False),
    "word-to-pdf": (office_action("Word.Application"), WORD, False),
    "excel-to-pdf": (office_action("Excel.Application"), EXCEL, False),
    "ppt-to-pdf": (office_action("PowerPoint.Application"), POWERPOINT, False),
    "jpg-to-pdf": (images_to_pdf, IMAGES, True),
    "pdf-to-word": (pdf_to_word, PDF, False),
    "pdf-to-jpg": (pdf_to_jpg, PDF, False),
    "pdf-to-excel": (pdf_to_excel, PDF, False),
}


def save_uploads(uploads, tmp):
    """Grava os envios com nomes gerados aqui; o nome original só vira nome do download."""
    files = []
    for i, f in enumerate(uploads):
        original = Path(f.filename.replace('\\', '/').rsplit('/', 1)[-1])
        suffix = original.suffix.lower() if re.fullmatch(r'\.[a-z0-9]{1,5}', original.suffix.lower()) else ''
        path = tmp / f"entrada{i}{suffix}"
        f.save(path)
        files.append((path, original.stem or "arquivo"))
    return files


@app.route('/')
def index():
    return render_template('index.html')

def check_uploads(files, accepted, multiple):
    if not files:
        raise UserError("Escolha um arquivo.")
    if len(files) > 1 and not multiple:
        raise UserError("Esta ferramenta aceita um arquivo por vez.")
    for f in files:
        if Path(f.filename).suffix.lower() not in accepted:
            raise UserError(f"\"{f.filename}\" não é do tipo aceito aqui ({', '.join(accepted)}).")


@app.errorhandler(413)
def too_large(e):
    return f"Arquivo grande demais: o limite é {MAX_UPLOAD_MB} MB.", 413

@app.route('/convert', methods=['POST'])
def handle_conversion():
    action = request.form.get('action')
    files = [f for f in request.files.getlist('file') if f.filename]
    if action not in ACTIONS: return "Ferramenta desconhecida.", 400
    handler, accepted, multiple = ACTIONS[action]

    try:
        check_uploads(files, accepted, multiple)
        # Tudo acontece numa pasta temporária própria, apagada ao fim da requisição
        with tempfile.TemporaryDirectory(prefix='euamopdf-') as tmp:
            tmp = Path(tmp)
            saved = save_uploads(files, tmp)
            if any(path.stat().st_size == 0 for path, _ in saved):
                raise UserError("O arquivo enviado está vazio.")
            data, download_name = handler(saved, request.form, tmp)
    except UserError as e:
        return str(e), 400
    except Exception:
        app.logger.exception("Falha em %s", action)
        return "Não foi possível processar o arquivo. Veja se ele abre normalmente em outro programa.", 500

    return send_file(io.BytesIO(data), as_attachment=True, download_name=download_name)

def free_port(preferred=5000):
    """Usa a porta preferida se estiver livre; senão, qualquer porta livre."""
    with socket.socket() as s:
        try:
            s.bind(('127.0.0.1', preferred))
        except OSError:
            s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]

if __name__ == '__main__':
    port = free_port()
    # O Timer aguarda 1 segundo para garantir que o servidor Flask já subiu
    Timer(1, webbrowser.open_new, [f"http://127.0.0.1:{port}/"]).start()
    app.run(host='127.0.0.1', port=port)
