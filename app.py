from flask import Flask, render_template, request, send_file
import io
import re
import sys
import shutil
import subprocess
import tempfile
import webbrowser # Biblioteca para abrir o navegador
from pathlib import Path
from threading import Timer # Para atrasar a abertura em 1 segundo
from urllib.parse import urlparse

if sys.platform == 'win32':
    import win32com.client
    import pythoncom

from PIL import Image
from pdf2image import convert_from_path
from pdf2docx import Converter
import zipfile
from PyPDF2 import PdfReader, PdfWriter, PdfMerger


app = Flask(__name__)

# --- CONFIGURAÇÕES ---
BASE_DIR = Path(__file__).resolve().parent
POPPLER_PATH = BASE_DIR / 'poppler-25.12.0' / 'Library' / 'bin'
LOCAL_HOSTS = {'127.0.0.1', 'localhost'}
MAX_UPLOAD_MB = 500
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


def office_to_pdf(input_path, output_path, app_name):
    in_p = str(input_path)
    out_p = str(output_path)
    out_dir = str(output_path.parent)

    # No Windows, tenta primeiro via COM / win32com (Microsoft Office)
    if sys.platform == 'win32':
        try:
            pythoncom.CoInitialize()
            try:
                app_inst = win32com.client.Dispatch(app_name)
                if "Word" in app_name:
                    doc = app_inst.Documents.Open(in_p)
                    doc.SaveAs(out_p, FileFormat=17)
                    doc.Close()
                elif "Excel" in app_name:
                    app_inst.Visible = False
                    wb = app_inst.Workbooks.Open(in_p)
                    wb.ExportAsFixedFormat(0, out_p)
                    wb.Close()
                elif "PowerPoint" in app_name:
                    pres = app_inst.Presentations.Open(in_p, WithWindow=False)
                    pres.SaveAs(out_p, 32)
                    pres.Close()
                app_inst.Quit()
                return
            finally:
                pythoncom.CoUninitialize()
        except Exception:
            pass # Fallback para LibreOffice se falhar no Windows

    # No Linux (ou como fallback no Windows), usa LibreOffice
    cmd = shutil.which('libreoffice') or shutil.which('soffice')
    if not cmd:
        raise RuntimeError("LibreOffice não encontrado no sistema para conversão de documentos Office.")

    res = subprocess.run([cmd, '--headless', '-env:UserInstallation=file:///tmp/LibreOffice_Profile', '--convert-to', 'pdf', in_p, '--outdir', out_dir], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"Erro ao converter com LibreOffice: {res.stderr}")

    generated_pdf = input_path.with_suffix('.pdf')
    if generated_pdf.exists() and generated_pdf != output_path:
        generated_pdf.replace(output_path)


def open_pdf(path):
    try:
        reader = PdfReader(path)
    except Exception:
        raise UserError("O arquivo não é um PDF válido ou está corrompido.")
    if reader.is_encrypted:
        raise UserError("Este PDF está protegido por senha.")
    return reader


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
    merger = PdfMerger()
    for path, _ in files:
        merger.append(open_pdf(path))
    buf = io.BytesIO()
    merger.write(buf); merger.close()
    return buf.getvalue(), "PDF_Unido.pdf"


def split_pdf(files, form, tmp):
    path, base = files[0]
    reader = open_pdf(path)
    spec = form.get('pages', '').strip()

    if spec:
        writer = PdfWriter()
        for i in parse_pages(spec, len(reader.pages)):
            writer.add_page(reader.pages[i])
        buf = io.BytesIO()
        writer.write(buf)
        return buf.getvalue(), f"{base}_paginas.pdf"

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for i, page in enumerate(reader.pages):
            w = PdfWriter(); w.add_page(page)
            page_buf = io.BytesIO()
            w.write(page_buf)
            z.writestr(f"pag_{i+1}.pdf", page_buf.getvalue())
    return buf.getvalue(), f"{base}_dividido.zip"


def office_action(app_name):
    def convert(files, form, tmp):
        path, base = files[0]
        out = tmp / 'saida.pdf'
        office_to_pdf(path, out, app_name)
        return out.read_bytes(), f"{base}.pdf"
    return convert


def jpg_to_pdf(files, form, tmp):
    path, base = files[0]
    buf = io.BytesIO()
    try:
        img = Image.open(path)
    except Exception:
        raise UserError("O arquivo não é uma imagem válida.")
    img.convert('RGB').save(buf, 'PDF')
    return buf.getvalue(), f"{base}.pdf"


def pdf_to_word(files, form, tmp):
    path, base = files[0]
    open_pdf(path)
    out = tmp / 'saida.docx'
    cv = Converter(str(path)); cv.convert(str(out)); cv.close()
    return out.read_bytes(), f"{base}.docx"


def pdf_to_jpg(files, form, tmp):
    path, base = files[0]
    open_pdf(path)
    p_path = POPPLER_PATH if (sys.platform == 'win32' and POPPLER_PATH.exists()) else None
    imgs = convert_from_path(path, poppler_path=p_path)
    buf = io.BytesIO()
    imgs[0].save(buf, 'JPEG')
    return buf.getvalue(), f"{base}.jpg"


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
    "jpg-to-pdf": (jpg_to_pdf, IMAGES, False),
    "pdf-to-word": (pdf_to_word, PDF, False),
    "pdf-to-jpg": (pdf_to_jpg, PDF, False),
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

def open_browser():
    """Abre o navegador no endereço local"""
    webbrowser.open_new("http://127.0.0.1:5000/")

if __name__ == '__main__':
    # O Timer aguarda 1 segundo para garantir que o servidor Flask já subiu
    Timer(1, open_browser).start()
    app.run(host='127.0.0.1', port=5000)
