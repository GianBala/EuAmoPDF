from flask import Flask, render_template, request, send_file
import base64
import io
import json
import os
import re
import sys
import shutil
import socket
import subprocess
import tempfile
import webbrowser # Biblioteca para abrir o navegador
import zipfile
from pathlib import Path
from threading import Timer # Para atrasar a abertura em 1 segundo
from urllib.parse import quote, urlparse

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


def program_file(*parts):
    """Caminho dentro de Program Files, no Windows, onde os instaladores não mexem no PATH."""
    if sys.platform != 'win32':
        return None
    for base in (os.environ.get('PROGRAMFILES'), os.environ.get('PROGRAMFILES(X86)')):
        if base and (path := Path(base, *parts)).exists():
            return str(path)
    return None


def find_soffice():
    return shutil.which('soffice') or shutil.which('libreoffice') or \
        program_file('LibreOffice', 'program', 'soffice.exe')


def find_tessdata():
    """Pasta de idiomas do Tesseract, ou None se ele não estiver instalado."""
    try:
        return pymupdf.get_tessdata()
    except Exception:
        return program_file('Tesseract-OCR', 'tessdata')


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


def open_pdf(source, password=None):
    """Abre um PDF a partir do caminho ou dos bytes do arquivo, com a senha se ele pedir uma."""
    try:
        # Abre da memória: no Windows, um arquivo que fica aberto impede apagar a pasta temporária
        data = source if isinstance(source, bytes) else Path(source).read_bytes()
        doc = pymupdf.open(stream=data, filetype='pdf')
    except Exception:
        raise UserError("O arquivo não é um PDF válido ou está corrompido.")
    if doc.needs_pass:
        if password is None:
            raise UserError("Este PDF está protegido por senha. Use \"Remover senha\" antes.")
        if not doc.authenticate(password):
            raise UserError("Senha incorreta.")
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
# o formulário e a pasta temporária da requisição, e devolve (bytes, nome do download)
# e, opcionalmente, uma mensagem para a interface mostrar.

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
        with Image.open(path) as original:
            fmt = original.format
            img = ImageOps.exif_transpose(original)
            img.load()
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
    cv = Converter(str(path))
    try:
        cv.convert(str(out))
    finally:
        cv.close()
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


def pdf_to_ppt(files, form, tmp):
    """Cada página vira um slide com a imagem da página (o texto não fica editável)."""
    from pptx import Presentation
    from pptx.util import Pt
    path, base = files[0]
    doc = open_pdf(path)
    prs = Presentation()
    first = doc[0].rect
    scale = min(1, 4000 / max(first.width, first.height))  # o PowerPoint limita o slide a 56 polegadas
    prs.slide_width, prs.slide_height = Pt(first.width * scale), Pt(first.height * scale)
    for page in doc:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # layout em branco
        # Encaixa a página no slide sem distorcer, caso ela tenha outro formato
        fit = min(prs.slide_width / page.rect.width, prs.slide_height / page.rect.height)
        w, h = round(page.rect.width * fit), round(page.rect.height * fit)
        image = io.BytesIO(page.get_pixmap(dpi=150).tobytes('jpg', jpg_quality=92))
        slide.shapes.add_picture(image, (prs.slide_width - w) // 2, (prs.slide_height - h) // 2, w, h)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue(), f"{base}.pptx"


def organize_pdf(files, form, tmp):
    """Monta o PDF na ordem que o usuário escolheu nas miniaturas: [{"pagina": 3, "giro": 90}, ...]."""
    path, base = files[0]
    doc = open_pdf(path)
    try:
        order = [(int(item['pagina']), int(item.get('giro', 0))) for item in json.loads(form.get('order', ''))]
    except (ValueError, TypeError, KeyError, AttributeError):
        raise UserError("Não entendi a nova ordem das páginas.")
    if not order:
        raise UserError("Deixe ao menos uma página.")
    out = pymupdf.open()
    for number, turn in order:
        if not 1 <= number <= doc.page_count or turn % 90:
            raise UserError("Não entendi a nova ordem das páginas.")
        out.insert_pdf(doc, from_page=number - 1, to_page=number - 1)
        out[-1].set_rotation((out[-1].rotation + turn) % 360)
    return pdf_bytes(out), f"{base}_organizado.pdf"


def thumbnail(page, width=150):
    zoom = width / page.rect.width
    jpg = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom)).tobytes('jpg', jpg_quality=70)
    return 'data:image/jpeg;base64,' + base64.b64encode(jpg).decode()


def rotate_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    angle = {'90': 90, '180': 180, '270': 270}.get(form.get('angle'))
    if angle is None:
        raise UserError("Escolha para que lado girar.")
    spec = form.get('pages', '').strip()
    for i in set(parse_pages(spec, doc.page_count)) if spec else range(doc.page_count):
        doc[i].set_rotation((doc[i].rotation + angle) % 360)
    return pdf_bytes(doc), f"{base}_girado.pdf"


def protect_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    password = form.get('password', '')
    if not password:
        raise UserError("Escolha uma senha.")
    if password != form.get('password2'):
        raise UserError("As senhas digitadas não são iguais.")
    data = doc.tobytes(garbage=3, deflate=True, encryption=pymupdf.PDF_ENCRYPT_AES_256,
                       user_pw=password, owner_pw=password)
    return data, f"{base}_protegido.pdf"


def unlock_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path, password=form.get('password', ''))
    # Também remove restrições (impressão, cópia...) de PDFs que abrem sem senha
    if not doc.metadata.get('encryption'):
        raise UserError("Este PDF não tem senha nem restrições para remover.")
    return doc.tobytes(garbage=3, deflate=True, encryption=pymupdf.PDF_ENCRYPT_NONE), f"{base}_sem_senha.pdf"


def check_text(text, what):
    """As fontes padrão do PDF só têm os caracteres do Windows-1252 (acentos do português incluídos)."""
    if not text:
        raise UserError(f"Escreva o texto {what}.")
    try:
        text.encode('cp1252')
    except UnicodeEncodeError:
        raise UserError("Use só letras, números, acentos e pontuação comuns (sem emojis ou símbolos especiais).")


def watermark_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    text = form.get('text', '').strip()
    check_text(text, "da marca d'água")
    for page in doc:
        r = page.rect  # área visível, já considerando a rotação da página
        size = min(100, 0.75 * abs(r.br - r.tl) / pymupdf.get_text_length(text, 'helv', 1))
        center = pymupdf.Point(r.width / 2, r.height / 2) * page.derotation_matrix
        start = pymupdf.Point(center.x - pymupdf.get_text_length(text, 'helv', size) / 2, center.y + size * 0.35)
        # Diagonal subindo da esquerda para a direita, qualquer que seja a rotação da página
        page.insert_text(start, text, fontsize=size, fontname='helv', color=(0.5, 0.5, 0.5),
                         fill_opacity=0.3, morph=(center, pymupdf.Matrix(45 + page.rotation)))
    return pdf_bytes(doc), f"{base}_marca_dagua.pdf"


NUMBER_FORMATS = {
    'n': "{n}",
    'n-de-t': "{n} / {t}",
    'pagina': "Página {n} de {t}",
}


def number_pages(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    label = NUMBER_FORMATS.get(form.get('format'), NUMBER_FORMATS['n'])
    for page in doc:
        text = label.format(n=page.number + 1, t=doc.page_count)
        r = page.rect  # área visível, já considerando a rotação da página
        x = (r.width - pymupdf.get_text_length(text, 'helv', 10)) / 2
        # Centralizado no rodapé como o usuário vê, mesmo em páginas giradas
        page.insert_text(pymupdf.Point(x, r.height - 20) * page.derotation_matrix, text,
                         fontsize=10, fontname='helv', rotate=page.rotation)
    return pdf_bytes(doc), f"{base}_numerado.pdf"


def add_invisible_text(page, words):
    """Escreve as palavras do OCR invisíveis sobre a imagem, para buscar e selecionar."""
    # Todas as palavras de uma linha usam a mesma base e altura, para a linha não se partir
    lines = {}
    for x0, y0, x1, y1, word, block, line, _ in words:
        bottom, height = lines.get((block, line), (0, 0))
        lines[(block, line)] = (max(bottom, y1), max(height, y1 - y0))
    for x0, y0, x1, y1, word, block, line, _ in words:
        bottom, height = lines[(block, line)]
        word = word.encode('cp1252', 'replace').decode('cp1252')  # a fonte padrão só tem esses caracteres
        if x1 <= x0 or height <= 0 or not word.strip():
            continue  # caixa degenerada do OCR: não há onde escrever
        start = pymupdf.Point(x0, bottom - 0.2 * height)
        stretch = (x1 - x0) / pymupdf.get_text_length(word, 'helv', height)  # ocupa a largura da palavra
        page.insert_text(start, word + ' ', fontsize=height, fontname='helv', render_mode=3,
                         morph=(start, pymupdf.Matrix(stretch, 1)))


def ocr_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    tessdata = find_tessdata()
    languages = [lang for lang in ('por', 'eng') if tessdata and Path(tessdata, f'{lang}.traineddata').exists()]
    if 'por' not in languages:
        raise UserError("O OCR precisa do Tesseract instalado com o idioma português. Veja as instruções no README.")
    recognized = 0
    for page in doc:
        if page.get_text().strip():
            continue  # a página já tem texto
        page.remove_rotation()  # mesma aparência, mas com coordenadas iguais às que o OCR devolve
        textpage = page.get_textpage_ocr(dpi=300, language='+'.join(languages), tessdata=tessdata, full=True)
        add_invisible_text(page, page.get_text('words', textpage=textpage))
        recognized += 1
    if not recognized:
        raise UserError("Todas as páginas deste PDF já têm texto: não há o que reconhecer.")
    return pdf_bytes(doc), f"{base}_ocr.pdf", f"Texto reconhecido em {recognized} página(s)."


def format_size(size):
    return f"{size / 1024:.0f} KB" if size < 1024 * 1024 else f"{size / 1024 / 1024:.1f} MB".replace('.', ',')


def compress_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    # Imagens acima da resolução alvo são reduzidas e recomprimidas em JPEG
    dpi, quality = (96, 55) if form.get('level') == 'forte' else (150, 75)
    doc.rewrite_images(dpi_threshold=dpi + 10, dpi_target=dpi, quality=quality)
    data = doc.tobytes(garbage=4, deflate=True, clean=True, use_objstms=1)
    before = path.stat().st_size
    if len(data) >= before:
        return path.read_bytes(), f"{base}.pdf", "Este PDF já está otimizado: não deu para reduzir mais."
    return data, f"{base}_comprimido.pdf", \
        f"Reduzido de {format_size(before)} para {format_size(len(data))} ({1 - len(data) / before:.0%} menor)."


PDF = ('.pdf',)
WORD = ('.doc', '.docx', '.odt', '.rtf')
EXCEL = ('.xls', '.xlsx', '.ods', '.csv')
POWERPOINT = ('.ppt', '.pptx', '.odp')
IMAGES = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.tif', '.tiff')

# ação: (função, extensões aceitas, aceita vários arquivos)
ACTIONS = {
    "merge-pdf": (merge_pdf, PDF, True),
    "compress-pdf": (compress_pdf, PDF, False),
    "ocr-pdf": (ocr_pdf, PDF, False),
    "protect-pdf": (protect_pdf, PDF, False),
    "unlock-pdf": (unlock_pdf, PDF, False),
    "watermark-pdf": (watermark_pdf, PDF, False),
    "number-pages": (number_pages, PDF, False),
    "split-pdf": (split_pdf, PDF, False),
    "rotate-pdf": (rotate_pdf, PDF, False),
    "organize-pdf": (organize_pdf, PDF, False),
    "word-to-pdf": (office_action("Word.Application"), WORD, False),
    "excel-to-pdf": (office_action("Excel.Application"), EXCEL, False),
    "ppt-to-pdf": (office_action("PowerPoint.Application"), POWERPOINT, False),
    "jpg-to-pdf": (images_to_pdf, IMAGES, True),
    "pdf-to-word": (pdf_to_word, PDF, False),
    "pdf-to-jpg": (pdf_to_jpg, PDF, False),
    "pdf-to-excel": (pdf_to_excel, PDF, False),
    "pdf-to-ppt": (pdf_to_ppt, PDF, False),
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
    # Extensões aceitas e se aceita vários arquivos, para os botões da interface
    tools = {action: {'accept': ','.join(accepted), 'multiple': multiple}
             for action, (_, accepted, multiple) in ACTIONS.items()}
    return render_template('index.html', tools=tools)

@app.route('/pages', methods=['POST'])
def page_info():
    """Número de páginas do PDF escolhido e, se pedidas, as miniaturas para organizar."""
    f = request.files.get('file')
    try:
        doc = open_pdf(f.read() if f else b'')
    except UserError as e:
        return {'erro': str(e)}, 400
    info = {'paginas': doc.page_count}
    if request.form.get('miniaturas'):
        # ponytail: todas as miniaturas numa resposta só; paginar se PDFs de centenas de páginas ficarem lentos
        info['miniaturas'] = [thumbnail(page) for page in doc]
    return info

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
            data, download_name, *message = handler(saved, request.form, tmp)
    except UserError as e:
        return str(e), 400
    except Exception:
        app.logger.exception("Falha em %s", action)
        return "Não foi possível processar o arquivo. Veja se ele abre normalmente em outro programa.", 500

    response = send_file(io.BytesIO(data), as_attachment=True, download_name=download_name)
    if message:  # cabeçalhos HTTP só aceitam ASCII
        response.headers['X-Mensagem'] = quote(message[0])
    return response

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
