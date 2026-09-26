from flask import Flask, render_template, request, send_file
from werkzeug.exceptions import HTTPException
import base64
import hashlib
import io
import json
import math
import os
import random
import re
import sys
import shutil
import socket
import subprocess
import tempfile
import threading
import time
import urllib.request
import webbrowser # Biblioteca para abrir o navegador
import zipfile
import zlib
from pathlib import Path
from threading import Timer # Para atrasar a abertura em 1 segundo
from urllib.parse import quote, urlparse

import numpy as np
import pymupdf
from PIL import Image, ImageChops, ImageOps


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


def has_alpha(img):
    return img.mode in ('RGBA', 'LA', 'PA') or 'transparency' in img.info


def on_white(img):
    """Imagem em RGB, com a transparência sobre fundo branco."""
    if not has_alpha(img):
        return img.convert('RGB')
    img = img.convert('RGBA')
    bg = Image.new('RGB', img.size, 'white')
    bg.paste(img, mask=img.getchannel('A'))
    return bg


def open_image(path, reduce_to=None):
    """Abre a imagem já na orientação da foto: (imagem, formato original). Com reduce_to, um JPEG
    grande é decodificado direto num tamanho menor (mas ainda o dobro de reduce_to): bem mais rápido."""
    try:
        with Image.open(path) as original:
            fmt = 'JPEG' if original.format == 'MPO' else original.format  # MPO: JPEG de alguns celulares
            if reduce_to:
                original.draft(original.mode, (reduce_to * 2, reduce_to * 2))
            img = ImageOps.exif_transpose(original)
            img.load()
    except Exception:
        raise UserError("O arquivo não é uma imagem válida.")
    return img, fmt


def load_image(path):
    """Imagem pronta para o PDF: na orientação da foto e com a transparência sobre fundo branco."""
    img, fmt = open_image(path)
    img = on_white(img)
    buf = io.BytesIO()
    if fmt == 'JPEG':  # foto continua JPEG; o resto vira PNG para não borrar texto
        img.save(buf, 'JPEG', quality=95)
    else:
        img.save(buf, 'PNG')
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


PDF_LEVELS = {'recomendada': (150, 75), 'forte': (96, 55)}  # (DPI alvo, qualidade do JPEG)


def pdf_images(doc):
    """Imagens que a compressão pode mexer: {xref: (página, largura, altura, menor DPI em que aparecem)}."""
    found = {}
    for page in doc:
        for item in page.get_images(full=True):
            xref, smask, width, height, bpc = item[:5]
            if smask or bpc == 1 or doc.xref_get_key(xref, 'Mask')[0] != 'null':
                continue  # transparência e imagens de 1 bit (texto escaneado em preto e branco) ficam como estão
            rect = page.get_image_bbox(item)  # bem mais rápido que get_image_rects, que decodifica a imagem
            if rect.is_empty or rect.is_infinite:
                continue
            dpi = 72 * math.sqrt(width * height / (rect.width * rect.height))  # vale também para imagem girada
            if xref not in found or dpi < found[xref][3]:
                found[xref] = (page.number, width, height, dpi)
    return found


def stored_size(doc, xref, raw=None):
    """Tamanho da imagem como fica no PDF gravado: stream sem compressão ganha deflate ao salvar."""
    raw = doc.xref_stream_raw(xref) if raw is None else raw
    return len(zlib.compress(raw)) if doc.xref_get_key(xref, 'Filter')[0] == 'null' else len(raw)


def shrink_pdf_image(doc, xref, width, height, dpi, target_dpi, quality):
    """A imagem em JPEG, reduzida à resolução alvo; None se isso não a deixar menor."""
    raw = doc.xref_stream_raw(xref)
    is_jpeg = doc.xref_get_key(xref, 'Filter')[1] == '/DCTDecode'
    if dpi <= target_dpi + 10 and not is_jpeg:
        return None  # imagem sem perda (desenho, captura de tela) só vira JPEG se tiver resolução demais
    scale = min(1, target_dpi / dpi)
    size = (max(1, round(width * scale)), max(1, round(height * scale)))
    try:
        img = None
        if is_jpeg and doc.xref_get_key(xref, 'Decode')[0] == 'null':
            img = Image.open(io.BytesIO(raw))
            if img.mode in ('RGB', 'L'):
                img.draft(img.mode, size)  # decodifica o JPEG já reduzido: bem mais rápido
            else:
                img = None  # CMYK e afins: o MuPDF converte as cores melhor
        if img is None:
            pix = pymupdf.Pixmap(doc, xref)
            if pix.n not in (1, 3):
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            img = pix.pil_image()
        img = img.convert('L' if img.mode == 'L' else 'RGB')
    except Exception:
        return None  # formato de imagem que não dá para decodificar: fica como está
    if img.size != size:
        img = img.resize(size, Image.Resampling.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, 'JPEG', quality=quality, optimize=True)
    return buf.getvalue() if buf.tell() < stored_size(doc, xref, raw) else None


def compress_pdf(files, form, tmp):
    path, base = files[0]
    doc = open_pdf(path)
    target_dpi, quality = PDF_LEVELS.get(form.get('level'), PDF_LEVELS['recomendada'])
    for xref, (page_number, width, height, dpi) in pdf_images(doc).items():
        smaller = shrink_pdf_image(doc, xref, width, height, dpi, target_dpi, quality)
        if smaller:
            doc[page_number].replace_image(xref, stream=smaller)  # vale para todas as páginas que a usam
    data = doc.tobytes(garbage=4, deflate=True, clean=True, use_objstms=1)  # garbage=4 junta as cópias
    before = path.stat().st_size
    if len(data) >= before:
        return path.read_bytes(), f"{base}.pdf", "Este PDF já está otimizado: não deu para reduzir mais."
    return data, f"{base}_comprimido.pdf", \
        f"Reduzido de {format_size(before)} para {format_size(len(data))} ({1 - len(data) / before:.0%} menor)."


# nível: (qualidade do JPG e do WebP, cores do PNG; None = PNG sem perda)
IMAGE_LEVELS = {
    'leve': (85, None),
    'recomendada': (72, 256),
    'forte': (55, 128),
    'extrema': (35, 64),
}
IMAGE_EXTENSIONS = {'JPEG': '.jpg', 'PNG': '.png', 'WEBP': '.webp'}
MAX_SIDES = {'2560': 2560, '1920': 1920, '1280': 1280, '800': 800}


def image_settings(form):
    """Opções do Comprimir imagem: (qualidade, cores do PNG, lado máximo, formato, remover dados da foto)."""
    level = form.get('image_level') if form.get('image_level') in IMAGE_LEVELS else 'recomendada'
    return (*IMAGE_LEVELS[level], MAX_SIDES.get(form.get('max_size')),
            form.get('image_format', '').upper(), bool(form.get('strip_metadata')))


def prepare_image(path, max_side, target, strip_metadata):
    """A imagem já reduzida e o que é preciso para gravá-la:
    (imagem, formato final, opções de gravação, se muda algo além da compressão)."""
    img, fmt = open_image(path, reduce_to=max_side)
    out_fmt = target if target in IMAGE_EXTENSIONS else fmt if fmt in IMAGE_EXTENSIONS else 'PNG'
    resized = bool(max_side) and max(img.size) > max_side
    if resized:
        img.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)

    options = {}
    if img.info.get('icc_profile'):
        options['icc_profile'] = img.info['icc_profile']  # perfil de cor: sem ele as cores mudam
    has_exif = bool(img.getexif())
    if has_exif and not strip_metadata:
        options['exif'] = img.getexif()  # a orientação já foi aplicada e retirada
    return img, out_fmt, options, out_fmt != fmt or resized or (has_exif and strip_metadata)


def encode_image(img, fmt, quality, colors, options=None, fast_png=False):
    options = options or {}
    buf = io.BytesIO()
    if fmt == 'JPEG':
        on_white(img).save(buf, 'JPEG', quality=quality, optimize=True, progressive=True, **options)
    elif fmt == 'WEBP':
        # method 6 é o mais compacto, mas com transparência fica 20 vezes mais lento para ganhar ~3%
        img = img.convert('RGBA' if has_alpha(img) else 'RGB')
        img.save(buf, 'WEBP', quality=quality, method=5 if img.mode == 'RGBA' else 6, **options)
    else:
        img = img.convert('RGBA' if has_alpha(img) else 'RGB')
        if colors:  # PNG com menos cores, como faz o TinyPNG
            method = Image.Quantize.FASTOCTREE if img.mode == 'RGBA' else Image.Quantize.MEDIANCUT
            img = img.quantize(colors, method=method, dither=Image.Dither.FLOYDSTEINBERG)
        img.save(buf, 'PNG', **({'compress_level': 6} if fast_png else {'optimize': True}), **options)
    return buf.getvalue()


def compress_image(path, settings):
    """Devolve (bytes, extensão, se ficou como estava)."""
    quality, colors, max_side, target, strip_metadata = settings
    img, fmt, options, changed = prepare_image(path, max_side, target, strip_metadata)
    data = encode_image(img, fmt, quality, colors, options)
    # Se não deu para diminuir e não havia nada a mudar (formato, tamanho, dados da foto), fica o original
    if len(data) >= path.stat().st_size and not changed:
        return path.read_bytes(), IMAGE_EXTENSIONS[fmt], True
    return data, IMAGE_EXTENSIONS[fmt], False


def pack(results, zip_name):
    """Um resultado vai direto; vários vão num ZIP. results: [(nome sem extensão, extensão, bytes)]."""
    if len(results) == 1:
        stem, ext, data = results[0]
        return data, f"{stem}{ext}"
    buf, used = io.BytesIO(), set()
    with zipfile.ZipFile(buf, 'w') as z:  # imagens já vêm comprimidas
        for stem, ext, data in results:
            name, n = f"{stem}{ext}", 2
            while name in used:  # dois arquivos com o mesmo nome não podem se sobrescrever no ZIP
                name, n = f"{stem}_{n}{ext}", n + 1
            used.add(name)
            z.writestr(name, data)
    return buf.getvalue(), zip_name


def compress_images(files, form, tmp):
    settings = image_settings(form)
    results, kept = [], 0
    before = after = 0
    for path, base in files:
        data, ext, unchanged = compress_image(path, settings)
        results.append((f"{base}_comprimida", ext, data))
        kept += unchanged
        before += path.stat().st_size
        after += len(data)

    if len(files) == 1 and kept:
        message = "A imagem já estava bem comprimida e ficou como estava."
    else:
        prefix = f"{len(files)} imagens: de" if len(files) > 1 else "Reduzida de"
        message = f"{prefix} {format_size(before)} para {format_size(after)} ({max(0, 1 - after / before):.0%} menor)."
        if kept == 1:
            message += " Uma já estava bem comprimida e ficou como estava."
        elif kept:
            message += f" {kept} já estavam bem comprimidas e ficaram como estavam."
    return *pack(results, "Imagens_comprimidas.zip"), message


# --- REMOVER FUNDO ---
# Modelo de IA ISNet (projeto DIS, licença Apache-2.0), no formato ONNX publicado pelo rembg.
# É baixado só na primeira vez e conferido pelo SHA-256; as imagens nunca saem do computador.
# Dos modelos testados, foi o de melhor recorte que roda bem num computador comum: ~1 s por
# foto e ~800 MB de memória (o BiRefNet passou de 8 GB).
BG_MODEL_URL = 'https://github.com/danielgatis/rembg/releases/download/v0.0.0/isnet-general-use.onnx'
BG_MODEL_SHA256 = '60920e99c45464f2ba57bee2ad08c919a52bbf852739e96947fbb4358c0d964a'
BG_MODEL_MB = 170
BG_MODEL_SIDE = 1024  # o modelo vê a imagem em 1024x1024; a máscara volta ao tamanho original
_bg_session = None
_bg_lock = threading.Lock()


def bg_model_path():
    """Na pasta de cache do sistema: %LOCALAPPDATA% no Windows, ~/.cache no Linux."""
    base = os.environ.get('LOCALAPPDATA') or os.environ.get('XDG_CACHE_HOME') or Path.home() / '.cache'
    return Path(base) / 'euamopdf' / 'isnet-general-use.onnx'


def download_bg_model(path):
    """Baixa o modelo para um arquivo temporário e só o põe no lugar se o SHA-256 conferir."""
    path.parent.mkdir(parents=True, exist_ok=True)
    partial = path.with_name(path.name + '.part')
    digest = hashlib.sha256()
    try:
        with urllib.request.urlopen(BG_MODEL_URL, timeout=30) as response, open(partial, 'wb') as out:
            while chunk := response.read(1 << 20):
                digest.update(chunk)
                out.write(chunk)
    except OSError:
        partial.unlink(missing_ok=True)
        raise UserError("Não foi possível baixar o modelo de IA que remove o fundo. "
                        "Só na primeira vez é preciso estar conectado à internet.")
    if digest.hexdigest() != BG_MODEL_SHA256:
        partial.unlink(missing_ok=True)
        raise UserError("O modelo de IA baixado veio corrompido. Tente de novo.")
    os.replace(partial, path)


def bg_session():
    """Sessão do modelo: baixado na primeira vez e carregado uma vez só."""
    global _bg_session
    with _bg_lock:  # também impede dois downloads ao mesmo tempo
        if _bg_session is None:
            import onnxruntime  # importação lenta: só quando usada
            path = bg_model_path()
            if not path.exists():
                download_bg_model(path)
            options = onnxruntime.SessionOptions()
            options.enable_cpu_mem_arena = False  # pico de ~800 MB em vez de ~1,2 GB, por ~0,3 s a mais
            _bg_session = onnxruntime.InferenceSession(str(path), options, providers=['CPUExecutionProvider'])
        return _bg_session


def subject_mask(img):
    """Máscara do objeto principal (0 a 255), no tamanho da imagem."""
    session = bg_session()
    small = img.convert('RGB').resize((BG_MODEL_SIDE, BG_MODEL_SIDE), Image.Resampling.LANCZOS)
    x = np.asarray(small, dtype=np.float32) / 255
    x = (x - 0.5).transpose(2, 0, 1)[None]  # normalização do ISNet: média 0,5 e desvio 1, canais primeiro
    out = session.run(None, {session.get_inputs()[0].name: x})[0][0, 0]
    # A confiança do modelo fica em ~0,985 no objeto e raramente é 0 no fundo: esticar 0,05-0,95
    # para 0-1 deixa o objeto opaco e o fundo limpo, mantendo a borda suave
    out = (out - 0.05) / 0.9
    mask = Image.fromarray((np.clip(out, 0, 1) * 255).round().astype(np.uint8))
    # Ampliação bicúbica: nos testes, refinar a borda com guided filter piorou o recorte
    return mask.resize(img.size, Image.Resampling.BICUBIC)


BACKGROUNDS = {'branco': '#ffffff', 'preto': '#000000'}
BG_FORMATS = {'png': ('PNG', '.png', {}), 'webp': ('WEBP', '.webp', {'lossless': True})}  # sempre sem perda


def background_color(form):
    """Cor do novo fundo, ou None para transparente."""
    choice = form.get('background', 'transparente')
    if choice != 'cor':
        return BACKGROUNDS.get(choice)
    color = form.get('background_color', '')
    if not re.fullmatch(r'#[0-9a-fA-F]{6}', color):
        raise UserError("Escolha uma cor válida para o fundo.")
    return color


def load_mask(upload, size, what="máscara"):
    """Máscara em tons de cinza enviada pela interface; tem de ter o tamanho da imagem."""
    if upload is None:
        raise UserError(f"Falta a {what}.")
    try:
        mask = Image.open(upload.stream)
        mask.load()
    except Exception:
        raise UserError(f"A {what} enviada não é uma imagem válida.")
    if mask.size != size:
        raise UserError(f"A {what} não tem o tamanho da imagem.")
    return mask.convert('L')


def load_strokes(upload, size):
    """Traço pintado na prévia (onde o canal alfa não é zero), levado ao tamanho da imagem."""
    if upload is None:
        raise UserError("Falta o traço.")
    try:
        strokes = Image.open(upload.stream).convert('RGBA').getchannel('A')
    except Exception:
        raise UserError("O traço enviado não é uma imagem válida.")
    painted = np.asarray(strokes.resize(size, Image.Resampling.NEAREST)) > 0
    if not painted.any():
        raise UserError("Pinte por cima da área que quer corrigir.")
    return painted


REFINE_SIDE = 1024  # o GrabCut roda num recorte de até 1024 px em volta do traço


def grabcut_region(img, mask, painted, restore, box, reach):
    """Região a mudar dentro do recorte box: (região no tamanho do recorte, se encosta na borda)."""
    import cv2  # importação lenta: só quando usada
    h, w = mask.shape
    x0, y0, x1, y1 = box
    scale = min(1, REFINE_SIDE / max(x1 - x0, y1 - y0))
    size = (max(1, round((x1 - x0) * scale)), max(1, round((y1 - y0) * scale)))
    crop = np.asarray(img.convert('RGB').crop(box).resize(size, Image.Resampling.BILINEAR))[..., ::-1].copy()
    cm = np.asarray(Image.fromarray(mask[y0:y1, x0:x1]).resize(size, Image.Resampling.BILINEAR))
    cs = np.asarray(Image.fromarray(painted[y0:y1, x0:x1]).resize(size, Image.Resampling.NEAREST))
    # Perto do traço, parte-se do que o usuário pediu, e o GrabCut devolve ao outro lado o que
    # tiver a cor e as bordas do outro lado. Longe dele vale a máscara atual, travada onde o
    # modelo teve certeza.
    near = cv2.distanceTransform((~cs).astype(np.uint8), cv2.DIST_L2, 3) <= reach * scale
    labels = np.where(cm > 127, cv2.GC_PR_FGD, cv2.GC_PR_BGD).astype(np.uint8)
    labels[(cm >= 250) & ~near] = cv2.GC_FGD
    labels[(cm <= 5) & ~near] = cv2.GC_BGD
    labels[near] = cv2.GC_PR_FGD if restore else cv2.GC_PR_BGD
    labels[cs] = cv2.GC_FGD if restore else cv2.GC_BGD
    cv2.setRNGSeed(0)  # mesmo traço, mesmo resultado
    try:
        cv2.grabCut(crop, labels, None, np.zeros((1, 65)), np.zeros((1, 65)), 4, cv2.GC_INIT_WITH_MASK)
        wanted = np.isin(labels, (cv2.GC_FGD, cv2.GC_PR_FGD) if restore else (cv2.GC_BGD, cv2.GC_PR_BGD))
    except cv2.error:  # recorte só de objeto ou só de fundo: não há com o que comparar
        wanted = cs
    # Restaurar vale para o que não está totalmente opaco; apagar, para o que está minimamente
    # visível. O que foi pintado sempre muda, e além dele só a região ligada ao traço.
    to_change = (wanted & ((cm < 255) if restore else (cm > 0))) | cs
    _, parts = cv2.connectedComponents(to_change.astype(np.uint8), connectivity=8)
    touched = np.unique(parts[cs])
    region = np.isin(parts, touched[touched > 0])
    edge = (region[0].any() and y0 > 0) or (region[-1].any() and y1 < h) or \
           (region[:, 0].any() and x0 > 0) or (region[:, -1].any() and x1 < w)
    full = Image.fromarray(region.astype(np.uint8) * 255).resize((x1 - x0, y1 - y0), Image.Resampling.BILINEAR)
    return np.asarray(full), edge


def refine_mask(img, mask, painted, restore):
    """Estende o traço do usuário à região que ele quis marcar (restaurar ou apagar) com o
    GrabCut, e devolve a máscara corrigida. O resto da máscara não muda."""
    m = np.asarray(mask, dtype=np.uint8)
    h, w = m.shape
    ys, xs = np.nonzero(painted)
    reach = max(64, int(0.25 * max(h, w)), 2 * int(max(np.ptp(xs), np.ptp(ys))))  # vizinhança do traço
    pad = 2 * reach
    while True:  # se a região encosta na borda do recorte, ela continua além dele: amplia e refaz
        box = (max(0, xs.min() - pad), max(0, ys.min() - pad), min(w, xs.max() + pad + 1), min(h, ys.max() + pad + 1))
        region, edge = grabcut_region(img, m, painted, restore, box, reach)
        if not edge or box == (0, 0, w, h):
            break
        pad *= 2
    x0, y0, x1, y1 = box
    out = m.copy()
    part = out[y0:y1, x0:x1]
    out[y0:y1, x0:x1] = np.maximum(part, region) if restore else np.minimum(part, 255 - region)
    return Image.fromarray(out)


def remove_background(path, background, fmt, mask_upload=None):
    """A imagem original sem o fundo, na mesma resolução e com os mesmos pixels no objeto:
    (bytes, extensão, se foi encontrado um objeto em destaque). Com mask_upload, usa a máscara
    corrigida na prévia em vez de rodar o modelo."""
    img, _ = open_image(path)
    # Perfil de cor CMYK ou de cinza não serve para a imagem RGBA gravada
    icc = img.info.get('icc_profile') if img.mode in ('RGB', 'RGBA', 'P') else None
    alpha = load_mask(mask_upload, img.size) if mask_upload else subject_mask(img)
    if has_alpha(img):  # a transparência que a imagem já tinha continua valendo
        alpha = ImageChops.multiply(alpha, img.convert('RGBA').getchannel('A'))
    out = img.convert('RGBA')
    out.putalpha(alpha)
    if background:
        canvas = Image.new('RGB', img.size, background)
        canvas.paste(out, mask=alpha)
        out = canvas
    else:
        # Pixel totalmente transparente guardaria a cor do fundo original, que qualquer um veria
        # tirando a transparência. A borda semitransparente mantém a cor da fonte.
        out = Image.composite(out, Image.new('RGBA', img.size), alpha.point(lambda v: 255 if v else 0))
    pil_fmt, ext, options = BG_FORMATS[fmt]
    buf = io.BytesIO()
    out.save(buf, pil_fmt, **options, **({'icc_profile': icc} if icc else {}))
    return buf.getvalue(), ext, alpha.getextrema()[1] >= 128


def remove_backgrounds(files, form, tmp):
    background = background_color(form)
    fmt = form.get('bg_format') if form.get('bg_format') in BG_FORMATS else 'png'
    masks = request.files.getlist('mascara')  # corrigidas na prévia, uma por imagem, na mesma ordem
    if masks and len(masks) != len(files):
        raise UserError("Envie uma máscara para cada imagem.")
    results, missed = [], []
    for i, (path, base) in enumerate(files):
        data, ext, found = remove_background(path, background, fmt, masks[i] if masks else None)
        results.append((f"{base}_sem_fundo", ext, data))
        if not found:
            missed.append(base)
    message = f"Fundo removido de {len(files)} imagem(ns)."
    if missed:
        message = f"Nenhum objeto em destaque encontrado em: {', '.join(missed)}. Confira o resultado."
    return *pack(results, "Imagens_sem_fundo.zip"), message


# --- ESTIMATIVA DE TAMANHO ---
# Rápida o bastante para rodar a cada mudança de opção: comprime de verdade só uma amostra
# e extrapola. Devolve (tamanho atual, tamanho estimado) em bytes.

def spread(items, n):
    """Até n itens espalhados pela lista."""
    return items if len(items) <= n else [items[i * len(items) // n] for i in range(n)]


# O optimize do PNG deixa o arquivo ~5% menor que o nível padrão, mas é 5 a 10 vezes mais lento
PNG_OPTIMIZE_GAIN = 0.95


def mosaic(img, tile=64, grid=23):
    """Pedaços espalhados pela imagem, montados numa amostra de ~2 MP: (amostra, quantas vezes a
    imagem é maior que ela). Imagens pequenas vão inteiras. Com fotos reais, o erro médio da
    estimativa fica em ~5%; pedaços maiores (e em menor número) erram mais."""
    w, h = img.size
    if w * h <= 2 * (tile * grid) ** 2:
        return img, 1
    img = img.convert('RGBA' if has_alpha(img) else 'L' if img.mode == 'L' else 'RGB')
    sample = Image.new(img.mode, (tile * grid, tile * grid))
    for i in range(grid):
        for j in range(grid):
            x, y = (w - tile) * i // (grid - 1), (h - tile) * j // (grid - 1)
            sample.paste(img.crop((x, y, x + tile, y + tile)), (i * tile, j * tile))
    return sample, w * h / (tile * grid) ** 2


def estimate_image(path, settings):
    quality, colors, max_side, target, strip_metadata = settings
    img, fmt, options, changed = prepare_image(path, max_side, target, strip_metadata)
    sample, scale = mosaic(img)
    # Perfil de cor e EXIF não crescem com a imagem: entram uma vez só
    extra = sum(len(value if isinstance(value, bytes) else value.tobytes()) for value in options.values())
    lossless_png = fmt == 'PNG' and not colors
    size = len(encode_image(sample, fmt, quality, colors, fast_png=lossless_png)) * scale + extra
    if lossless_png:
        size *= PNG_OPTIMIZE_GAIN
    before = path.stat().st_size
    return before if size >= before and not changed else size  # como no compress_image


ESTIMATE_SECONDS = 2  # tempo para estimar imagens; as que não couberem são extrapoladas
ESTIMATE_MIN_IMAGES = 3  # mesmo passando do tempo: extrapolar de uma imagem só erra muito num lote misturado


def estimate_images(files, form):
    settings = image_settings(form)
    sample_before = sample_after = 0
    deadline = time.monotonic() + ESTIMATE_SECONDS
    order = random.Random(0).sample(files, len(files))  # embaralhada: a amostra fica espalhada
    for count, (path, _) in enumerate(order, 1):
        sample_before += path.stat().st_size
        sample_after += estimate_image(path, settings)
        if count >= ESTIMATE_MIN_IMAGES and time.monotonic() > deadline:
            break
    before = sum(path.stat().st_size for path, _ in files)
    return before, before * sample_after / sample_before


def estimate_pdf(files, form):
    path, _ = files[0]
    doc = open_pdf(path)
    target_dpi, quality = PDF_LEVELS.get(form.get('level'), PDF_LEVELS['recomendada'])
    images = pdf_images(doc)
    raw = {xref: stored_size(doc, xref) for xref in images}
    sample = spread(list(images.items()), 8)
    sample_before = sum(raw[xref] for xref, _ in sample)
    sample_after = 0
    for xref, (_, width, height, dpi) in sample:
        smaller = shrink_pdf_image(doc, xref, width, height, dpi, target_dpi, quality)
        sample_after += len(smaller) if smaller else raw[xref]
    # O PDF limpo como o compress_pdf grava, trocando o tamanho das imagens pelo estimado
    cleaned = len(doc.tobytes(garbage=4, deflate=True, clean=True, use_objstms=1))
    images_before = sum(raw.values())
    after = cleaned - images_before + images_before * (sample_after / sample_before if sample_before else 1)
    before = path.stat().st_size
    return before, min(before, after)  # o compress_pdf devolve o original quando não diminui


ESTIMATORS = {
    'compress-pdf': estimate_pdf,
    'compress-image': estimate_images,
}


PDF = ('.pdf',)
WORD = ('.doc', '.docx', '.odt', '.rtf')
EXCEL = ('.xls', '.xlsx', '.ods', '.csv')
POWERPOINT = ('.ppt', '.pptx', '.odp')
IMAGES = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif', '.tif', '.tiff')
WEB_IMAGES = ('.jpg', '.jpeg', '.png', '.webp')

# ação: (função, extensões aceitas, aceita vários arquivos)
ACTIONS = {
    "merge-pdf": (merge_pdf, PDF, True),
    "compress-pdf": (compress_pdf, PDF, False),
    "compress-image": (compress_images, WEB_IMAGES, True),
    "remove-background": (remove_backgrounds, WEB_IMAGES, True),
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
    return render_template('index.html', tools=tools, bg_model_ready=bg_model_path().exists(), bg_model_mb=BG_MODEL_MB)

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


def process_uploads(action, work):
    """Valida os arquivos enviados para a ferramenta, grava numa pasta temporária e devolve
    work(arquivos, pasta). A pasta é apagada ao fim, com o que houver dentro."""
    if action not in ACTIONS:
        raise UserError("Ferramenta desconhecida.")
    _, accepted, multiple = ACTIONS[action]
    files = [f for f in request.files.getlist('file') if f.filename]
    check_uploads(files, accepted, multiple)
    with tempfile.TemporaryDirectory(prefix='euamopdf-') as tmp:
        tmp = Path(tmp)
        saved = save_uploads(files, tmp)
        if any(path.stat().st_size == 0 for path, _ in saved):
            raise UserError("O arquivo enviado está vazio.")
        return work(saved, tmp)


@app.errorhandler(UserError)
def user_error(e):
    return str(e), 400

@app.errorhandler(413)
def too_large(e):
    return f"Arquivo grande demais: o limite é {MAX_UPLOAD_MB} MB.", 413

@app.errorhandler(Exception)
def unexpected_error(e):
    if isinstance(e, HTTPException):
        return e  # 404, 405...: a resposta padrão do Flask já serve
    app.logger.exception("Falha em %s", request.form.get('action'))
    return "Não foi possível processar o arquivo. Veja se ele abre normalmente em outro programa.", 500

@app.route('/convert', methods=['POST'])
def handle_conversion():
    action = request.form.get('action')
    data, download_name, *message = process_uploads(
        action, lambda saved, tmp: ACTIONS[action][0](saved, request.form, tmp))
    response = send_file(io.BytesIO(data), as_attachment=True, download_name=download_name)
    if message:  # cabeçalhos HTTP só aceitam ASCII
        response.headers['X-Mensagem'] = quote(message[0])
    return response

def png_response(img):
    buf = io.BytesIO()
    img.save(buf, 'PNG')
    buf.seek(0)
    return send_file(buf, mimetype='image/png')

def single_image(saved):
    if len(saved) != 1:
        raise UserError("Envie uma imagem por vez.")
    return open_image(saved[0][0])[0]

@app.route('/background/mask', methods=['POST'])
def background_mask():
    """Máscara do objeto (PNG em tons de cinza, do tamanho da imagem), para a prévia editável."""
    return png_response(process_uploads('remove-background', lambda saved, tmp: subject_mask(single_image(saved))))

@app.route('/background/refine', methods=['POST'])
def background_refine():
    """Corrige a máscara a partir do traço pintado na prévia: restaura ou apaga a região marcada."""
    mode = request.form.get('modo')
    if mode not in ('restaurar', 'apagar'):
        raise UserError("Escolha entre restaurar e apagar.")

    def work(saved, tmp):
        img = single_image(saved)
        mask = load_mask(request.files.get('mascara'), img.size)
        painted = load_strokes(request.files.get('traco'), img.size)
        return refine_mask(img, mask, painted, mode == 'restaurar')
    return png_response(process_uploads('remove-background', work))

@app.route('/estimate', methods=['POST'])
def estimate_size():
    """Tamanho aproximado do resultado das ferramentas de compressão, com as opções escolhidas."""
    action = request.form.get('action')
    if action not in ESTIMATORS:
        raise UserError("Esta ferramenta não tem estimativa de tamanho.")
    before, after = process_uploads(action, lambda saved, tmp: ESTIMATORS[action](saved, request.form))
    return {'antes': before, 'depois': round(after)}

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
