from flask import Flask, render_template, request, send_file
import os
import sys

if sys.platform == 'win32':
    import win32com.client
    import pythoncom

from PIL import Image
from pdf2image import convert_from_path
from pdf2docx import Converter
from pptx import Presentation
import pandas as pd
import pdfplumber
import subprocess
import shutil
import time
import zipfile
import webbrowser # Biblioteca para abrir o navegador
from threading import Timer # Para atrasar a abertura em 1 segundo
from PyPDF2 import PdfReader, PdfWriter, PdfMerger


app = Flask(__name__)

# --- CONFIGURAÇÕES ---
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
POPPLER_PATH = r'poppler-25.12.0\Library\bin'

for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
    os.makedirs(folder, exist_ok=True)

def cleanup_files():
    """Remove arquivos com mais de 10 minutos"""
    now = time.time()
    for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
        for f in os.listdir(folder):
            path = os.path.join(folder, f)
            if os.stat(path).st_mtime < now - 600:
                try: os.remove(path)
                except: pass

def office_to_pdf(input_path, output_path, app_name):
    in_p = os.path.abspath(input_path)
    out_p = os.path.abspath(output_path)
    out_dir = os.path.dirname(out_p)

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

    generated_pdf = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(in_p))[0]}.pdf")
    if os.path.exists(generated_pdf) and generated_pdf != out_p:
        os.replace(generated_pdf, out_p)

def open_browser():
    """Abre o navegador no endereço local"""
    webbrowser.open_new("http://127.0.0.1:5000/")

@app.route('/')
def index():
    cleanup_files()
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def handle_conversion():
    action = request.form.get('action')
    files = request.files.getlist('file')
    if not files or files[0].filename == '': return "Nenhum arquivo", 400

    try:
        # --- JUNTAR PDF ---
        if action == "merge-pdf":
            merger = PdfMerger()
            for f in files:
                path = os.path.join(UPLOAD_FOLDER, f.filename)
                f.save(path)
                merger.append(path)
            out_p = os.path.join(OUTPUT_FOLDER, "PDF_Unido.pdf")
            merger.write(out_p); merger.close()
            return send_file(os.path.abspath(out_p), as_attachment=True)

        # --- DIVIDIR / INTERVALO ---
        file = files[0]
        in_p = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(in_p)
        base = os.path.splitext(file.filename)[0]

        if action == "split-pdf":
            reader = PdfReader(in_p)
            start = request.form.get('page_start')
            end = request.form.get('page_end')
            
            if start and end:
                writer = PdfWriter()
                for i in range(int(start)-1, min(int(end), len(reader.pages))):
                    writer.add_page(reader.pages[i])
                out_p = os.path.join(OUTPUT_FOLDER, f"{base}_recorte.pdf")
                with open(out_p, "wb") as f: writer.write(f)
                return send_file(os.path.abspath(out_p), as_attachment=True)
            else:
                zip_p = os.path.join(OUTPUT_FOLDER, f"{base}_dividido.zip")
                with zipfile.ZipFile(zip_p, 'w') as z:
                    for i, page in enumerate(reader.pages):
                        p_path = os.path.join(OUTPUT_FOLDER, f"pag_{i+1}.pdf")
                        w = PdfWriter(); w.add_page(page)
                        with open(p_path, "wb") as f: w.write(f)
                        z.write(p_path, os.path.basename(p_path))
                return send_file(os.path.abspath(zip_p), as_attachment=True)

        # --- OUTRAS CONVERSÕES ---
        out_p = os.path.join(OUTPUT_FOLDER, f"{base}.pdf") # Default
        if action == "word-to-pdf": office_to_pdf(in_p, out_p, "Word.Application")
        elif action == "excel-to-pdf": office_to_pdf(in_p, out_p, "Excel.Application")
        elif action == "ppt-to-pdf": office_to_pdf(in_p, out_p, "PowerPoint.Application")
        elif action == "jpg-to-pdf": Image.open(in_p).convert('RGB').save(out_p)
        elif action == "pdf-to-word":
            out_p = os.path.join(OUTPUT_FOLDER, f"{base}.docx")
            cv = Converter(in_p); cv.convert(out_p); cv.close()
        elif action == "pdf-to-jpg":
            out_p = os.path.join(OUTPUT_FOLDER, f"{base}.jpg")
            p_path = POPPLER_PATH if (os.name == 'nt' and os.path.exists(POPPLER_PATH)) else None
            imgs = convert_from_path(in_p, poppler_path=p_path)
            imgs[0].save(out_p, 'JPEG')
        
        return send_file(os.path.abspath(out_p), as_attachment=True)

    except Exception as e: return f"Erro: {str(e)}", 500

if __name__ == '__main__':
    # O Timer aguarda 1 segundo para garantir que o servidor Flask já subiu
    # O parâmetro 'use_reloader=False' evita que o navegador abra duas vezes ao iniciar
    Timer(1, open_browser).start()
    app.run(debug=True, use_reloader=False)