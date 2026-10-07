"""Gera o EuAmoPDF num arquivo só, com o Python e tudo o que o app usa embutido:

    Linux:   dist/EuAmoPDF.AppImage
    Windows: dist/EuAmoPDF.exe

    pip install -r requirements-build.txt
    python packaging/build.py

O PyInstaller não gera para outro sistema: cada versão é construída no seu (o workflow
executaveis.yml faz as duas no GitHub). No fim, o arquivo pronto é aberto de verdade e processa
alguns documentos: se faltar um módulo ou um arquivo embutido, o build falha aqui, não no usuário.
"""
import hashlib
import io
import os
import re
import shutil
import signal
import subprocess
import sys
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path

import pymupdf
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
BUILD, DIST = ROOT / 'build', ROOT / 'dist'
ICON = ROOT / 'packaging' / 'icone.png'  # o PyInstaller converte para .ico no Windows
WINDOWS = sys.platform == 'win32'
# Tudo o que o build baixa vai numa versão fixa e conferido pelo SHA-256: o appimagetool roda aqui, e
# o runtime fica na frente de todo AppImage gerado (a versão "continuous" mudava a cada commit deles).
# Para atualizar, troque a versão e o hash juntos.
TESSDATA_URL = 'https://github.com/tesseract-ocr/tessdata_fast/raw/4.1.0/{}.traineddata'
TESSDATA_SHA256 = {
    'por': 'c4932b937207a9514b7514d518b931a99938c02a28a5a5a553f8599ed58b7deb',
    'eng': '7d4322bd2a7749724879683fc3912cb542f19906c83bcc1a52132556427170b2',
}
APPIMAGETOOL = ('https://github.com/AppImage/appimagetool/releases/download/1.9.1/appimagetool-x86_64.AppImage',
                'ed4ce84f0d9caff66f50bcca6ff6f35aae54ce8135408b3fa33abfc3cb384eb0')
RUNTIME = ('https://github.com/AppImage/type2-runtime/releases/download/20251108/runtime-x86_64',
           '2fca8b443c92510f1483a883f60061ad09b46b978b2631c807cd873a47ec260d')
# Terminal=true: o app não tem janela própria, e o terminal é como se fecha (igual ao console no Windows)
DESKTOP = """[Desktop Entry]
Type=Application
Name=EuAmoPDF
Comment=Ferramentas de PDF que rodam no seu computador
Exec=EuAmoPDF
Icon=euamopdf
Categories=Office;Utility;
Terminal=true
"""


def sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def download(url, path, digest):
    """Baixa para um arquivo parcial e só o põe no lugar se o SHA-256 conferir; um arquivo já baixado
    também é conferido, e baixado de novo se não bater."""
    if not path.exists() or sha256(path) != digest:
        path.parent.mkdir(parents=True, exist_ok=True)
        partial = path.with_name(path.name + '.part')
        urllib.request.urlretrieve(url, partial)
        if sha256(partial) != digest:
            partial.unlink()
            sys.exit(f'{url}: o SHA-256 não confere')
        partial.replace(path)
    return path


def pyinstaller(dist, mode, tessdata, model):
    """O modelo, os idiomas do OCR e a interface vão para dentro; o app os acha por app.bundled()."""
    data = [(ROOT / 'templates', 'templates'), (ROOT / 'static', 'static'), (tessdata, 'tessdata'), (model, '.')]
    cmd = [sys.executable, '-m', 'PyInstaller', '--noconfirm', '--name', 'EuAmoPDF', mode,
           '--distpath', str(dist), '--workpath', str(BUILD / 'pyinstaller'), '--specpath', str(BUILD)]
    if WINDOWS:
        cmd += ['--icon', str(ICON)]
    for src, dest in data:
        cmd += ['--add-data', f'{src}{os.pathsep}{dest}']
    subprocess.run([*cmd, str(ROOT / 'app.py')], check=True)


def build_appimage(tessdata, model):
    appdir = BUILD / 'EuAmoPDF.AppDir'
    shutil.rmtree(appdir, ignore_errors=True)
    pyinstaller(appdir / 'usr' / 'bin', '--onedir', tessdata, model)  # sem descompactar a cada abertura
    run = appdir / 'AppRun'
    shutil.copy(ROOT / 'packaging' / 'AppRun', run)
    run.chmod(0o755)
    (appdir / 'euamopdf.desktop').write_text(DESKTOP)
    shutil.copy(ICON, appdir / 'euamopdf.png')
    smoke_test(run)  # o mesmo binário e AppRun que vão no AppImage, sem extrair 570 MB em /tmp a cada teste
    tool = download(APPIMAGETOOL[0], BUILD / 'appimagetool-1.9.1', APPIMAGETOOL[1])
    tool.chmod(0o755)
    runtime = download(RUNTIME[0], BUILD / 'runtime-20251108', RUNTIME[1])
    out = DIST / 'EuAmoPDF.AppImage'
    # EXTRACT_AND_RUN: roda a ferramenta (que também é um AppImage) sem precisar do FUSE; sem o
    # --runtime-file, ela baixaria a versão do dia do runtime
    subprocess.run([tool, '--no-appstream', '--runtime-file', runtime, appdir, out], check=True,
                   env=os.environ | {'ARCH': 'x86_64', 'APPIMAGE_EXTRACT_AND_RUN': '1'})
    return out


def build_exe(tessdata, model):
    pyinstaller(BUILD / 'pyinstaller-dist', '--onefile', tessdata, model)
    out = Path(shutil.move(BUILD / 'pyinstaller-dist' / 'EuAmoPDF.exe', DIST / 'EuAmoPDF.exe'))
    smoke_test(out)
    return out


def post(base, action, filename, data):
    boundary = uuid.uuid4().hex
    parts = [('action', '', action.encode()), ('file', f'; filename="{filename}"', data)]
    body = b''.join(f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"{extra}\r\n\r\n'.encode()
                    + value + b'\r\n' for name, extra, value in parts) + f'--{boundary}--\r\n'.encode()
    request = urllib.request.Request(f'{base}/convert', body, {'Content-Type': f'multipart/form-data; boundary={boundary}'})
    try:
        with urllib.request.urlopen(request, timeout=300) as response:
            return response.read()
    except urllib.error.HTTPError as e:
        sys.exit(f'{action}: HTTP {e.code}: {e.read().decode(errors="replace")}')


def check_tools(base):
    with urllib.request.urlopen(base + '/') as page, urllib.request.urlopen(base + '/static/estilo.css') as css:
        assert b'EuAmoPDF' in page.read() and css.status == 200, 'interface (templates/static) não embutida'
    doc = pymupdf.open()
    doc.new_page().insert_text((72, 100), 'Ola mundo', fontsize=30)
    text_pdf = doc.tobytes()
    scan = pymupdf.open()  # a mesma página só como imagem: só o OCR consegue ler
    scan.new_page().insert_image(scan[0].rect, stream=doc[0].get_pixmap(dpi=200).tobytes('png'))
    circle = Image.new('RGB', (256, 256), 'white')
    ImageDraw.Draw(circle).ellipse((64, 64, 192, 192), fill='red')
    png = io.BytesIO()
    circle.save(png, 'PNG')

    assert post(base, 'pdf-to-word', 'a.pdf', text_pdf)[:2] == b'PK', 'pdf2docx'
    assert post(base, 'pdf-to-ppt', 'a.pdf', text_pdf)[:2] == b'PK', 'python-pptx'
    ocr = post(base, 'ocr-pdf', 'a.pdf', scan.tobytes())
    assert 'mundo' in pymupdf.open(stream=ocr)[0].get_text(), 'OCR: idiomas do Tesseract não embutidos'
    assert post(base, 'remove-background', 'a.png', png.getvalue())[:4] == b'\x89PNG', 'remover fundo (modelo)'


def smoke_test(cmd):
    """Abre o executável e usa as ferramentas que dependem do que foi embutido."""
    log = BUILD / 'smoke.log'
    with open(log, 'w') as f:
        proc = subprocess.Popen([cmd], stdout=f, stderr=subprocess.STDOUT, start_new_session=True)
    try:
        deadline = time.time() + 300  # o .exe se descompacta na pasta temporária a cada abertura
        while not (found := re.search(r'http://127\.0\.0\.1:(\d+)', log.read_text(errors='replace'))):
            if proc.poll() is not None or time.time() > deadline:
                sys.exit('O executável não subiu:\n' + log.read_text(errors='replace'))
            time.sleep(1)
        check_tools(f'http://127.0.0.1:{found[1]}')
    finally:
        if WINDOWS:  # o .exe é um lançador que cria outro processo: derruba os dois
            subprocess.run(['taskkill', '/F', '/T', '/PID', str(proc.pid)], capture_output=True)
        else:
            os.killpg(proc.pid, signal.SIGTERM)
        proc.wait()


def main():
    tessdata = BUILD / 'tessdata'
    for lang in ('por', 'eng'):
        download(TESSDATA_URL.format(lang), tessdata / f'{lang}.traineddata', TESSDATA_SHA256[lang])
    sys.path.insert(0, str(ROOT))
    import app  # reaproveita o download do modelo de remover fundo, que confere o SHA-256
    model = app.bg_model_path()
    if not model.exists() or sha256(model) != app.BG_MODEL_SHA256:  # o do cache do CI também é conferido
        app.download_bg_model(model)
    DIST.mkdir(exist_ok=True)
    out = (build_exe if WINDOWS else build_appimage)(tessdata, model)
    print(f'Pronto: {out} ({out.stat().st_size / 1024 / 1024:.0f} MB)')


if __name__ == '__main__':
    main()
