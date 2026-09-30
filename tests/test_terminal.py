"""O app vive preso a um terminal: fechar o terminal encerra o processo, e o AppRun do AppImage abre
um terminal quando o arquivo é aberto por duplo clique, sem terminal nenhum."""
import os
import select
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest

if sys.platform == 'win32':
    pytest.skip("pty, SIGHUP e AppRun são do Linux", allow_module_level=True)
import pty  # noqa: E402 (não existe no Windows)

ROOT = Path(__file__).resolve().parent.parent


def wait_for(terminal, text, timeout=30):
    """Lê o que o processo escreve no terminal até aparecer o texto."""
    out, deadline = b'', time.time() + timeout
    while text not in out:
        assert time.time() < deadline, f"{text!r} não apareceu:\n{out.decode(errors='replace')}"
        if select.select([terminal], [], [], 0.5)[0]:
            out += os.read(terminal, 4096)
    return out


@pytest.mark.filterwarnings("ignore:This process .* is multi-threaded:DeprecationWarning")  # o filho só faz exec
def test_fechar_o_terminal_encerra_o_app():
    pid, terminal = pty.fork()  # o filho ganha um terminal de controle, como numa janela de terminal
    if pid == 0:
        os.execve(sys.executable, [sys.executable, str(ROOT / 'app.py')], os.environ | {'BROWSER': 'true'})
    try:
        wait_for(terminal, b'Running on')
        os.close(terminal)  # o que o emulador faz ao fechar a janela: o kernel manda SIGHUP ao processo
        deadline = time.time() + 10
        while os.waitpid(pid, os.WNOHANG) == (0, 0):
            assert time.time() < deadline, "o app continuou rodando depois que o terminal foi fechado"
            time.sleep(0.1)
    finally:
        try:
            os.kill(pid, 9)
        except ProcessLookupError:
            pass


# --- AppRun ---

APPRUN = ROOT / 'packaging' / 'AppRun'
TERMINALS = [('x-terminal-emulator', '-e'), ('gnome-terminal', '--'), ('konsole', '-e'),
             ('xfce4-terminal', '-x'), ('xterm', '-e')]


@pytest.fixture
def appdir(tmp_path):
    """O AppRun verdadeiro numa pasta com um 'app' que só deixa um rastro. O PATH tem só o que o
    script usa, então o que o teste instalar em bin/ é o único terminal que existe."""
    shutil.copy(APPRUN, tmp_path / 'AppRun')
    app = tmp_path / 'usr' / 'bin' / 'EuAmoPDF' / 'EuAmoPDF'
    app.parent.mkdir(parents=True)
    (tmp_path / 'bin').mkdir()
    for tool in ('sh', 'dirname'):
        (tmp_path / 'bin' / tool).symlink_to(shutil.which(tool))
    script(app, 'echo "app $@" >> "$TRACE"')
    return tmp_path


def script(path, body):
    path.write_text('#!/bin/sh\n' + body + '\n')
    path.chmod(0o755)


def fake_terminal(appdir, name, runs_command=False):
    """Terminal de mentira: anota como foi chamado e, se pedido, roda o comando que recebeu."""
    body = 'echo "terminal $@" >> "$TRACE"' + ('\nshift\nexec "$@"' if runs_command else '')
    script(appdir / 'bin' / name, body)


def run_apprun(appdir, tty=False, **env):
    """Roda o AppRun sem terminal (como um duplo clique) ou com um. Devolve (rastro, saída)."""
    master, slave = os.openpty() if tty else (None, None)
    stdio = dict(stdin=slave, stdout=slave, stderr=slave) if tty else \
        dict(stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    env = {'PATH': str(appdir / 'bin'), 'TRACE': str(appdir / 'trace'), **env}
    try:
        done = subprocess.run([appdir / 'AppRun'], env=env, timeout=30, text=True, **stdio)
    finally:
        if tty:
            os.close(master), os.close(slave)
    trace = appdir / 'trace'
    return (trace.read_text().splitlines() if trace.exists() else []), done.stdout


DUPLO_CLIQUE = {'APPIMAGE': '/opt/EuAmoPDF.AppImage', 'DISPLAY': ':0'}


@pytest.mark.parametrize("name, flag", TERMINALS)
def test_sem_terminal_abre_um_e_roda_o_appimage_dentro_dele(appdir, name, flag):
    fake_terminal(appdir, name)
    trace, _ = run_apprun(appdir, **DUPLO_CLIQUE)
    assert len(trace) == 1 and trace[0].startswith(f'terminal {flag} sh -c ')
    assert trace[0].endswith('/opt/EuAmoPDF.AppImage')


@pytest.mark.parametrize("code, avisa", [(0, False), (130, False), (143, False), (1, True), (3, True)])
def test_o_terminal_roda_o_appimage_e_so_avisa_se_ele_parar_com_erro(appdir, code, avisa):
    fake_terminal(appdir, 'x-terminal-emulator', runs_command=True)
    appimage = appdir / 'EuAmoPDF.AppImage'
    script(appimage, f'echo "appimage" >> "$TRACE"\nexit {code}')
    trace, out = run_apprun(appdir, APPIMAGE=str(appimage), DISPLAY=':0')
    assert trace[-1] == 'appimage'
    assert ('parou com erro' in out) == avisa


def test_prefere_o_terminal_do_sistema(appdir):
    for name, _ in TERMINALS:
        fake_terminal(appdir, name)
    trace, _ = run_apprun(appdir, **DUPLO_CLIQUE)
    assert trace[0].startswith('terminal -e sh -c ')  # x-terminal-emulator, não o gnome-terminal


@pytest.mark.parametrize("tty, env", [
    (True, DUPLO_CLIQUE),  # já está num terminal
    (False, {'DISPLAY': ':0'}),  # AppDir aberto direto, como no teste do build: não há AppImage para reabrir
    (False, {'APPIMAGE': '/opt/EuAmoPDF.AppImage'}),  # sem ambiente gráfico
], ids=["ja-tem-terminal", "sem-appimage", "sem-display"])
def test_roda_direto_quando_nao_deve_ou_nao_pode_abrir_terminal(appdir, tty, env):
    fake_terminal(appdir, 'x-terminal-emulator')
    trace, _ = run_apprun(appdir, tty, **env)
    assert trace == ['app ']


def test_sem_nenhum_terminal_instalado_roda_direto(appdir):
    trace, _ = run_apprun(appdir, **DUPLO_CLIQUE)
    assert trace == ['app ']
