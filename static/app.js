const $ = (id) => document.getElementById(id);
const dialog = $('dialog');
const form = $('form');
const input = $('file-input');
const drop = $('drop');
const submit = $('submit');

let tool = null;   // data-* do botão da ferramenta aberta
let files = [];    // arquivos escolhidos, na ordem em que serão enviados
let pageInfoRequest = 0;

document.querySelectorAll('.tool').forEach((button) => button.addEventListener('click', () => openTool(button)));
$('close').addEventListener('click', () => dialog.close());

function openTool(button) {
    tool = button.dataset;
    files = [];
    form.reset();
    $('dialog-title').textContent = tool.title;
    $('dialog-desc').textContent = button.querySelector('small').textContent;
    input.accept = tool.accept;
    input.multiple = 'multiple' in tool;
    $('drop-text').textContent = input.multiple
        ? 'Arraste os arquivos aqui ou clique para escolher'
        : 'Arraste o arquivo aqui ou clique para escolher';

    // Mostra só os campos da ferramenta; campos escondidos ficam desativados e não são enviados
    const fields = tool.fields.split(' ');
    document.querySelectorAll('[data-field]').forEach((field) => {
        field.hidden = !fields.includes(field.dataset.field);
        field.querySelectorAll('input, select').forEach((el) => { el.disabled = field.hidden; });
    });
    $('hint').textContent = tool.hint;
    $('hint').hidden = !tool.hint;

    setStatus('');
    setBusy(false);
    render();
    dialog.showModal();
}

// --- Escolha de arquivos ---

function addFiles(chosen) {
    const extensions = tool.accept.split(',');
    const accepted = [];
    const rejected = [];
    for (const file of chosen) {
        (extensions.some((ext) => file.name.toLowerCase().endsWith(ext)) ? accepted : rejected).push(file);
    }
    let problem = '';
    if (input.multiple) {
        files = files.concat(accepted);
    } else if (accepted.length) {
        files = [accepted[0]];
        if (accepted.length > 1) problem = `Esta ferramenta usa um arquivo por vez. Fiquei com ${accepted[0].name}.`;
    }
    if (rejected.length) problem = `Tipo não aceito nesta ferramenta: ${rejected.map((f) => f.name).join(', ')}`;
    setStatus(problem, 'error');
    render();
}

input.addEventListener('change', () => {
    addFiles(input.files);
    input.value = '';  // permite escolher o mesmo arquivo de novo
});

drop.addEventListener('dragover', (event) => { event.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (event) => {
    event.preventDefault();
    drop.classList.remove('over');
    addFiles(event.dataTransfer.files);
});
// Arquivo solto fora da área não deve abrir no navegador e tirar o usuário do app
window.addEventListener('dragover', (event) => event.preventDefault());
window.addEventListener('drop', (event) => event.preventDefault());

function render() {
    const list = $('file-list');
    list.replaceChildren(...files.map((file, i) => {
        const item = document.createElement('li');
        const name = document.createElement('span');
        name.className = 'name';
        name.textContent = file.name;
        name.title = file.name;
        const size = document.createElement('span');
        size.className = 'size';
        size.textContent = formatSize(file.size);
        item.append(name, size);
        if (input.multiple) {
            item.append(
                iconButton('↑', `Mover ${file.name} para cima`, i === 0, () => move(i, -1)),
                iconButton('↓', `Mover ${file.name} para baixo`, i === files.length - 1, () => move(i, 1)),
            );
        }
        item.append(iconButton('✕', `Remover ${file.name}`, false, () => { files.splice(i, 1); render(); }));
        return item;
    }));
    submit.disabled = files.length === 0;
    showPageInfo();
}

function iconButton(text, label, disabled, onClick) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'icon-button';
    button.textContent = text;
    button.setAttribute('aria-label', label);
    button.disabled = disabled;
    button.addEventListener('click', onClick);
    return button;
}

function move(i, step) {
    [files[i], files[i + step]] = [files[i + step], files[i]];
    render();
}

function formatSize(bytes) {
    if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1).replace('.', ',')} MB`;
}

async function showPageInfo() {
    const info = $('page-info');
    const request = ++pageInfoRequest;  // ignora respostas de um arquivo que já foi trocado
    info.hidden = true;
    if (files.length !== 1 || !files[0].name.toLowerCase().endsWith('.pdf')) return;
    const data = new FormData();
    data.append('file', files[0]);
    try {
        const response = await fetch('/pages', { method: 'POST', body: data });
        const body = await response.json();
        if (request !== pageInfoRequest) return;
        info.textContent = response.ok
            ? `Este PDF tem ${body.paginas} página${body.paginas === 1 ? '' : 's'}.`
            : body.erro;
        info.hidden = false;
    } catch {
        // Sem a contagem de páginas a ferramenta continua funcionando
    }
}

// --- Envio ---

form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (!files.length) return;
    const data = new FormData(form);
    data.append('action', tool.action);
    files.forEach((file) => data.append('file', file));

    setBusy(true);
    setStatus('');
    try {
        const response = await fetch('/convert', { method: 'POST', body: data });
        if (!response.ok) {
            setStatus(await response.text(), 'error');
            return;
        }
        download(await response.blob(), fileName(response));
        setStatus('Pronto! O arquivo foi baixado.', 'ok');
    } catch {
        setStatus('Não foi possível falar com o EuAmoPDF. Ele ainda está aberto?', 'error');
    } finally {
        setBusy(false);
    }
});

function fileName(response) {
    const header = response.headers.get('Content-Disposition') || '';
    const utf8 = header.match(/filename\*=UTF-8''([^;]+)/i);
    if (utf8) return decodeURIComponent(utf8[1]);
    const plain = header.match(/filename="?([^";]+)"?/i);
    return plain ? plain[1] : 'resultado';
}

function download(blob, name) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = name;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 60_000);
}

function setBusy(busy) {
    submit.disabled = busy || !files.length;
    submit.classList.toggle('busy', busy);
    submit.textContent = busy ? 'Processando…' : tool.title;
}

function setStatus(message, kind) {
    const status = $('status');
    status.textContent = message;
    status.className = `status ${kind || ''}`;
    status.hidden = !message;
}
