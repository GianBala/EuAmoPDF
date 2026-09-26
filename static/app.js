const $ = (id) => document.getElementById(id);
const dialog = $('dialog');
const form = $('form');
const input = $('file-input');
const drop = $('drop');
const submit = $('submit');

let tool = null;   // data-* do botão da ferramenta aberta
let files = [];    // arquivos escolhidos, na ordem em que serão enviados
let pageInfoRequest = 0;
let pages = [];     // Organizar: {pagina, giro, src} na nova ordem
let dragged = null; // índice da miniatura sendo arrastada

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
    updateDropText();
    showPageInfo();
}

function updateDropText() {
    const text = $('drop-text');
    if (!input.multiple) {
        text.textContent = files.length
            ? 'Arraste ou clique para trocar o arquivo'
            : 'Arraste o arquivo aqui ou clique para escolher';
        return;
    }
    // Cada escolha soma à lista; na janela de arquivos, clicar sem Ctrl troca a seleção
    text.textContent = files.length
        ? 'Arraste ou clique para adicionar mais arquivos'
        : 'Arraste os arquivos aqui ou clique para escolher';
    const tip = document.createElement('small');
    tip.className = 'drop-tip';
    tip.textContent = 'Para marcar vários de uma vez na janela de arquivos, segure Ctrl ao clicar em cada um.';
    text.append(document.createElement('br'), tip);
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

function organizing() {
    return tool.fields.split(' ').includes('organizer');
}

async function showPageInfo() {
    const info = $('page-info');
    const request = ++pageInfoRequest;  // ignora respostas de um arquivo que já foi trocado
    info.hidden = true;
    pages = [];
    renderPages();
    if (files.length !== 1 || !files[0].name.toLowerCase().endsWith('.pdf')) return;
    const data = new FormData();
    data.append('file', files[0]);
    if (organizing()) {
        data.append('miniaturas', '1');
        info.textContent = 'Carregando as páginas…';
        info.hidden = false;
    }
    try {
        const response = await fetch('/pages', { method: 'POST', body: data });
        const body = await response.json();
        if (request !== pageInfoRequest) return;
        info.textContent = response.ok
            ? `Este PDF tem ${body.paginas} página${body.paginas === 1 ? '' : 's'}.`
            : body.erro;
        info.hidden = false;
        if (response.ok && body.miniaturas) {
            pages = body.miniaturas.map((src, i) => ({ pagina: i + 1, giro: 0, src }));
            renderPages();
        }
    } catch {
        // Sem a contagem de páginas a ferramenta continua funcionando
    }
}

// --- Organizar páginas ---

function renderPages() {
    $('organizer').replaceChildren(...pages.map((page, i) => {
        const card = document.createElement('figure');
        card.className = 'page-card';
        card.draggable = true;
        const thumb = document.createElement('div');
        thumb.className = 'page-thumb';
        const img = document.createElement('img');
        img.src = page.src;
        img.alt = `Página ${page.pagina}`;
        img.style.transform = `rotate(${page.giro}deg)`;
        thumb.append(img);
        const caption = document.createElement('figcaption');
        caption.textContent = `Página ${page.pagina}`;
        const actions = document.createElement('div');
        actions.className = 'page-actions';
        const name = `a página ${page.pagina}`;
        actions.append(
            iconButton('←', `Mover ${name} para trás`, i === 0, () => movePage(i, i - 1)),
            iconButton('↺', `Girar ${name} para a esquerda`, false, () => turnPage(i, -90)),
            iconButton('✕', `Excluir ${name}`, false, () => { pages.splice(i, 1); renderPages(); }),
            iconButton('↻', `Girar ${name} para a direita`, false, () => turnPage(i, 90)),
            iconButton('→', `Mover ${name} para a frente`, i === pages.length - 1, () => movePage(i, i + 1)),
        );
        card.addEventListener('dragstart', (event) => {
            dragged = i;
            event.dataTransfer.effectAllowed = 'move';
            card.classList.add('dragging');
        });
        card.addEventListener('dragend', () => card.classList.remove('dragging'));
        card.addEventListener('dragover', (event) => event.preventDefault());
        card.addEventListener('drop', (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (dragged !== null && dragged !== i) movePage(dragged, i);
            dragged = null;
        });
        card.append(thumb, caption, actions);
        return card;
    }));
}

function movePage(from, to) {
    const [page] = pages.splice(from, 1);
    pages.splice(to, 0, page);
    renderPages();
}

function turnPage(i, degrees) {
    pages[i].giro = (pages[i].giro + degrees + 360) % 360;
    renderPages();
}

// --- Envio ---

form.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (!files.length) return;
    if (organizing()) {
        if (!pages.length) {
            setStatus('Espere as páginas carregarem, ou deixe ao menos uma página.', 'error');
            return;
        }
        $('order').value = JSON.stringify(pages.map(({ pagina, giro }) => ({ pagina, giro })));
    }
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
        const message = response.headers.get('X-Mensagem');
        setStatus(`Pronto! O arquivo foi baixado.${message ? ` ${decodeURIComponent(message)}` : ''}`, 'ok');
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
