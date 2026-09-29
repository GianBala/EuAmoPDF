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
// Sem forçar 'fechado': o evento chega depois do close() e o diálogo pode já ter sido reaberto
dialog.addEventListener('close', () => scheduleEstimate());
// "Outra cor…" mostra o seletor de cor do fundo
$('background').addEventListener('change', () => { $('background-color').hidden = $('background').value !== 'cor'; });
// Mudou alguma opção (e não o arquivo): recalcula a estimativa de tamanho
form.addEventListener('change', (event) => { if (event.target !== input) scheduleEstimate(); });

function openTool(button) {
    tool = button.dataset;
    files = [];
    form.reset();
    $('background-color').hidden = true;  // o reset não dispara o change que o esconde
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
    resetEditor();  // mudaram os arquivos: a prévia antiga não vale mais
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
    scheduleEstimate();
    loadMetadata();
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
    if (previewing() && !edits) {
        await startEditor();
        return;
    }
    const data = formData();
    edits?.forEach((edit) => data.append('mascara', edit.mask, 'mascara.png'));  // na ordem dos arquivos

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

function formData() {
    const data = new FormData(form);
    data.append('action', tool.action);
    files.forEach((file) => data.append('file', file));
    return data;
}

// --- Estimativa de tamanho (ferramentas de compressão) ---

let estimateTimer = null;
let estimateRequest = null;  // AbortController da estimativa em andamento

function scheduleEstimate(active = dialog.open) {
    clearTimeout(estimateTimer);
    estimateRequest?.abort();
    const box = $('estimate');
    if (!active || !tool.fields.split(' ').includes('estimate') || !files.length) {
        box.textContent = '';
        return;
    }
    box.textContent = 'Calculando o tamanho final…';
    box.classList.add('calculating');
    estimateTimer = setTimeout(requestEstimate, 400);  // espera o usuário terminar de mexer nas opções
}

async function requestEstimate() {
    const box = $('estimate');
    const controller = new AbortController();
    estimateRequest = controller;
    try {
        const response = await fetch('/estimate', { method: 'POST', body: formData(), signal: controller.signal });
        if (!response.ok) {
            box.textContent = await response.text();  // ex.: PDF com senha, arquivo inválido
        } else {
            const { antes, depois } = await response.json();
            const percent = Math.round((1 - depois / antes) * 100);
            box.textContent = percent > 0
                ? `Tamanho estimado: ~${formatSize(depois)} (hoje ${formatSize(antes)}, ${percent}% menor)`
                : `Tamanho estimado: ~${formatSize(antes)}. Com estas opções, o arquivo praticamente não diminui.`;
        }
        box.classList.remove('calculating');
    } catch (error) {
        if (error.name !== 'AbortError') box.textContent = '';
    }
}

// --- Editar metadados ---

let metadataRequest = 0;

// Mostra os campos do tipo de arquivo (pdf ou imagem); os outros ficam desativados e não são enviados
function showMetadataKind(kind, info = {}) {
    const strip = $('meta-strip').checked;
    document.querySelectorAll('.meta-grid').forEach((grid) => {
        grid.hidden = grid.dataset.kind !== kind;
        grid.querySelectorAll('input').forEach((input) => { input.disabled = grid.hidden || strip; });
    });
    const details = [info.camera && `Câmera: ${info.camera}`, info.localizacao && `Localização: ${info.localizacao}`].filter(Boolean);
    $('meta-info').textContent = details.join(' · ');
    $('meta-info').hidden = !details.length;
    $('meta-gps').hidden = !info.localizacao;
    $('meta-gps').querySelector('input').disabled = !info.localizacao || strip;
}

async function loadMetadata() {
    if (!tool.fields.split(' ').includes('metadata')) return;
    const request = ++metadataRequest;  // ignora a resposta de um arquivo que já foi trocado
    document.querySelectorAll('.meta-grid input').forEach((input) => { input.value = ''; });
    showMetadataKind(null);
    if (files.length !== 1) return;
    const data = new FormData();
    data.append('file', files[0]);
    try {
        const response = await fetch('/metadata', { method: 'POST', body: data });
        if (request !== metadataRequest) return;
        if (!response.ok) {
            setStatus(await response.text(), 'error');
            return;
        }
        const values = await response.json();
        for (const [name, value] of Object.entries(values)) {
            if (form.elements[name]) form.elements[name].value = value;
        }
        showMetadataKind(values.tipo, values);
    } catch {
        // sem os valores atuais, os campos ficam em branco e ainda dá para preenchê-los
    }
}

// "Remover todos" desativa os campos, que então não são enviados
$('meta-strip').addEventListener('change', () => {
    document.querySelectorAll('.meta-grid:not([hidden]) input, #meta-gps:not([hidden]) input')
        .forEach((input) => { input.disabled = $('meta-strip').checked; });
});

// --- Remover fundo: prévia editável ---

let edits = null;           // depois da prévia: [{file, mask: PNG da máscara, history: [máscaras anteriores]}]
let current = 0;            // imagem aberta no editor
let brushMode = 'restaurar';
let bitmap = null;          // imagem aberta, decodificada
let painting = null;        // último ponto do traço em andamento
const paint = $('paint');

function previewing() {
    return tool.fields.split(' ').includes('preview');
}

function resetEditor() {
    edits = null;
    $('editor').hidden = true;
    form.classList.remove('editing');
    dialog.classList.remove('wide');
    selectMode('restaurar');
    $('view-final').setAttribute('aria-pressed', 'false');
}

async function startEditor() {
    setBusy(true);
    try {
        const masks = [];
        for (const [i, file] of files.entries()) {
            setStatus(files.length > 1 ? `Removendo o fundo… (${i + 1} de ${files.length})` : 'Removendo o fundo…');
            const data = new FormData();
            data.append('action', tool.action);
            data.append('file', file);
            const response = await fetch('/background/mask', { method: 'POST', body: data });
            if (!response.ok) throw new Error(await response.text());
            masks.push({ file, mask: await response.blob(), history: [] });
        }
        edits = masks;
        setStatus('');
        await openEditor(0);
    } catch (error) {
        setStatus(error instanceof TypeError ? 'Não foi possível falar com o EuAmoPDF. Ele ainda está aberto?' : error.message, 'error');
    } finally {
        setBusy(false);
    }
}

async function openEditor(index) {
    current = index;
    form.classList.add('editing');
    dialog.classList.add('wide');
    $('editor').hidden = false;
    $('editor-nav').hidden = edits.length < 2;
    $('editor-count').textContent = `Imagem ${current + 1} de ${edits.length}`;
    $('prev').disabled = current === 0;
    $('next').disabled = current === edits.length - 1;
    $('editor-status').textContent = '';
    bitmap = await createImageBitmap(edits[current].file, { imageOrientation: 'from-image' });
    // Cabe na janela sem ampliar além do original; o canvas acompanha a densidade da tela
    const fit = Math.min(1, $('editor').clientWidth / bitmap.width, Math.max(240, window.innerHeight * 0.55) / bitmap.height);
    const width = Math.round(bitmap.width * fit);
    const height = Math.round(bitmap.height * fit);
    const dpr = window.devicePixelRatio || 1;
    for (const canvas of [$('preview'), paint]) {
        canvas.width = Math.round(width * dpr);
        canvas.height = Math.round(height * dpr);
        canvas.style.width = `${width}px`;
        canvas.style.height = `${height}px`;
    }
    $('stage').style.width = `${width}px`;
    showBackground();
    await drawPreview();
}

async function drawPreview() {
    const edit = edits[current];
    const canvas = $('preview');
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    const { width, height } = canvas;
    ctx.clearRect(0, 0, width, height);
    ctx.drawImage(await createImageBitmap(edit.mask), 0, 0, width, height);
    const mask = ctx.getImageData(0, 0, width, height).data;
    ctx.clearRect(0, 0, width, height);
    ctx.drawImage(bitmap, 0, 0, width, height);
    const pixels = ctx.getImageData(0, 0, width, height);
    const data = pixels.data;
    const final = $('view-final').getAttribute('aria-pressed') === 'true';
    for (let i = 0; i < data.length; i += 4) {
        const removed = 1 - mask[i] / 255;
        if (final) {  // como vai ficar: o que saiu some
            data[i + 3] *= 1 - removed;
            continue;
        }
        // O que saiu continua bem visível (80% opaco), com um véu vermelho que o separa do que fica
        data[i] += (220 - data[i]) * removed * 0.4;
        data[i + 1] += (40 - data[i + 1]) * removed * 0.4;
        data[i + 2] += (40 - data[i + 2]) * removed * 0.4;
        data[i + 3] *= 1 - removed * 0.2;
    }
    ctx.putImageData(pixels, 0, 0);
    $('undo').disabled = !edit.history.length;
}

function showBackground() {
    const color = { branco: '#ffffff', preto: '#000000', cor: $('background-color').value }[$('background').value];
    $('stage').style.background = color || '';  // sem cor: o xadrez de transparência do CSS
}

function selectMode(mode) {
    brushMode = mode;
    document.querySelectorAll('.segmented .mode').forEach((button) => {
        button.setAttribute('aria-pressed', String(button.dataset.mode === mode));
    });
}

function pointOf(event) {
    const box = paint.getBoundingClientRect();
    return { x: (event.clientX - box.left) * paint.width / box.width, y: (event.clientY - box.top) * paint.height / box.height };
}

function strokeTo(point) {
    const ctx = paint.getContext('2d');
    ctx.strokeStyle = brushMode === 'restaurar' ? '#1ec85a' : '#e63232';
    ctx.lineWidth = $('brush').value * (window.devicePixelRatio || 1);
    ctx.lineCap = ctx.lineJoin = 'round';
    ctx.beginPath();
    ctx.moveTo(painting.x, painting.y);
    ctx.lineTo(point.x, point.y);
    ctx.stroke();
    painting = point;
}

async function finishStroke() {
    if (!painting) return;
    painting = null;
    const edit = edits[current];
    const data = new FormData();
    data.append('action', tool.action);
    data.append('modo', brushMode);
    data.append('inteligente', $('smart').checked ? '1' : '0');
    data.append('file', edit.file);
    data.append('mascara', edit.mask, 'mascara.png');
    data.append('traco', await new Promise((resolve) => paint.toBlob(resolve, 'image/png')), 'traco.png');
    $('stage').classList.add('working');
    $('editor-status').textContent = !$('smart').checked ? 'Aplicando o pincel…'
        : brushMode === 'restaurar' ? 'Restaurando a área marcada…' : 'Apagando a área marcada…';
    try {
        const response = await fetch('/background/refine', { method: 'POST', body: data });
        if (!response.ok) throw new Error(await response.text());
        edit.history.push(edit.mask);
        edit.mask = await response.blob();
        if (edits?.[current] === edit) await drawPreview();  // o usuário pode ter trocado de imagem
        $('editor-status').textContent = '';
    } catch (error) {
        $('editor-status').textContent = error instanceof TypeError ? 'Não foi possível falar com o EuAmoPDF.' : error.message;
    } finally {
        paint.getContext('2d').clearRect(0, 0, paint.width, paint.height);
        $('stage').classList.remove('working');
    }
}

async function undo() {
    const edit = edits?.[current];
    if (!edit?.history.length || $('stage').classList.contains('working')) return;
    edit.mask = edit.history.pop();
    await drawPreview();
}

paint.addEventListener('pointerdown', (event) => {
    if ($('stage').classList.contains('working')) return;
    try { paint.setPointerCapture(event.pointerId); } catch { /* sem captura, o traço só para ao sair do canvas */ }
    painting = pointOf(event);
    strokeTo(painting);  // um clique sem arrastar também marca
});
paint.addEventListener('pointermove', (event) => { if (painting) strokeTo(pointOf(event)); });
paint.addEventListener('pointerup', finishStroke);
paint.addEventListener('pointercancel', () => {
    painting = null;
    paint.getContext('2d').clearRect(0, 0, paint.width, paint.height);
});
document.querySelectorAll('.segmented .mode').forEach((button) => button.addEventListener('click', () => selectMode(button.dataset.mode)));
$('undo').addEventListener('click', undo);
$('view-final').addEventListener('click', () => {
    const button = $('view-final');
    button.setAttribute('aria-pressed', String(button.getAttribute('aria-pressed') !== 'true'));
    drawPreview();
});
document.addEventListener('keydown', (event) => {
    if (edits && dialog.open && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') {
        event.preventDefault();
        undo();
    }
});
$('prev').addEventListener('click', () => openEditor(current - 1));
$('next').addEventListener('click', () => openEditor(current + 1));
$('back').addEventListener('click', () => { resetEditor(); setBusy(false); });
$('background').addEventListener('change', () => { if (edits) showBackground(); });
$('background-color').addEventListener('input', () => { if (edits) showBackground(); });

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
    submit.textContent = busy ? 'Processando…' : edits ? 'Baixar' : tool.title;
}

function setStatus(message, kind) {
    const status = $('status');
    status.textContent = message;
    status.className = `status ${kind || ''}`;
    status.hidden = !message;
}
