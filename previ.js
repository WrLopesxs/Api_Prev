// POR ESTA (apenas para teste local):
const API_URL = 'https://script.google.com/macros/s/AKfycbw44AWco0xaVegRU-19DJI8Qd8dijZxJdnzMCfehFuY-yOkTDVFNldZ-RVelDgD1qa-lw/exec';
const STORAGE_THEME_KEY = 'dashboard-theme';
const AUTO_REFRESH_INTERVAL_MS = 3 * 60 * 1000;
const AUTO_REFRESH_TICK_MS = 1000;

let dadosCompletos = [];
let dadosFiltrados = [];
let filtros = {
    modelo: 'todos',
    turno: 'todos',
    status: 'todos',
    ordem: 'desc'
};
let autoRefreshAtivo = true;
let autoRefreshIntervalId = null;
let tempoRestanteAutoMs = AUTO_REFRESH_INTERVAL_MS;
let inicioCicloAutoMs = Date.now();
let carregamentoEmAndamento = false;
let modoVisualizacao = 'cards';
let referenciaCalculoMs = Date.now();

const ultimaAtualizacaoEl = document.getElementById('ultimaAtualizacao');
const cardsContainer = document.getElementById('cardsContainer');
const btnAtualizar = document.getElementById('btnAtualizar');
const filtroModeloEl = document.getElementById('filtroModelo');
const filtroTurnoEl = document.getElementById('filtroTurno');
const filtroStatusEl = document.getElementById('filtroStatus');
const filtroOrdemEl = document.getElementById('filtroOrdem');
const kpiTotal = document.getElementById('kpiTotal');
const kpiCritico = document.getElementById('kpiCritico');
const kpiAtencao = document.getElementById('kpiAtencao');
const kpiNormal = document.getElementById('kpiNormal');
const modalDetalhes = document.getElementById('modalDetalhes');
const modalTitulo = document.getElementById('modalTitulo');
const modalBody = document.getElementById('modalBody');
const btnTema = document.getElementById('btnTema');
const temaIcone = document.getElementById('temaIcone');
const temaLabel = document.getElementById('temaLabel');
const btnPausarAuto = document.getElementById('btnPausarAuto');
const barraAutoRefresh = document.getElementById('barraAutoRefresh');
const tempoRestanteAutoEl = document.getElementById('tempoRestanteAuto');
const statusAutoRefreshEl = document.getElementById('statusAutoRefresh');
const btnAlternarVisualizacao = document.getElementById('btnAlternarVisualizacao');
const tabelaContainer = document.getElementById('tabelaContainer');
const btnImprimirTabela = document.getElementById('btnImprimirTabela');

document.addEventListener('DOMContentLoaded', () => {
    inicializarTema();
    configurarEventos();
    iniciarAutoRefresh();
    carregarDados();
});

function configurarEventos() {
    btnAtualizar.addEventListener('click', atualizarAgora);

    if (btnPausarAuto) {
        btnPausarAuto.addEventListener('click', alternarAutoRefresh);
    }

    if (btnAlternarVisualizacao) {
        btnAlternarVisualizacao.addEventListener('click', alternarVisualizacao);
    }

    if (btnImprimirTabela) {
        btnImprimirTabela.addEventListener('click', imprimirTabelaFiltrada);
    }

    filtroModeloEl.addEventListener('click', (e) => {
        if (e.target.classList.contains('filtro-btn')) {
            atualizarFiltro('modelo', e.target.dataset.modelo);
        }
    });

    filtroTurnoEl.addEventListener('click', (e) => {
        if (e.target.classList.contains('filtro-btn')) {
            atualizarFiltro('turno', e.target.dataset.turno);
        }
    });

    filtroStatusEl.addEventListener('click', (e) => {
        if (e.target.classList.contains('filtro-btn')) {
            atualizarFiltro('status', e.target.dataset.status);
        }
    });

    filtroOrdemEl.addEventListener('click', (e) => {
        if (e.target.classList.contains('filtro-btn')) {
            atualizarFiltro('ordem', e.target.dataset.ordem);
        }
    });

    if (btnTema) {
        btnTema.addEventListener('click', alternarTema);
    }

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && modalDetalhes.classList.contains('mostrar')) {
            fecharModal();
        }
    });

    modalDetalhes.addEventListener('click', (e) => {
        if (e.target === modalDetalhes) {
            fecharModal();
        }
    });

    atualizarBotaoVisualizacao();
}

function inicializarTema() {
    const temaSalvo = localStorage.getItem(STORAGE_THEME_KEY);
    const prefereEscuro = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    const temaInicial = temaSalvo || (prefereEscuro ? 'dark' : 'light');

    aplicarTema(temaInicial, false);
}

function alternarTema() {
    const temaAtual = document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
    const proximoTema = temaAtual === 'dark' ? 'light' : 'dark';

    aplicarTema(proximoTema, true);
}

function aplicarTema(tema, salvar) {
    const temaFinal = tema === 'dark' ? 'dark' : 'light';
    document.documentElement.dataset.theme = temaFinal;

    if (salvar) {
        localStorage.setItem(STORAGE_THEME_KEY, temaFinal);
    }

    atualizarBotaoTema(temaFinal);
}

function atualizarBotaoTema(tema) {
    if (!temaIcone || !temaLabel) {
        return;
    }

    if (tema === 'dark') {
        temaIcone.textContent = '☀';
        temaLabel.textContent = 'Modo claro';
    } else {
        temaIcone.textContent = '🌙';
        temaLabel.textContent = 'Modo escuro';
    }
}

function atualizarAgora() {
    reiniciarContagemAutoRefresh();
    carregarDados();
}

function alternarVisualizacao() {
    modoVisualizacao = modoVisualizacao === 'cards' ? 'tabela' : 'cards';
    atualizarBotaoVisualizacao();
    aplicarFiltros();
}

function atualizarBotaoVisualizacao() {
    if (!btnAlternarVisualizacao) {
        return;
    }

    if (modoVisualizacao === 'cards') {
        btnAlternarVisualizacao.textContent = '📋 Ver tabela';
        return;
    }

    btnAlternarVisualizacao.textContent = '🧩 Ver cards';
}

function iniciarAutoRefresh() {
    reiniciarContagemAutoRefresh();

    if (autoRefreshIntervalId) {
        clearInterval(autoRefreshIntervalId);
    }

    autoRefreshIntervalId = setInterval(processarAutoRefresh, AUTO_REFRESH_TICK_MS);
    atualizarEstadoAutoRefreshUI();
}

function processarAutoRefresh() {
    if (!autoRefreshAtivo) {
        return;
    }

    const decorridoMs = Date.now() - inicioCicloAutoMs;
    tempoRestanteAutoMs = Math.max(AUTO_REFRESH_INTERVAL_MS - decorridoMs, 0);

    if (tempoRestanteAutoMs <= 0) {
        reiniciarContagemAutoRefresh();
        carregarDados();
        return;
    }

    atualizarEstadoAutoRefreshUI();
}

function alternarAutoRefresh() {
    autoRefreshAtivo = !autoRefreshAtivo;

    if (autoRefreshAtivo) {
        if (tempoRestanteAutoMs <= 0 || tempoRestanteAutoMs > AUTO_REFRESH_INTERVAL_MS) {
            tempoRestanteAutoMs = AUTO_REFRESH_INTERVAL_MS;
        }

        inicioCicloAutoMs = Date.now() - (AUTO_REFRESH_INTERVAL_MS - tempoRestanteAutoMs);
    }

    atualizarEstadoAutoRefreshUI();
}

function reiniciarContagemAutoRefresh() {
    inicioCicloAutoMs = Date.now();
    tempoRestanteAutoMs = AUTO_REFRESH_INTERVAL_MS;
    atualizarEstadoAutoRefreshUI();
}

function atualizarEstadoAutoRefreshUI() {
    if (tempoRestanteAutoEl) {
        tempoRestanteAutoEl.textContent = formatarTempoRestante(tempoRestanteAutoMs);
    }

    if (barraAutoRefresh) {
        const porcentagem = Math.max(0, Math.min(100, (tempoRestanteAutoMs / AUTO_REFRESH_INTERVAL_MS) * 100));
        barraAutoRefresh.style.width = `${porcentagem}%`;
        barraAutoRefresh.classList.toggle('pausado', !autoRefreshAtivo);
    }

    if (statusAutoRefreshEl) {
        statusAutoRefreshEl.textContent = autoRefreshAtivo
            ? 'Atualização automática ativa'
            : 'Atualização automática pausada';
    }

    if (btnPausarAuto) {
        btnPausarAuto.classList.toggle('pausado', !autoRefreshAtivo);
        btnPausarAuto.textContent = autoRefreshAtivo
            ? '⏸ Pausar atualização automática'
            : '▶ Retomar atualização automática';
    }
}

function formatarTempoRestante(tempoMs) {
    const totalSegundos = Math.max(0, Math.ceil(tempoMs / 1000));
    const minutos = Math.floor(totalSegundos / 60);
    const segundos = totalSegundos % 60;
    return `${String(minutos).padStart(2, '0')}:${String(segundos).padStart(2, '0')}`;
}

async function carregarDados() {
    if (carregamentoEmAndamento) {
        return;
    }

    carregamentoEmAndamento = true;
    mostrarLoading(true);

    try {
        const response = await fetch(API_URL);

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();

        if (data.status !== 'ok' || !Array.isArray(data.dados)) {
            throw new Error(data.erro || 'Resposta inválida da API');
        }

        dadosCompletos = data.dados.filter((item) => itemPodeSerExibido(item));

        const modelos = [...new Set(dadosCompletos.map(item => item.modelo).filter(Boolean))].sort();
        atualizarFiltrosModelo(modelos);

        if (data.ultimaAtualizacao) {
            const dataAtualizacao = new Date(data.ultimaAtualizacao);
            ultimaAtualizacaoEl.textContent = `Atualizado: ${formatarDataHora(dataAtualizacao)}`;
        } else {
            ultimaAtualizacaoEl.textContent = 'Atualizado agora';
        }

        aplicarFiltros();
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        mostrarErro('Erro de conexão com a API. Tente novamente.');
    } finally {
        mostrarLoading(false);
        carregamentoEmAndamento = false;
    }
}

function atualizarFiltrosModelo(modelos) {
    if (!modelos.includes(filtros.modelo)) {
        filtros.modelo = 'todos';
    }

    const html = [
        `<button class="filtro-btn ${filtros.modelo === 'todos' ? 'ativo' : ''}" data-modelo="todos">Todos</button>`
    ];

    modelos.forEach((modelo) => {
        const ativo = filtros.modelo === modelo ? 'ativo' : '';
        html.push(`<button class="filtro-btn ${ativo}" data-modelo="${modelo}">${modelo}</button>`);
    });

    filtroModeloEl.innerHTML = html.join('');
}

function atualizarFiltro(tipo, valor) {
    filtros[tipo] = valor;

    const config = {
        modelo: { container: filtroModeloEl, dataAttr: 'modelo' },
        turno: { container: filtroTurnoEl, dataAttr: 'turno' },
        status: { container: filtroStatusEl, dataAttr: 'status' },
        ordem: { container: filtroOrdemEl, dataAttr: 'ordem' }
    };

    const { container, dataAttr } = config[tipo];

    container.querySelectorAll('.filtro-btn').forEach((btn) => {
        btn.classList.toggle('ativo', btn.dataset[dataAttr] === valor);
    });

    aplicarFiltros();
}

function aplicarFiltros() {
    referenciaCalculoMs = Date.now();

    dadosFiltrados = dadosCompletos.filter((item) => {
        if (!itemPodeSerExibido(item)) {
            return false;
        }

        if (filtros.modelo !== 'todos' && item.modelo !== filtros.modelo) {
            return false;
        }

        if (filtros.turno !== 'todos' && String(item.turnoPrevisto) !== String(filtros.turno)) {
            return false;
        }

        if (filtros.status !== 'todos' && normalizarStatus(item.status) !== normalizarStatus(filtros.status)) {
            return false;
        }

        return true;
    });

    ordenarDadosFiltrados();
    atualizarKPIs();
    renderizarVisualizacao();
}

function ordenarDadosFiltrados() {
    dadosFiltrados.sort((a, b) => {
        if (modoVisualizacao === 'tabela') {
            const dataA = obterTimestampPrevisao(a);
            const dataB = obterTimestampPrevisao(b);

            if (dataA === null && dataB === null) {
                return 0;
            }

            if (dataA === null) {
                return 1;
            }

            if (dataB === null) {
                return -1;
            }

            return dataA - dataB;
        }

        const dataA = obterTimestampPrevisao(a);
        const dataB = obterTimestampPrevisao(b);

        if (dataA === null && dataB === null) {
            return 0;
        }

        if (dataA === null) {
            return 1;
        }

        if (dataB === null) {
            return -1;
        }

        if (filtros.ordem === 'asc') {
            return dataA - dataB;
        }

        return dataB - dataA;
    });
}

function renderizarVisualizacao() {
    if (modoVisualizacao === 'tabela') {
        cardsContainer.classList.add('hidden');
        tabelaContainer.classList.remove('hidden');
        renderizarTabela();
        return;
    }

    tabelaContainer.classList.add('hidden');
    cardsContainer.classList.remove('hidden');
    renderizarCards();
}

function atualizarKPIs() {
    const total = dadosFiltrados.length;
    const criticos = dadosFiltrados.filter(item => normalizarStatus(item.status) === 'critico').length;
    const atencao = dadosFiltrados.filter(item => normalizarStatus(item.status) === 'atencao').length;
    const normal = dadosFiltrados.filter(item => normalizarStatus(item.status) === 'normal').length;

    kpiTotal.textContent = total;
    kpiCritico.textContent = criticos;
    kpiAtencao.textContent = atencao;
    kpiNormal.textContent = normal;
}

function renderizarCards() {
    if (dadosFiltrados.length === 0) {
        cardsContainer.innerHTML = '<div class="loading-cards">Nenhum item encontrado com os filtros atuais.</div>';
        return;
    }

    cardsContainer.innerHTML = dadosFiltrados.map(item => criarCardHTML(item)).join('');

    document.querySelectorAll('.pn-card').forEach((card) => {
        card.addEventListener('click', () => abrirModal(card.dataset.pn));
    });
}

function renderizarTabela() {
    if (dadosFiltrados.length === 0) {
        tabelaContainer.innerHTML = '<div class="loading-cards">Nenhum item encontrado com os filtros atuais.</div>';
        return;
    }

    const linhas = dadosFiltrados.map((item) => {
        const previsao = obterPrevisaoComDiasUteis(item);
        const dataFim = previsao.dataFim;
        const autonomiaHoras = previsao.autonomiaHoras;
        const statusInfo = obterStatusInfo(item.status);

        return `
            <tr class="linha-tabela" data-pn="${item.pn}">
                <td>${item.pn}</td>
                <td>${item.modelo || '-'}</td>
                <td>${item.estoque}</td>
                <td>${Number(item.consumo || 0).toFixed(1)}</td>
                <td>${autonomiaHoras}h</td>
                <td>${formatarDataHora(dataFim)}</td>
                <td>${item.turnoPrevisto}º</td>
                <td><span class="status-chip ${statusInfo.classe}">${statusInfo.icone} ${statusInfo.label}</span></td>
            </tr>
        `;
    }).join('');

    tabelaContainer.innerHTML = `
        <div class="tabela-scroll">
            <table class="pn-tabela">
                <thead>
                    <tr>
                        <th>PN</th>
                        <th>Modelo</th>
                        <th>Estoque</th>
                        <th>Consumo/h</th>
                        <th>Autonomia</th>
                        <th>Acaba em</th>
                        <th>Turno</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${linhas}
                </tbody>
            </table>
        </div>
    `;

    tabelaContainer.querySelectorAll('.linha-tabela').forEach((linha) => {
        linha.addEventListener('click', () => abrirModal(linha.dataset.pn));
    });
}

function imprimirTabelaFiltrada() {
    if (!Array.isArray(dadosFiltrados) || dadosFiltrados.length === 0) {
        alert('Não há dados para imprimir com os filtros atuais.');
        return;
    }

    const linhas = dadosFiltrados.map((item) => {
        const previsao = obterPrevisaoComDiasUteis(item);
        const dataFim = previsao.dataFim;
        const autonomiaHoras = previsao.autonomiaHoras;
        const statusInfo = obterStatusInfo(item.status);

        return `
            <tr>
                <td>${item.pn}</td>
                <td>${item.modelo || '-'}</td>
                <td>${item.estoque}</td>
                <td>${Number(item.consumo || 0).toFixed(1)}</td>
                <td>${autonomiaHoras}h</td>
                <td>${formatarDataHora(dataFim)}</td>
                <td>${item.turnoPrevisto}º</td>
                <td>${statusInfo.label}</td>
            </tr>
        `;
    }).join('');

    const agora = new Date();
    const popup = window.open('', '_blank');

    if (!popup) {
        alert('Não foi possível abrir a janela de impressão. Verifique o bloqueador de pop-up.');
        return;
    }

    popup.document.write(`
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <title>Tabela - Dashboard Toyota</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 20px;
                    color: #111827;
                }
                h1 {
                    font-size: 20px;
                    margin-bottom: 8px;
                }
                .meta {
                    font-size: 12px;
                    color: #4b5563;
                    margin-bottom: 14px;
                }
                table {
                    width: 100%;
                    border-collapse: collapse;
                    font-size: 12px;
                }
                th, td {
                    border: 1px solid #d1d5db;
                    padding: 8px;
                    text-align: left;
                }
                th {
                    background: #f3f4f6;
                    font-weight: 700;
                }
                @media print {
                    body {
                        margin: 10mm;
                    }
                }
            </style>
        </head>
        <body>
            <h1>Toyota - Tabela de Previsão</h1>
            <div class="meta">Gerado em: ${formatarDataHora(agora)} | Total de itens: ${dadosFiltrados.length}</div>
            <table>
                <thead>
                    <tr>
                        <th>PN</th>
                        <th>Modelo</th>
                        <th>Estoque</th>
                        <th>Consumo/h</th>
                        <th>Autonomia</th>
                        <th>Acaba em</th>
                        <th>Turno</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${linhas}
                </tbody>
            </table>
        </body>
        </html>
    `);

    popup.document.close();
    popup.focus();
    popup.print();
}

function criarCardHTML(item) {
    const previsao = obterPrevisaoComDiasUteis(item);
    const dataFim = previsao.dataFim;
    const autonomiaHoras = previsao.autonomiaHoras;
    const autonomiaDias = (autonomiaHoras / 24).toFixed(1);
    const statusInfo = obterStatusInfo(item.status);

    return `
        <article class="pn-card ${statusInfo.classe}" data-pn="${item.pn}">
            <div class="pn-header">
                <span class="pn-nome">${item.pn}</span>
                <span class="pn-modelo">${item.modelo}</span>
            </div>

            <div class="pn-info">
                <div class="info-item">
                    <div class="info-label">Estoque</div>
                    <div class="info-value">${item.estoque}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Consumo/h</div>
                    <div class="info-value">${Number(item.consumo || 0).toFixed(1)}</div>
                </div>
            </div>

            <div class="pn-previsao">
                <div class="previsao-item">
                    <span class="previsao-label">Autonomia:</span>
                    <span class="previsao-value">${autonomiaHoras}h (${autonomiaDias} dias)</span>
                </div>
                <div class="previsao-item">
                    <span class="previsao-label">Acaba em:</span>
                    <span class="previsao-value">${formatarDataHora(dataFim)}</span>
                </div>
                <div class="previsao-item">
                    <span class="previsao-label">Turno:</span>
                    <span class="pn-turno">${item.turnoPrevisto}º Turno</span>
                </div>
            </div>

            <div class="pn-status ${statusInfo.classe}">
                ${statusInfo.icone} ${statusInfo.label.toUpperCase()} ${statusInfo.icone}
            </div>
        </article>
    `;
}

function abrirModal(pn) {
    const item = dadosCompletos.find((dado) => String(dado.pn) === String(pn));
    if (!item) {
        return;
    }

    const previsao = obterPrevisaoComDiasUteis(item);
    const dataFim = previsao.dataFim;
    const autonomiaHoras = previsao.autonomiaHoras;
    const autonomiaDias = (autonomiaHoras / 24).toFixed(1);
    const statusInfo = obterStatusInfo(item.status);

    modalTitulo.textContent = item.pn;

    modalBody.innerHTML = `
        <div class="modal-details">
            <div class="modal-grid">
                <div class="detail-box center">
                    <div class="detail-caption">Modelo</div>
                    <div class="detail-value">${item.modelo}</div>
                </div>
                <div class="detail-box center">
                    <div class="detail-caption">Linha</div>
                    <div class="detail-value">${item.linha}</div>
                </div>
            </div>

            <div class="detail-section">
                <h4>Detalhes do estoque</h4>
                <div class="detail-row">
                    <span>Quantidade atual:</span>
                    <strong>${item.estoque} peças</strong>
                </div>
                <div class="detail-row">
                    <span>Consumo por hora:</span>
                    <strong>${Number(item.consumo || 0).toFixed(1)} peças/h</strong>
                </div>
                <div class="detail-row">
                    <span>Autonomia:</span>
                    <strong>${autonomiaHoras} horas (${autonomiaDias} dias)</strong>
                </div>
            </div>

            <div class="detail-section">
                <h4>Previsão de esgotamento</h4>
                <div class="detail-row">
                    <span>Data:</span>
                    <strong>${formatarData(dataFim)}</strong>
                </div>
                <div class="detail-row">
                    <span>Horário:</span>
                    <strong>${formatarHora(dataFim)}</strong>
                </div>
                <div class="detail-row">
                    <span>Turno:</span>
                    <strong>${item.turnoPrevisto}º Turno</strong>
                </div>
            </div>

            <div class="status-box ${statusInfo.classe}">
                Status: ${statusInfo.icone} ${statusInfo.label.toUpperCase()}
            </div>
        </div>
    `;

    modalDetalhes.classList.add('mostrar');
}

function fecharModal() {
    modalDetalhes.classList.remove('mostrar');
}

function obterStatusInfo(status) {
    const statusNormalizado = normalizarStatus(status);

    if (statusNormalizado === 'critico') {
        return { classe: 'critico', icone: '🔴', label: 'Crítico' };
    }

    if (statusNormalizado === 'atencao') {
        return { classe: 'atencao', icone: '🟡', label: 'Atenção' };
    }

    return { classe: 'normal', icone: '🟢', label: 'Normal' };
}

function normalizarStatus(status) {
    return String(status || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase();
}

function obterTimestampSeguro(valorData) {
    const timestamp = new Date(valorData).getTime();
    return Number.isFinite(timestamp) ? timestamp : null;
}

function obterTimestampPrevisao(item) {
    const previsao = obterPrevisaoComDiasUteis(item);
    return obterTimestampSeguro(previsao.dataFim);
}

function obterPrevisaoComDiasUteis(item) {
    const autonomiaHoras = obterAutonomiaHorasBase(item);
    const inicioCalculo = new Date(referenciaCalculoMs);
    const dataFim = adicionarHorasConsumindoSomenteDiasUteis(inicioCalculo, autonomiaHoras);

    return { autonomiaHoras, dataFim };
}

function obterAutonomiaHorasBase(item) {
    const autonomiaApi = Number(item?.autonomiaHoras);
    if (Number.isFinite(autonomiaApi) && autonomiaApi > 0) {
        return autonomiaApi;
    }

    const estoque = Number(item?.estoque);
    const consumo = Number(item?.consumo);

    if (Number.isFinite(estoque) && Number.isFinite(consumo) && consumo > 0) {
        return estoque / consumo;
    }

    return 0;
}

function adicionarHorasConsumindoSomenteDiasUteis(dataInicial, horasParaConsumir) {
    const dataBase = new Date(dataInicial);
    if (!Number.isFinite(dataBase.getTime())) {
        return new Date(NaN);
    }

    let atual = new Date(dataBase);
    let horasRestantes = Math.max(Number(horasParaConsumir) || 0, 0);

    while (horasRestantes > 0) {
        if (ehFimDeSemana(atual)) {
            atual = avancarParaProximaSegunda(atual);
            continue;
        }

        const fimDoDia = new Date(atual);
        fimDoDia.setHours(24, 0, 0, 0);
        const horasAteFimDoDia = (fimDoDia.getTime() - atual.getTime()) / (1000 * 60 * 60);
        const horasDoBloco = Math.min(horasRestantes, horasAteFimDoDia);

        atual = new Date(atual.getTime() + (horasDoBloco * 60 * 60 * 1000));
        horasRestantes -= horasDoBloco;
    }

    return atual;
}

function ehFimDeSemana(data) {
    const diaSemana = data.getDay();
    return diaSemana === 0 || diaSemana === 6;
}

function avancarParaProximaSegunda(data) {
    const proximaData = new Date(data);
    const diaSemana = proximaData.getDay();

    if (diaSemana === 6) {
        proximaData.setDate(proximaData.getDate() + 2);
    } else if (diaSemana === 0) {
        proximaData.setDate(proximaData.getDate() + 1);
    }

    proximaData.setHours(0, 0, 0, 0);
    return proximaData;
}

function temEstoqueDisponivel(item) {
    const estoque = Number(item?.estoque);
    return Number.isFinite(estoque) && estoque > 0;
}

function temConsumoValido(item) {
    const consumo = Number(item?.consumo);
    return Number.isFinite(consumo) && consumo > 0;
}

function itemPodeSerExibido(item) {
    return temEstoqueDisponivel(item) && temConsumoValido(item);
}

function formatarDataHora(data) {
    if (!(data instanceof Date) || Number.isNaN(data.getTime())) {
        return '-';
    }

    return data.toLocaleString('pt-BR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function formatarData(data) {
    if (!(data instanceof Date) || Number.isNaN(data.getTime())) {
        return '-';
    }

    return data.toLocaleDateString('pt-BR');
}

function formatarHora(data) {
    if (!(data instanceof Date) || Number.isNaN(data.getTime())) {
        return '-';
    }

    return data.toLocaleTimeString('pt-BR', {
        hour: '2-digit',
        minute: '2-digit'
    });
}

function mostrarLoading(ativo) {
    if (ativo) {
        btnAtualizar.disabled = true;
        btnAtualizar.innerHTML = '<span class="loading"></span> Atualizando...';
        const loadingHtml = '<div class="loading-cards"><span class="loading"></span> Carregando dados da API...</div>';
        cardsContainer.innerHTML = loadingHtml;
        tabelaContainer.innerHTML = loadingHtml;
    } else {
        btnAtualizar.disabled = false;
        btnAtualizar.innerHTML = '<span>↻</span> Atualizar agora';
    }
}

function mostrarErro(mensagem) {
    const erroHtml = `<div class="loading-cards" style="color: #d42f4a;">${mensagem}</div>`;
    cardsContainer.innerHTML = erroHtml;
    tabelaContainer.innerHTML = erroHtml;
}
