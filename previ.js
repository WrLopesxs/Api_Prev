
const API_URL = 'https://script.google.com/macros/s/AKfycbzbXBOkRN2YiYremndvUXKXbvMAbqkJRPV8NvJEssARPoScg_bX5ysWbUrgZegEQ1FOGw/exec';
const STORAGE_THEME_KEY = 'dashboard-theme';

let dadosCompletos = [];
let dadosFiltrados = [];
let modalTrendChart = null;
let modoVisualizacao = 'cards';
let ordenacaoTabela = { campo: 'dataFim', direcao: 'asc' };
let autoRefreshTimer = null;
let autoRefreshIntervalSeconds = 0;
let autoRefreshPausado = false;
let autoRefreshStartedAt = null;
let autoRefreshProgressTimer = null;

let filtros = {
    modelo: 'todos',
    turno: 'todos',
    status: 'todos',
    ordem: 'urgencia',
    busca: ''
};

const ultimaAtualizacaoEl = document.getElementById('ultimaAtualizacao');
const cardsContainer = document.getElementById('cardsContainer');
const btnAtualizar = document.getElementById('btnAtualizar');
const btnLimparFiltros = document.getElementById('btnLimparFiltros');
const btnExportarPdf = document.getElementById('btnExportarPdf');
const btnModoVisao = document.getElementById('btnModoVisao');
const btnPausarAutoRefresh = document.getElementById('btnPausarAutoRefresh');
const autoRefreshSelect = document.getElementById('autoRefreshSelect');
const autoRefreshStatus = document.getElementById('autoRefreshStatus');
const autoRefreshCountdown = document.getElementById('autoRefreshCountdown');
const autoRefreshProgressBar = document.getElementById('autoRefreshProgressBar');
const buscaInput = document.getElementById('buscaInput');
const resultadoContadorEl = document.getElementById('resultadoContador');

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

document.addEventListener('DOMContentLoaded', () => {
    inicializarTema();
    configurarEventos();
    carregarDados();
});

function configurarEventos() {
    btnAtualizar.addEventListener('click', carregarDados);

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

    buscaInput.addEventListener('input', (e) => {
        filtros.busca = e.target.value.trim();
        aplicarFiltros();
    });

    btnLimparFiltros.addEventListener('click', limparFiltros);
    btnExportarPdf.addEventListener('click', exportarPDF);
    btnPausarAutoRefresh.addEventListener('click', alternarPausaAutoRefresh);

    autoRefreshSelect.addEventListener('change', (e) => {
        autoRefreshIntervalSeconds = Number(e.target.value) || 0;
        autoRefreshPausado = false;
        atualizarBotaoAutoRefresh();
        configurarAutoRefresh();
    });

    btnModoVisao.addEventListener('click', () => {
        modoVisualizacao = modoVisualizacao === 'cards' ? 'tabela' : 'cards';
        atualizarTextoModoVisao();
        renderizarCards(obterTopUrgentes(dadosFiltrados));
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

    atualizarTextoModoVisao();
    atualizarBotaoAutoRefresh();
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

async function carregarDados() {
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
        html.push(`<button class="filtro-btn ${ativo}" data-modelo="${escapeHtml(modelo)}">${escapeHtml(modelo)}</button>`);
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

    const itemConfig = config[tipo];
    if (itemConfig?.container) {
        const { container, dataAttr } = itemConfig;

        container.querySelectorAll('.filtro-btn').forEach((btn) => {
            btn.classList.toggle('ativo', btn.dataset[dataAttr] === valor);
        });
    }

    aplicarFiltros();
}

function aplicarFiltros() {
    dadosFiltrados = dadosCompletos.filter((item) => itemPassaFiltros(item));
    ordenarDadosFiltrados();

    const topUrgentes = obterTopUrgentes(dadosFiltrados);

    atualizarKPIs();
    atualizarContadorResultados();
    renderizarCards(topUrgentes);
}

function itemPassaFiltros(item, opcoes = {}) {
    const {
        ignorarStatus = false,
        ignorarTurno = false,
        ignorarBusca = false,
        ignorarModelo = false
    } = opcoes;

    if (!itemPodeSerExibido(item)) {
        return false;
    }

    if (!ignorarModelo && filtros.modelo !== 'todos' && item.modelo !== filtros.modelo) {
        return false;
    }

    if (!ignorarTurno && filtros.turno !== 'todos' && String(item.turnoPrevisto) !== String(filtros.turno)) {
        return false;
    }

    const statusInfo = classificarCriticidade(item);
    if (!ignorarStatus && filtros.status !== 'todos' && statusInfo.grupo !== filtros.status) {
        return false;
    }

    if (!ignorarBusca && filtros.busca) {
        const busca = normalizarTexto(filtros.busca);
        const campos = [item.pn, item.modelo, item.linha, statusInfo.label, statusInfo.grupo];
        const encontrou = campos.some((campo) => normalizarTexto(campo).includes(busca));
        if (!encontrou) {
            return false;
        }
    }

    return true;
}

function ordenarDadosFiltrados() {
    dadosFiltrados.sort((a, b) => {
        const statusA = classificarCriticidade(a);
        const statusB = classificarCriticidade(b);
        const dataA = obterTimestampSeguro(a?.dataFim);
        const dataB = obterTimestampSeguro(b?.dataFim);

        if (statusA.grupo === 'sem-consumo' && statusB.grupo !== 'sem-consumo') {
            return 1;
        }

        if (statusA.grupo !== 'sem-consumo' && statusB.grupo === 'sem-consumo') {
            return -1;
        }

        if (filtros.ordem === 'urgencia') {
            const autonomiaA = numeroSeguro(a.autonomiaHoras, Number.POSITIVE_INFINITY);
            const autonomiaB = numeroSeguro(b.autonomiaHoras, Number.POSITIVE_INFINITY);
            if (autonomiaA !== autonomiaB) {
                return autonomiaA - autonomiaB;
            }

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

function obterTopUrgentes(lista) {
    const criticos = lista.filter((item) => classificarCriticidade(item).grupo === 'critico');
    const base = criticos.length > 0 ? criticos : lista;

    const top = [...base]
        .sort((a, b) => numeroSeguro(a.autonomiaHoras, Number.POSITIVE_INFINITY) - numeroSeguro(b.autonomiaHoras, Number.POSITIVE_INFINITY))
        .slice(0, 3);

    const mapa = new Map();
    top.forEach((item, index) => {
        mapa.set(String(item.pn), index + 1);
    });

    return mapa;
}
function atualizarKPIs() {
    const total = dadosFiltrados.length;
    const criticos = dadosFiltrados.filter(item => classificarCriticidade(item).grupo === 'critico').length;
    const atencao = dadosFiltrados.filter(item => classificarCriticidade(item).grupo === 'atencao').length;
    const normal = dadosFiltrados.filter(item => classificarCriticidade(item).grupo === 'normal').length;

    kpiTotal.textContent = total;
    kpiCritico.textContent = criticos;
    kpiAtencao.textContent = atencao;
    kpiNormal.textContent = normal;
}

function atualizarContadorResultados() {
    resultadoContadorEl.textContent = `${dadosFiltrados.length} resultado(s)`;
}
function renderizarCards(topUrgentes) {
    if (dadosFiltrados.length === 0) {
        cardsContainer.classList.remove('modo-tabela');
        cardsContainer.innerHTML = '<div class="loading-cards">Nenhum item encontrado com os filtros atuais.</div>';
        return;
    }

    if (modoVisualizacao === 'tabela') {
        renderizarTabela(topUrgentes);
        return;
    }

    cardsContainer.classList.remove('modo-tabela');
    cardsContainer.innerHTML = dadosFiltrados.map(item => criarCardHTML(item, topUrgentes)).join('');

    cardsContainer.querySelectorAll('.pn-card').forEach((card) => {
        card.addEventListener('click', () => abrirModal(card.dataset.pn));
    });
}

function criarCardHTML(item, topUrgentes) {
    const dataFim = new Date(item.dataFim);
    const autonomiaHoras = numeroSeguro(item.autonomiaHoras);
    const autonomiaDias = (autonomiaHoras / 24).toFixed(1);
    const statusInfo = classificarCriticidade(item);
    const rankUrgencia = topUrgentes.get(String(item.pn));
    const termoBusca = filtros.busca;

    const classes = [
        'pn-card',
        statusInfo.grupo,
        statusInfo.classeFaixa,
        rankUrgencia ? 'urgente' : ''
    ].filter(Boolean).join(' ');

    const badgeUrgente = rankUrgencia
        ? `<span class="urgente-badge">TOP ${rankUrgencia} URGENTE</span>`
        : '';

    return `
        <article class="${classes}" id="card-${criarIdSeguro(item.pn)}" data-pn="${escapeHtml(item.pn)}">
            <div class="pn-topbar">
                <div class="pn-header">
                    <span class="pn-nome">${destacarTexto(item.pn, termoBusca)}</span>
                </div>
                ${badgeUrgente}
                <span class="pn-modelo">${destacarTexto(item.modelo, termoBusca)}</span>
            </div>

            <div class="pn-info">
                <div class="info-item">
                    <div class="info-label">Estoque</div>
                    <div class="info-value">${numeroSeguro(item.estoque)}</div>
                </div>
                <div class="info-item">
                    <div class="info-label">Consumo/h</div>
                    <div class="info-value">${numeroSeguro(item.consumo).toFixed(1)}</div>
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
                    <span class="pn-turno">${escapeHtml(String(item.turnoPrevisto))}º Turno</span>
                </div>
            </div>

            <div class="pn-status ${statusInfo.grupo}" title="${escapeHtml(statusInfo.tooltip)}">
                ${destacarTexto(`${statusInfo.icones} ${statusInfo.label.toUpperCase()}`, termoBusca)}
            </div>
        </article>
    `;
}

function renderizarTabela(topUrgentes) {
    cardsContainer.classList.add('modo-tabela');

    const dadosTabela = [...dadosFiltrados];
    ordenarDadosTabela(dadosTabela);

    const setaCampo = (campo) => {
        if (ordenacaoTabela.campo !== campo) {
            return '';
        }

        return ordenacaoTabela.direcao === 'asc' ? ' ▲' : ' ▼';
    };

    const linhas = dadosTabela.map((item) => {
        const statusInfo = classificarCriticidade(item);
        const rankUrgencia = topUrgentes.get(String(item.pn));
        const termoBusca = filtros.busca;

        return `
            <tr class="linha-tabela ${rankUrgencia ? 'linha-urgente' : ''}" data-pn="${escapeHtml(item.pn)}">
                <td>${rankUrgencia ? `TOP ${rankUrgencia}` : '-'}</td>
                <td>${destacarTexto(item.pn, termoBusca)}</td>
                <td>${destacarTexto(item.modelo, termoBusca)}</td>
                <td>${numeroSeguro(item.estoque)}</td>
                <td>${numeroSeguro(item.consumo).toFixed(1)}</td>
                <td>${numeroSeguro(item.autonomiaHoras).toFixed(1)}h</td>
                <td>${formatarDataHora(new Date(item.dataFim))}</td>
                <td>${escapeHtml(String(item.turnoPrevisto))}º</td>
                <td><span class="status-cell pn-status ${statusInfo.grupo}" title="${escapeHtml(statusInfo.tooltip)}">${destacarTexto(`${statusInfo.icones} ${statusInfo.label}`, termoBusca)}</span></td>
            </tr>
        `;
    }).join('');

    cardsContainer.innerHTML = `
        <div class="tabela-wrapper">
            <table class="tabela-lista">
                <thead>
                    <tr>
                        <th data-sort="urgencia">Urgência${setaCampo('urgencia')}</th>
                        <th data-sort="pn">PN${setaCampo('pn')}</th>
                        <th data-sort="modelo">Modelo${setaCampo('modelo')}</th>
                        <th data-sort="estoque">Estoque${setaCampo('estoque')}</th>
                        <th data-sort="consumo">Consumo/h${setaCampo('consumo')}</th>
                        <th data-sort="autonomiaHoras">Autonomia${setaCampo('autonomiaHoras')}</th>
                        <th data-sort="dataFim">Previsão${setaCampo('dataFim')}</th>
                        <th data-sort="turnoPrevisto">Turno${setaCampo('turnoPrevisto')}</th>
                        <th data-sort="status">Status${setaCampo('status')}</th>
                    </tr>
                </thead>
                <tbody>
                    ${linhas}
                </tbody>
            </table>
        </div>
    `;

    cardsContainer.querySelectorAll('.linha-tabela').forEach((linha) => {
        linha.addEventListener('click', () => abrirModal(linha.dataset.pn));
    });

    cardsContainer.querySelectorAll('th[data-sort]').forEach((th) => {
        th.addEventListener('click', (e) => {
            e.stopPropagation();
            const campo = th.dataset.sort;
            if (ordenacaoTabela.campo === campo) {
                ordenacaoTabela.direcao = ordenacaoTabela.direcao === 'asc' ? 'desc' : 'asc';
            } else {
                ordenacaoTabela.campo = campo;
                ordenacaoTabela.direcao = 'asc';
            }

            renderizarTabela(topUrgentes);
        });
    });
}

function ordenarDadosTabela(lista) {
    const { campo, direcao } = ordenacaoTabela;

    lista.sort((a, b) => {
        let valorA;
        let valorB;

        if (campo === 'status') {
            valorA = classificarCriticidade(a).label;
            valorB = classificarCriticidade(b).label;
        } else if (campo === 'urgencia') {
            valorA = numeroSeguro(a.autonomiaHoras, Number.POSITIVE_INFINITY);
            valorB = numeroSeguro(b.autonomiaHoras, Number.POSITIVE_INFINITY);
        } else if (campo === 'dataFim') {
            valorA = obterTimestampSeguro(a.dataFim) ?? Number.POSITIVE_INFINITY;
            valorB = obterTimestampSeguro(b.dataFim) ?? Number.POSITIVE_INFINITY;
        } else if (['estoque', 'consumo', 'autonomiaHoras', 'turnoPrevisto'].includes(campo)) {
            valorA = numeroSeguro(a[campo], Number.POSITIVE_INFINITY);
            valorB = numeroSeguro(b[campo], Number.POSITIVE_INFINITY);
        } else {
            valorA = normalizarTexto(a[campo]);
            valorB = normalizarTexto(b[campo]);
        }

        if (typeof valorA === 'string' || typeof valorB === 'string') {
            const comparacao = String(valorA).localeCompare(String(valorB), 'pt-BR');
            return direcao === 'asc' ? comparacao : -comparacao;
        }

        const comparacao = valorA - valorB;
        return direcao === 'asc' ? comparacao : -comparacao;
    });
}
function abrirModal(pn) {
    const item = dadosCompletos.find((dado) => String(dado.pn) === String(pn));
    if (!item) {
        return;
    }

    const dataFim = new Date(item.dataFim);
    const autonomiaHoras = numeroSeguro(item.autonomiaHoras);
    const autonomiaDias = (autonomiaHoras / 24).toFixed(1);
    const statusInfo = classificarCriticidade(item);
    const serieEstoque = gerarSerieEstoqueReal(item);

    modalTitulo.textContent = item.pn;

    modalBody.innerHTML = `
        <div class="modal-details">
            <div class="modal-grid">
                <div class="detail-box center">
                    <div class="detail-caption">Modelo</div>
                    <div class="detail-value">${escapeHtml(item.modelo)}</div>
                </div>
                <div class="detail-box center">
                    <div class="detail-caption">Linha</div>
                    <div class="detail-value">${escapeHtml(item.linha || '-')}</div>
                </div>
            </div>

            <div class="detail-section">
                <h4>Detalhes do estoque</h4>
                <div class="detail-row">
                    <span>Quantidade atual:</span>
                    <strong>${numeroSeguro(item.estoque)} peças</strong>
                </div>
                <div class="detail-row">
                    <span>Consumo por hora:</span>
                    <strong>${numeroSeguro(item.consumo).toFixed(1)} peças/h</strong>
                </div>
                <div class="detail-row">
                    <span>Autonomia:</span>
                    <strong>${autonomiaHoras.toFixed(1)} horas (${autonomiaDias} dias)</strong>
                </div>
            </div>

            <div class="trend-chart-wrap">
                <h4>📉 Queda do estoque até o esgotamento</h4>
                <div class="trend-chart-subtitle">Estoque real projetado até ${escapeHtml(formatarDataHora(dataFim))}</div>
                <canvas id="modalTrendChart" aria-label="Gráfico de tendência"></canvas>
                <div class="trend-stats">
                    <div class="trend-stat">
                        <span class="trend-stat-label">Estoque atual</span>
                        <span class="trend-stat-value">${serieEstoque.estoqueInicial.toFixed(0)} peças</span>
                    </div>
                    <div class="trend-stat">
                        <span class="trend-stat-label">Consumo real</span>
                        <span class="trend-stat-value">${serieEstoque.consumoHora.toFixed(2)} p/h</span>
                    </div>
                    <div class="trend-stat">
                        <span class="trend-stat-label">Fim previsto</span>
                        <span class="trend-stat-value">${escapeHtml(serieEstoque.fimPrevisto)}</span>
                    </div>
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
                    <strong>${escapeHtml(String(item.turnoPrevisto))}º Turno</strong>
                </div>
            </div>

            <div class="status-box ${statusInfo.grupo}" title="${escapeHtml(statusInfo.tooltip)}">
                Status: ${escapeHtml(statusInfo.icones)} ${escapeHtml(statusInfo.label.toUpperCase())}
            </div>
        </div>
    `;

    modalDetalhes.classList.add('mostrar');
    renderizarGraficoEstoqueReal(serieEstoque);
}

function renderizarGraficoEstoqueReal(serieEstoque) {
    if (typeof Chart === 'undefined') {
        return;
    }

    const canvas = document.getElementById('modalTrendChart');
    if (!canvas) {
        return;
    }

    if (modalTrendChart) {
        modalTrendChart.destroy();
    }

    const rootStyle = getComputedStyle(document.documentElement);
    const textColor = rootStyle.getPropertyValue('--text').trim() || '#0f2036';

    modalTrendChart = new Chart(canvas, {
        type: 'line',
        data: {
            labels: serieEstoque.labels,
            datasets: [
                {
                    label: 'Estoque restante',
                    data: serieEstoque.valores,
                    borderColor: '#3b9dff',
                    backgroundColor: 'rgba(59, 157, 255, 0.15)',
                    tension: 0,
                    fill: true,
                    pointRadius: 2.2,
                    pointHoverRadius: 4
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: {
                        color: textColor,
                        font: { size: 11, weight: 700 }
                    }
                },
                tooltip: {
                    callbacks: {
                        label: (context) => `Estoque: ${Number(context.parsed.y).toFixed(0)} peças`
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        color: textColor,
                        maxTicksLimit: 8
                    },
                    grid: {
                        color: 'rgba(120, 140, 170, 0.18)'
                    }
                },
                y: {
                    beginAtZero: true,
                    ticks: { color: textColor },
                    grid: {
                        color: 'rgba(120, 140, 170, 0.18)'
                    },
                    title: {
                        display: true,
                        text: 'Peças em estoque',
                        color: textColor
                    }
                }
            }
        }
    });
}

function fecharModal() {
    modalDetalhes.classList.remove('mostrar');
    if (modalTrendChart) {
        modalTrendChart.destroy();
        modalTrendChart = null;
    }
}

function classificarCriticidade(item) {
    const estoque = numeroSeguro(item.estoque, 0);
    const consumo = numeroSeguro(item.consumo, 0);
    const autonomia = numeroSeguro(item.autonomiaHoras, 0);

    if (estoque > 0 && consumo <= 0) {
        return {
            grupo: 'sem-consumo',
            classeFaixa: 'sem-consumo-faixa',
            label: 'Sem consumo',
            icones: '⚪ ⏸',
            tooltip: 'Item com estoque disponível, mas sem consumo registrado'
        };
    }

    if (autonomia <= 0) {
        return {
            grupo: 'esgotado',
            classeFaixa: 'esgotado-faixa',
            label: 'Esgotado',
            icones: '⚫ ❌',
            tooltip: 'Item esgotado'
        };
    }

    if (autonomia < 4) {
        return {
            grupo: 'critico',
            classeFaixa: 'critico-extremo',
            label: 'Crítico <4h',
            icones: '🔴 ⚠️',
            tooltip: 'Nível máximo de urgência (menor que 4h)'
        };
    }

    if (autonomia < 8) {
        return {
            grupo: 'atencao',
            classeFaixa: 'atencao-faixa',
            label: 'Atenção 4-8h',
            icones: '🟡 ⏳',
            tooltip: 'Acompanhar de perto (4h a 8h)'
        };
    }

    return {
        grupo: 'normal',
        classeFaixa: 'normal-faixa',
        label: 'OK >8h',
        icones: '🟢 ✅',
        tooltip: 'Situação estável (acima de 8h)'
    };
}

function gerarSerieConsumo24h(item) {
    return gerarSerieEstoqueReal(item);
}

function gerarSerieEstoqueReal(item) {
    const agora = new Date();
    const dataFim = new Date(item.dataFim);
    const estoqueInicial = numeroSeguro(item.estoque, 0);
    const consumoHora = numeroSeguro(item.consumo, 0);
    const labels = [];
    const valores = [];
    const timestampAgora = agora.getTime();
    const timestampFim = Number.isNaN(dataFim.getTime())
        ? timestampAgora + (numeroSeguro(item.autonomiaHoras, 24) * 3600000)
        : dataFim.getTime();
    const duracaoHoras = Math.max(0.1, (timestampFim - timestampAgora) / 3600000);
    const numPontos = 20;
    const passoHoras = duracaoHoras / (numPontos - 1);

    for (let i = 0; i < numPontos; i += 1) {
        const horasPassadas = i * passoHoras;
        const dataHora = new Date(timestampAgora + horasPassadas * 60 * 60 * 1000);
        const label = duracaoHoras <= 48
            ? `${String(dataHora.getHours()).padStart(2, '0')}:00`
            : `${String(dataHora.getDate()).padStart(2, '0')}/${String(dataHora.getMonth() + 1).padStart(2, '0')} ${String(dataHora.getHours()).padStart(2, '0')}:00`;
        const valor = Math.max(0, estoqueInicial - (consumoHora * horasPassadas));

        labels.push(label);
        valores.push(Number(valor.toFixed(2)));
    }

    return {
        labels,
        valores,
        estoqueInicial,
        consumoHora,
        fimPrevisto: formatarDataHora(new Date(timestampFim))
    };
}
function exportarPDF() {
    if (dadosFiltrados.length === 0) {
        alert('Nenhum dado filtrado para exportar.');
        return;
    }

    const popup = window.open('', '_blank');
    if (!popup) {
        alert('Não foi possível abrir a janela de impressão. Verifique bloqueio de pop-up.');
        return;
    }

    const linhasHtml = dadosFiltrados.map((item) => {
        const statusInfo = classificarCriticidade(item);
        return `
            <tr>
                <td>${escapeHtml(item.pn)}</td>
                <td>${escapeHtml(item.modelo)}</td>
                <td>${numeroSeguro(item.estoque)}</td>
                <td>${numeroSeguro(item.consumo).toFixed(1)}</td>
                <td>${numeroSeguro(item.autonomiaHoras).toFixed(1)}h</td>
                <td>${escapeHtml(formatarDataHora(new Date(item.dataFim)))}</td>
                <td>${escapeHtml(String(item.turnoPrevisto))}º</td>
                <td>${escapeHtml(statusInfo.icones)} ${escapeHtml(statusInfo.label)}</td>
            </tr>
        `;
    }).join('');

    popup.document.write(`
        <html>
        <head>
            <title>Relatório - ${gerarNomeArquivo('previsao')}</title>
            <style>
                body { font-family: Arial, sans-serif; padding: 24px; color: #111; }
                h1 { margin: 0 0 8px; font-size: 20px; }
                .meta { margin-bottom: 16px; color: #444; font-size: 12px; }
                table { border-collapse: collapse; width: 100%; font-size: 12px; }
                th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
                th { background: #f0f0f0; }
            </style>
        </head>
        <body>
            <h1>Dashboard de Previsão - Relatório</h1>
            <div class="meta">Gerado em ${formatarDataHora(new Date())} | ${dadosFiltrados.length} itens</div>
            <table>
                <thead>
                    <tr>
                        <th>PN</th>
                        <th>Modelo</th>
                        <th>Estoque</th>
                        <th>Consumo/h</th>
                        <th>Autonomia</th>
                        <th>Previsão</th>
                        <th>Turno</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>${linhasHtml}</tbody>
            </table>
        </body>
        </html>
    `);

    popup.document.close();
    popup.focus();
    popup.print();
}

function atualizarTextoModoVisao() {
    btnModoVisao.textContent = modoVisualizacao === 'cards' ? 'Visão: Cards' : 'Visão: Tabela';
}

function limparFiltros() {
    filtros = {
        modelo: 'todos',
        turno: 'todos',
        status: 'todos',
        ordem: 'urgencia',
        busca: ''
    };

    buscaInput.value = '';

    [
        [filtroModeloEl, 'modelo', 'todos'],
        [filtroTurnoEl, 'turno', 'todos'],
        [filtroStatusEl, 'status', 'todos'],
        [filtroOrdemEl, 'ordem', 'urgencia']
    ].forEach(([container, dataAttr, valor]) => {
        container.querySelectorAll('.filtro-btn').forEach((btn) => {
            btn.classList.toggle('ativo', btn.dataset[dataAttr] === valor);
        });
    });

    aplicarFiltros();
}

function configurarAutoRefresh() {
    if (autoRefreshTimer) {
        clearInterval(autoRefreshTimer);
        autoRefreshTimer = null;
    }

    if (autoRefreshIntervalSeconds <= 0 || autoRefreshPausado) {
        autoRefreshStartedAt = null;
        iniciarProgressBarAutoRefresh();
        return;
    }

    autoRefreshStartedAt = Date.now();
    autoRefreshTimer = setInterval(() => {
        autoRefreshStartedAt = Date.now();
        carregarDados();
    }, autoRefreshIntervalSeconds * 1000);

    iniciarProgressBarAutoRefresh();
}

function alternarPausaAutoRefresh() {
    if (autoRefreshIntervalSeconds <= 0) {
        return;
    }

    autoRefreshPausado = !autoRefreshPausado;
    atualizarBotaoAutoRefresh();
    configurarAutoRefresh();
}

function atualizarBotaoAutoRefresh() {
    if (autoRefreshIntervalSeconds <= 0) {
        btnPausarAutoRefresh.textContent = 'Auto desligada';
        btnPausarAutoRefresh.disabled = true;
        autoRefreshStatus.textContent = 'Auto atualização desligada';
        autoRefreshCountdown.textContent = '--';
        autoRefreshProgressBar.style.width = '0%';
        autoRefreshProgressBar.classList.add('desligado');
        autoRefreshProgressBar.classList.remove('pausado');
        return;
    }

    btnPausarAutoRefresh.disabled = false;
    btnPausarAutoRefresh.textContent = autoRefreshPausado ? 'Retomar auto' : 'Pausar auto';
    autoRefreshStatus.textContent = autoRefreshPausado
        ? `Auto atualização pausada (${formatarIntervaloAutoRefresh(autoRefreshIntervalSeconds)})`
        : `Auto atualização a cada ${formatarIntervaloAutoRefresh(autoRefreshIntervalSeconds)}`;
    autoRefreshProgressBar.classList.toggle('pausado', autoRefreshPausado);
    autoRefreshProgressBar.classList.remove('desligado');
}

function iniciarProgressBarAutoRefresh() {
    if (autoRefreshProgressTimer) {
        clearInterval(autoRefreshProgressTimer);
        autoRefreshProgressTimer = null;
    }

    atualizarProgressoAutoRefresh();

    if (autoRefreshIntervalSeconds <= 0 || autoRefreshPausado) {
        return;
    }

    autoRefreshProgressTimer = setInterval(atualizarProgressoAutoRefresh, 1000);
}

function atualizarProgressoAutoRefresh() {
    if (autoRefreshIntervalSeconds <= 0) {
        autoRefreshCountdown.textContent = '--';
        autoRefreshProgressBar.style.width = '0%';
        return;
    }

    if (autoRefreshPausado || !autoRefreshStartedAt) {
        autoRefreshCountdown.textContent = 'Pausado';
        return;
    }

    const totalMs = autoRefreshIntervalSeconds * 1000;
    const elapsedMs = Date.now() - autoRefreshStartedAt;
    const restanteMs = Math.max(0, totalMs - elapsedMs);
    const progresso = Math.min(100, (elapsedMs / totalMs) * 100);

    autoRefreshCountdown.textContent = `Próxima em ${formatarTempoRestante(restanteMs)}`;
    autoRefreshProgressBar.style.width = `${progresso}%`;
}

function formatarTempoRestante(restanteMs) {
    const totalSegundos = Math.ceil(restanteMs / 1000);
    const minutos = Math.floor(totalSegundos / 60);
    const segundos = totalSegundos % 60;
    return minutos > 0
        ? `${minutos}:${String(segundos).padStart(2, '0')}`
        : `${segundos}s`;
}

function formatarIntervaloAutoRefresh(segundos) {
    if (segundos % 60 === 0) {
        return `${segundos / 60} min`;
    }

    return `${segundos}s`;
}

function baixarArquivo(blob, nomeArquivo) {
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);

    link.href = url;
    link.download = nomeArquivo;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    URL.revokeObjectURL(url);
}

function gerarNomeArquivo(prefixo) {
    const agora = new Date();
    const yyyy = agora.getFullYear();
    const mm = String(agora.getMonth() + 1).padStart(2, '0');
    const dd = String(agora.getDate()).padStart(2, '0');
    const hh = String(agora.getHours()).padStart(2, '0');
    const mi = String(agora.getMinutes()).padStart(2, '0');
    return `${prefixo}_${yyyy}${mm}${dd}_${hh}${mi}`;
}

function numeroSeguro(valor, fallback = 0) {
    const numero = Number(valor);
    return Number.isFinite(numero) ? numero : fallback;
}

function normalizarTexto(valor) {
    return String(valor || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase();
}

function destaqueSeguro(texto) {
    return escapeHtml(String(texto ?? ''));
}

function destacarTexto(texto, termo) {
    const textoSeguro = destaqueSeguro(texto);
    if (!termo) {
        return textoSeguro;
    }

    const termoLimpo = escapeRegExp(String(termo));
    if (!termoLimpo) {
        return textoSeguro;
    }

    const regex = new RegExp(`(${termoLimpo})`, 'ig');
    return textoSeguro.replace(regex, '<mark>$1</mark>');
}

function escapeRegExp(texto) {
    return texto.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function escapeHtml(valor) {
    return String(valor ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function criarIdSeguro(valor) {
    return String(valor || '')
        .trim()
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '-')
        .replace(/(^-|-$)/g, '');
}

function obterTimestampSeguro(valorData) {
    const timestamp = new Date(valorData).getTime();
    return Number.isFinite(timestamp) ? timestamp : null;
}

function itemPodeSerExibido(item) {
    const estoque = numeroSeguro(item?.estoque, 0);
    const consumo = numeroSeguro(item?.consumo, 0);
    return !(estoque <= 0 && consumo <= 0);
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
        cardsContainer.classList.remove('modo-tabela');
        cardsContainer.innerHTML = '<div class="loading-cards"><span class="loading"></span> Carregando dados da API...</div>';
    } else {
        btnAtualizar.disabled = false;
        btnAtualizar.innerHTML = '<span>↻</span> Atualizar dados';
    }
}

function mostrarErro(mensagem) {
    cardsContainer.classList.remove('modo-tabela');
    cardsContainer.innerHTML = `<div class="loading-cards" style="color: #d42f4a;">${escapeHtml(mensagem)}</div>`;
}
