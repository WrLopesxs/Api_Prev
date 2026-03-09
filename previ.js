
const API_URL = 'https://script.google.com/macros/s/AKfycbzbXBOkRN2YiYremndvUXKXbvMAbqkJRPV8NvJEssARPoScg_bX5ysWbUrgZegEQ1FOGw/exec';
const STORAGE_THEME_KEY = 'dashboard-theme';
const STORAGE_ECONOMICO_KEY = 'dashboard-economico';
const DEBOUNCE_BUSCA_MS = 280;
const CARDS_POR_LOTE = 24;
const NUM_PONTOS_GRAFICO = 12;
const LIMIAR_PERFORMANCE_VISUAL = 30;
const LIMITE_PADRAO_GRAFICO_CONSUMO = '10';
const LIMITE_SEGURANCA_GRAFICO_CONSUMO = 40;

let dadosCompletos = [];
let dadosFiltrados = [];
let graficoConsumoChart = null;
let graficoConsumoFingerprint = '';
let modoVisualizacao = 'cards';
let ordenacaoTabela = { campo: 'dataFim', direcao: 'asc' };
let autoRefreshTimer = null;
let autoRefreshIntervalSeconds = 0;
let autoRefreshPausado = false;
let autoRefreshStartedAt = null;
let autoRefreshProgressTimer = null;
let filtroModeloGraficoConsumo = 'todos';
let limitePnGraficoConsumo = 'todos';
let pnSelecionadoGraficoConsumo = 'todos';
let handlerHoverGraficoConsumo = null;
let ultimoHoverTooltipConsumo = '';
let buscaDebounceTimer = null;
let cardsVisiveis = [];
let cursorCardsVisiveis = 0;
let observerLazyCards = null;
let ultimoFingerprintDados = '';
let ultimoFingerprintModelos = '';
let carregamentoEmAndamento = false;
let limiteSegurancaGraficoAplicado = false;
let totalItensGraficoAntesLimite = 0;

let filtros = {
    modelo: 'todos',
    turno: 'todos',
    status: 'todos',
    ordem: 'urgencia',
    limiteCards: 'todos',
    busca: ''
};

const ultimaAtualizacaoEl = document.getElementById('ultimaAtualizacao');
const cardsContainer = document.getElementById('cardsContainer');
const btnAtualizar = document.getElementById('btnAtualizar');
const btnLimparFiltros = document.getElementById('btnLimparFiltros');
const btnGerarGraficoConsumo = document.getElementById('btnGerarGraficoConsumo');
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
const filtroLimiteCardsEl = document.getElementById('filtroLimiteCards');

const kpiTotal = document.getElementById('kpiTotal');
const kpiCritico = document.getElementById('kpiCritico');
const kpiAtencao = document.getElementById('kpiAtencao');
const kpiNormal = document.getElementById('kpiNormal');

const modalDetalhes = document.getElementById('modalDetalhes');
const modalTitulo = document.getElementById('modalTitulo');
const modalBody = document.getElementById('modalBody');
const modalContent = modalDetalhes?.querySelector('.modal-content');

const btnTema = document.getElementById('btnTema');
const temaIcone = document.getElementById('temaIcone');
const temaLabel = document.getElementById('temaLabel');
const btnModoEconomico = document.getElementById('btnModoEconomico');

document.addEventListener('DOMContentLoaded', () => {
    inicializarTema();
    inicializarModoEconomico();
    configurarEventos();
    carregarDados({ silencioso: false });
});

function configurarEventos() {
    btnAtualizar.addEventListener('click', () => carregarDados({ silencioso: false }));

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

    filtroLimiteCardsEl?.addEventListener('click', (e) => {
        if (e.target.classList.contains('filtro-btn')) {
            atualizarFiltro('limiteCards', e.target.dataset.limiteCards);
        }
    });

    buscaInput.addEventListener('input', (e) => {
        if (buscaDebounceTimer) {
            clearTimeout(buscaDebounceTimer);
        }

        const termo = e.target.value.trim();
        buscaDebounceTimer = setTimeout(() => {
            filtros.busca = termo;
            aplicarFiltros();
        }, DEBOUNCE_BUSCA_MS);
    });

    btnLimparFiltros.addEventListener('click', limparFiltros);
    btnGerarGraficoConsumo.addEventListener('click', abrirModalGraficoConsumo);
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
        renderizarCards();
    });

    if (btnTema) {
        btnTema.addEventListener('click', alternarTema);
    }
    if (btnModoEconomico) {
        btnModoEconomico.addEventListener('click', alternarModoEconomico);
    }

    cardsContainer.addEventListener('click', (e) => {
        const card = e.target.closest('.pn-card');
        if (!card || !cardsContainer.contains(card)) {
            return;
        }

        abrirModal(card.dataset.pn);
    });

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

function inicializarModoEconomico() {
    const economicoSalvo = localStorage.getItem(STORAGE_ECONOMICO_KEY);
    const ativo = economicoSalvo === '1';
    aplicarModoEconomico(ativo, false);
}

function alternarModoEconomico() {
    const atual = document.documentElement.dataset.economico === 'true';
    aplicarModoEconomico(!atual, true);
}

function aplicarModoEconomico(ativo, salvar) {
    document.documentElement.dataset.economico = ativo ? 'true' : 'false';

    if (salvar) {
        localStorage.setItem(STORAGE_ECONOMICO_KEY, ativo ? '1' : '0');
    }

    if (btnModoEconomico) {
        btnModoEconomico.textContent = ativo ? 'Modo econômico: ligado' : 'Modo econômico: desligado';
        btnModoEconomico.setAttribute('aria-pressed', ativo ? 'true' : 'false');
    }

    atualizarModoPerformanceAutomatico(modoVisualizacao === 'cards' ? cardsVisiveis.length : dadosFiltrados.length);
}

async function carregarDados(opcoes = {}) {
    const { silencioso = false } = opcoes;
    if (carregamentoEmAndamento) {
        return;
    }

    carregamentoEmAndamento = true;
    const limparCardsDuranteLoading = dadosCompletos.length === 0;

    if (!silencioso) {
        mostrarLoading(true, limparCardsDuranteLoading);
    }

    try {
        const response = await fetch(API_URL);

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();

        if (data.status !== 'ok' || !Array.isArray(data.dados)) {
            throw new Error(data.erro || 'Resposta inválida da API');
        }

        const novosDados = data.dados.filter((item) => itemPodeSerExibido(item));
        const novoFingerprint = gerarFingerprintDados(novosDados);
        const dadosMudaram = novoFingerprint !== ultimoFingerprintDados;

        if (!dadosMudaram) {
            atualizarTextoUltimaAtualizacao(data.ultimaAtualizacao, true);
            return;
        }

        dadosCompletos = novosDados;
        ultimoFingerprintDados = novoFingerprint;

        const modelos = [...new Set(dadosCompletos.map((item) => item.modelo).filter(Boolean))].sort();
        atualizarFiltrosModelo(modelos);
        atualizarTextoUltimaAtualizacao(data.ultimaAtualizacao, false);
        aplicarFiltros();
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        if (silencioso) {
            ultimaAtualizacaoEl.textContent = `Falha na checagem automática (${formatarHora(new Date())})`;
        } else {
            mostrarErro('Erro de conexão com a API. Tente novamente.');
        }
    } finally {
        carregamentoEmAndamento = false;
        if (!silencioso) {
            mostrarLoading(false);
        }
    }
}

function atualizarTextoUltimaAtualizacao(ultimaAtualizacaoApi, semMudanca) {
    if (ultimaAtualizacaoApi) {
        const dataAtualizacao = new Date(ultimaAtualizacaoApi);
        const prefixo = semMudanca ? 'Sem mudanças (checado em):' : 'Atualizado:';
        ultimaAtualizacaoEl.textContent = `${prefixo} ${formatarDataHora(dataAtualizacao)}`;
        return;
    }

    ultimaAtualizacaoEl.textContent = semMudanca ? 'Sem mudanças detectadas' : 'Atualizado agora';
}

function atualizarFiltrosModelo(modelos) {
    if (!modelos.includes(filtros.modelo)) {
        filtros.modelo = 'todos';
    }

    const fingerprintModelos = gerarFingerprintListaModelos(modelos);
    if (fingerprintModelos === ultimoFingerprintModelos) {
        return;
    }

    const html = [
        `<button class="filtro-btn ${filtros.modelo === 'todos' ? 'ativo' : ''}" data-modelo="todos">Todos</button>`
    ];

    modelos.forEach((modelo) => {
        const ativo = filtros.modelo === modelo ? 'ativo' : '';
        html.push(`<button class="filtro-btn ${ativo}" data-modelo="${escapeHtml(modelo)}">${escapeHtml(modelo)}</button>`);
    });

    filtroModeloEl.innerHTML = html.join('');
    ultimoFingerprintModelos = fingerprintModelos;
}

function atualizarFiltro(tipo, valor) {
    filtros[tipo] = valor;

    const config = {
        modelo: { container: filtroModeloEl, dataAttr: 'modelo' },
        turno: { container: filtroTurnoEl, dataAttr: 'turno' },
        status: { container: filtroStatusEl, dataAttr: 'status' },
        ordem: { container: filtroOrdemEl, dataAttr: 'ordem' },
        limiteCards: { container: filtroLimiteCardsEl, dataAttr: 'limiteCards' }
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

    atualizarKPIs();
    atualizarContadorResultados();
    renderizarCards();
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
        const dataA = obterTimestampEsgotamento(a);
        const dataB = obterTimestampEsgotamento(b);

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
function renderizarCards() {
    if (dadosFiltrados.length === 0) {
        desconectarObserverLazyCards();
        cardsVisiveis = [];
        cursorCardsVisiveis = 0;
        atualizarModoPerformanceAutomatico(0);
        cardsContainer.classList.remove('modo-tabela');
        cardsContainer.innerHTML = '<div class="loading-cards">Nenhum item encontrado com os filtros atuais.</div>';
        return;
    }

    if (modoVisualizacao === 'tabela') {
        desconectarObserverLazyCards();
        atualizarModoPerformanceAutomatico(dadosFiltrados.length);
        renderizarTabela();
        return;
    }

    cardsContainer.classList.remove('modo-tabela');
    cardsVisiveis = obterListaLimitadaCards(dadosFiltrados);
    atualizarModoPerformanceAutomatico(cardsVisiveis.length);
    cursorCardsVisiveis = 0;
    cardsContainer.innerHTML = '';
    carregarProximoLoteCards();
}

function obterListaLimitadaCards(lista) {
    if (filtros.limiteCards === 'todos') {
        return [...lista];
    }

    const limite = Number(filtros.limiteCards);
    if (!Number.isFinite(limite) || limite <= 0) {
        return [...lista];
    }

    return lista.slice(0, limite);
}

function atualizarModoPerformanceAutomatico(totalItens) {
    const root = document.documentElement;
    const modoEconomicoAtivo = root.dataset.economico === 'true';

    if (modoEconomicoAtivo) {
        root.dataset.performance = 'low';
        return;
    }

    root.dataset.performance = totalItens >= LIMIAR_PERFORMANCE_VISUAL ? 'low' : 'normal';
}

function carregarProximoLoteCards() {
    if (cursorCardsVisiveis >= cardsVisiveis.length) {
        atualizarSentinelaLazyCards();
        return;
    }

    cardsContainer.querySelector('.cards-sentinel')?.remove();

    const fimLote = Math.min(cursorCardsVisiveis + CARDS_POR_LOTE, cardsVisiveis.length);
    const lote = cardsVisiveis.slice(cursorCardsVisiveis, fimLote);
    const htmlLote = lote.map((item) => criarCardHTML(item)).join('');

    cardsContainer.insertAdjacentHTML('beforeend', htmlLote);
    cursorCardsVisiveis = fimLote;
    atualizarSentinelaLazyCards();
}

function atualizarSentinelaLazyCards() {
    if (cursorCardsVisiveis >= cardsVisiveis.length) {
        desconectarObserverLazyCards();
        cardsContainer.querySelector('.cards-sentinel')?.remove();
        return;
    }

    let sentinela = cardsContainer.querySelector('.cards-sentinel');
    if (!sentinela) {
        sentinela = document.createElement('div');
        sentinela.className = 'cards-sentinel';
    }

    sentinela.textContent = `Role para carregar mais (${cursorCardsVisiveis}/${cardsVisiveis.length})`;
    cardsContainer.appendChild(sentinela);
    observarSentinelaCards(sentinela);
}

function observarSentinelaCards(sentinela) {
    if (typeof IntersectionObserver === 'undefined') {
        cardsContainer.querySelector('.cards-sentinel')?.remove();

        while (cursorCardsVisiveis < cardsVisiveis.length) {
            const fimLote = Math.min(cursorCardsVisiveis + CARDS_POR_LOTE, cardsVisiveis.length);
            const lote = cardsVisiveis.slice(cursorCardsVisiveis, fimLote);
            const htmlLote = lote.map((item) => criarCardHTML(item)).join('');
            cardsContainer.insertAdjacentHTML('beforeend', htmlLote);
            cursorCardsVisiveis = fimLote;
        }

        return;
    }

    if (!observerLazyCards) {
        observerLazyCards = new IntersectionObserver((entries) => {
            const entrouNaTela = entries.some((entry) => entry.isIntersecting);
            if (entrouNaTela) {
                carregarProximoLoteCards();
            }
        }, {
            root: null,
            rootMargin: '320px 0px',
            threshold: 0.01
        });
    }

    observerLazyCards.disconnect();
    observerLazyCards.observe(sentinela);
}

function desconectarObserverLazyCards() {
    if (observerLazyCards) {
        observerLazyCards.disconnect();
    }
}

function criarCardHTML(item) {
    const dataFim = new Date(item.dataFim);
    const autonomiaHoras = numeroSeguro(item.autonomiaHoras);
    const autonomiaDias = (autonomiaHoras / 24).toFixed(1);
    const statusInfo = classificarCriticidade(item);
    const termoBusca = filtros.busca;

    const classes = [
        'pn-card',
        statusInfo.grupo,
        statusInfo.classeFaixa
    ].filter(Boolean).join(' ');

    return `
        <article class="${classes}" id="card-${criarIdSeguro(item.pn)}" data-pn="${escapeHtml(item.pn)}">
            <div class="pn-topbar">
                <div class="pn-header">
                    <span class="pn-nome">${destacarTexto(item.pn, termoBusca)}</span>
                </div>
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

function renderizarTabela() {
    cardsContainer.classList.add('modo-tabela');

    const dadosTabela = dadosFiltrados.filter((item) => {
        const status = classificarCriticidade(item).grupo;
        return status !== 'sem-consumo' && status !== 'esgotado';
    });

    if (dadosTabela.length === 0) {
        cardsContainer.innerHTML = '<div class="loading-cards">Nenhum item com consumo ativo para exibir na tabela.</div>';
        return;
    }

    ordenarDadosTabela(dadosTabela);

    const setaCampo = (campo) => {
        if (ordenacaoTabela.campo !== campo) {
            return '';
        }

        return ordenacaoTabela.direcao === 'asc' ? ' ▲' : ' ▼';
    };

    const linhas = dadosTabela.map((item) => {
        const statusInfo = classificarCriticidade(item);
        const termoBusca = filtros.busca;

        return `
            <tr class="linha-tabela" data-pn="${escapeHtml(item.pn)}">
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

            renderizarTabela();
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
            valorA = obterTimestampEsgotamento(a) ?? Number.POSITIVE_INFINITY;
            valorB = obterTimestampEsgotamento(b) ?? Number.POSITIVE_INFINITY;
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

    destruirGraficosModal();
    modalContent?.classList.remove('modal-consumo');
    const dataFim = new Date(item.dataFim);
    const autonomiaHoras = numeroSeguro(item.autonomiaHoras);
    const autonomiaDias = (autonomiaHoras / 24).toFixed(1);
    const statusInfo = classificarCriticidade(item);
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
}

function abrirModalGraficoConsumo() {
    const itensDisponiveis = dadosFiltrados.filter((item) => itemPodeAparecerNoGraficoConsumo(item));
    if (itensDisponiveis.length === 0) {
        alert('Nenhum part number com estoque e consumo por hora disponível para gerar o gráfico.');
        return;
    }

    destruirGraficosModal();
    filtroModeloGraficoConsumo = filtros.modelo !== 'todos' ? filtros.modelo : 'todos';
    limitePnGraficoConsumo = LIMITE_PADRAO_GRAFICO_CONSUMO;
    pnSelecionadoGraficoConsumo = 'todos';
    modalContent?.classList.add('modal-consumo');
    modalTitulo.textContent = 'Gráfico de consumo por part number';
    modalBody.innerHTML = criarConteudoModalGraficoConsumo();
    modalDetalhes.classList.add('mostrar');
    configurarFiltroModalGraficoConsumo();
    renderizarGraficoConsumo();
}

function criarConteudoModalGraficoConsumo() {
    const modelos = obterModelosGraficoConsumo();
    const botoesModelo = [
        `<button class="filtro-btn ${filtroModeloGraficoConsumo === 'todos' ? 'ativo' : ''}" data-modelo-grafico="todos">Todos</button>`,
        ...modelos.map((modelo) => `
            <button class="filtro-btn ${filtroModeloGraficoConsumo === modelo ? 'ativo' : ''}" data-modelo-grafico="${escapeHtml(modelo)}">
                ${escapeHtml(modelo)}
            </button>
        `)
    ].join('');

    return `
        <div class="modal-details consumo-modal">
            <div class="detail-section">
                <h4>Filtro em cascata por modelo</h4>
                <div class="trend-chart-subtitle">Selecione o modelo para reorganizar o consumo por part number.</div>
                <div class="filtro-buttons consumo-cascata-filtros" id="filtroModeloGraficoConsumo">${botoesModelo}</div>
            </div>

            <div class="detail-section" id="blocoLimiteGraficoConsumo">
                <h4>Quantidade de PN no gráfico</h4>
                <div class="trend-chart-subtitle">Defina quantos part numbers serão exibidos para manter boa performance.</div>
                <div class="consumo-toolbar">
                    <div class="filtro-buttons consumo-cascata-filtros" id="filtroLimiteGraficoConsumo">
                        <button class="filtro-btn ${limitePnGraficoConsumo === '5' ? 'ativo' : ''}" data-limite-grafico="5">5</button>
                        <button class="filtro-btn ${limitePnGraficoConsumo === '10' ? 'ativo' : ''}" data-limite-grafico="10">10</button>
                        <button class="filtro-btn ${limitePnGraficoConsumo === '20' ? 'ativo' : ''}" data-limite-grafico="20">20</button>
                        <button class="filtro-btn ${limitePnGraficoConsumo === 'todos' ? 'ativo' : ''}" data-limite-grafico="todos">Todos</button>
                    </div>
                    <div class="consumo-select-wrap">
                        <label class="consumo-select-label" for="selectPnGraficoConsumo">PN</label>
                        <select id="selectPnGraficoConsumo" class="auto-refresh-select consumo-select"></select>
                    </div>
                </div>
            </div>

            <div class="trend-chart-wrap consumo-chart-wrap">
                <h4>Curva de consumo por part number</h4>
                <div class="trend-chart-subtitle" id="graficoConsumoResumo"></div>
                <canvas id="graficoConsumoCanvas" aria-label="Gráfico de consumo por part number"></canvas>
            </div>

            <div class="detail-section">
                <h4>Resumo do filtro</h4>
                <div class="consumo-resumo-grid" id="graficoConsumoCards"></div>
            </div>
        </div>
    `;
}

function configurarFiltroModalGraficoConsumo() {
    const filtroEl = document.getElementById('filtroModeloGraficoConsumo');
    const limiteEl = document.getElementById('filtroLimiteGraficoConsumo');
    const selectPnEl = document.getElementById('selectPnGraficoConsumo');
    if (!filtroEl) {
        return;
    }

    filtroEl.addEventListener('click', (e) => {
        const botao = e.target.closest('[data-modelo-grafico]');
        if (!botao) {
            return;
        }

        filtroModeloGraficoConsumo = botao.dataset.modeloGrafico;
        pnSelecionadoGraficoConsumo = 'todos';
        filtroEl.querySelectorAll('[data-modelo-grafico]').forEach((item) => {
            item.classList.toggle('ativo', item.dataset.modeloGrafico === filtroModeloGraficoConsumo);
        });
        renderizarGraficoConsumo();
    });

    limiteEl?.addEventListener('click', (e) => {
        const botao = e.target.closest('[data-limite-grafico]');
        if (!botao) {
            return;
        }

        limitePnGraficoConsumo = botao.dataset.limiteGrafico;
        limiteEl.querySelectorAll('[data-limite-grafico]').forEach((item) => {
            item.classList.toggle('ativo', item.dataset.limiteGrafico === limitePnGraficoConsumo);
        });
        renderizarGraficoConsumo();
    });

    selectPnEl?.addEventListener('change', (e) => {
        pnSelecionadoGraficoConsumo = e.target.value || 'todos';
        renderizarGraficoConsumo();
    });
}

function obterModelosGraficoConsumo() {
    return [...new Set(
        dadosFiltrados
            .filter((item) => itemPodeAparecerNoGraficoConsumo(item))
            .map((item) => item.modelo)
            .filter(Boolean)
    )].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR'));
}

function obterDadosGraficoConsumo() {
    const itensBase = dadosFiltrados.filter((item) => itemPodeAparecerNoGraficoConsumo(item));
    const itensFiltrados = filtroModeloGraficoConsumo === 'todos'
        ? itensBase
        : itensBase.filter((item) => item.modelo === filtroModeloGraficoConsumo);

    let itensOrdenados = [...itensFiltrados].sort((a, b) => {
        const diferencaConsumo = numeroSeguro(b.consumo, 0) - numeroSeguro(a.consumo, 0);
        if (diferencaConsumo !== 0) {
            return diferencaConsumo;
        }

        return String(a.pn).localeCompare(String(b.pn), 'pt-BR');
    });

    totalItensGraficoAntesLimite = itensOrdenados.length;
    limiteSegurancaGraficoAplicado = false;

    if (pnSelecionadoGraficoConsumo !== 'todos') {
        itensOrdenados = itensOrdenados.filter((item) => String(item.pn) === String(pnSelecionadoGraficoConsumo));
        totalItensGraficoAntesLimite = itensOrdenados.length;
        return itensOrdenados;
    }

    if (limitePnGraficoConsumo !== 'todos') {
        return itensOrdenados.slice(0, Number(limitePnGraficoConsumo));
    }

    if (itensOrdenados.length > LIMITE_SEGURANCA_GRAFICO_CONSUMO) {
        limiteSegurancaGraficoAplicado = true;
        return itensOrdenados.slice(0, LIMITE_SEGURANCA_GRAFICO_CONSUMO);
    }

    return itensOrdenados;
}

function obterPnDisponiveisGraficoConsumo() {
    const itensBase = dadosFiltrados.filter((item) => itemPodeAparecerNoGraficoConsumo(item));
    const itensModelo = filtroModeloGraficoConsumo === 'todos'
        ? itensBase
        : itensBase.filter((item) => item.modelo === filtroModeloGraficoConsumo);

    return [...new Set(itensModelo.map((item) => String(item.pn)).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b, 'pt-BR'));
}

function renderizarGraficoConsumo() {
    const canvas = document.getElementById('graficoConsumoCanvas');
    const resumoEl = document.getElementById('graficoConsumoResumo');
    const cardsResumoEl = document.getElementById('graficoConsumoCards');

    if (!canvas || !resumoEl || !cardsResumoEl || typeof Chart === 'undefined') {
        return;
    }

    const itens = obterDadosGraficoConsumo();
    const modeloTexto = filtroModeloGraficoConsumo === 'todos'
        ? 'todos os modelos'
        : `modelo ${filtroModeloGraficoConsumo}`;
    const blocoLimiteEl = document.getElementById('blocoLimiteGraficoConsumo');
    const limiteEl = document.getElementById('filtroLimiteGraficoConsumo');
    const selectPnEl = document.getElementById('selectPnGraficoConsumo');
    const pnDisponiveis = obterPnDisponiveisGraficoConsumo();

    if (blocoLimiteEl) {
        blocoLimiteEl.classList.remove('hidden');
    }

    if (limiteEl) {
        limiteEl.querySelectorAll('[data-limite-grafico]').forEach((item) => {
            item.classList.toggle('ativo', item.dataset.limiteGrafico === limitePnGraficoConsumo);
        });
    }

    if (pnSelecionadoGraficoConsumo !== 'todos' && !pnDisponiveis.includes(pnSelecionadoGraficoConsumo)) {
        pnSelecionadoGraficoConsumo = 'todos';
    }

    if (selectPnEl) {
        const opcoesPn = [
            '<option value="todos">Todos os PN</option>',
            ...pnDisponiveis.map((pn) => `<option value="${escapeHtml(pn)}" ${pnSelecionadoGraficoConsumo === pn ? 'selected' : ''}>${escapeHtml(pn)}</option>`)
        ];
        selectPnEl.innerHTML = opcoesPn.join('');
        selectPnEl.disabled = filtroModeloGraficoConsumo === 'todos' || pnDisponiveis.length === 0;
        selectPnEl.value = pnSelecionadoGraficoConsumo;
    }

    if (itens.length === 0) {
        resumoEl.textContent = `Nenhum item disponível para ${modeloTexto}.`;
        cardsResumoEl.innerHTML = '<div class="trend-stat">Sem dados para exibir.</div>';
        if (graficoConsumoChart) {
            removerHoverTooltipConsumo();
            graficoConsumoChart.destroy();
            graficoConsumoChart = null;
            graficoConsumoFingerprint = '';
        }
        return;
    }

    const totalConsumo = itens.reduce((acc, item) => acc + numeroSeguro(item.consumo), 0);
    const itemMaisCritico = [...itens].sort((a, b) => {
        const dataA = obterTimestampSeguro(a.dataFim) ?? Number.POSITIVE_INFINITY;
        const dataB = obterTimestampSeguro(b.dataFim) ?? Number.POSITIVE_INFINITY;
        return dataA - dataB;
    })[0];
    const rootStyle = getComputedStyle(document.documentElement);
    const textColor = rootStyle.getPropertyValue('--text').trim() || '#0f2036';

    const avisoSeguranca = limiteSegurancaGraficoAplicado
        ? ` Exibição limitada automaticamente em ${LIMITE_SEGURANCA_GRAFICO_CONSUMO} PN para manter desempenho.`
        : '';
    resumoEl.textContent = `${itens.length} part number(s) exibidos para ${modeloTexto}${pnSelecionadoGraficoConsumo !== 'todos' ? `, PN ${pnSelecionadoGraficoConsumo}` : ''}${limitePnGraficoConsumo !== 'todos' && pnSelecionadoGraficoConsumo === 'todos' ? `, limite ${limitePnGraficoConsumo}` : ''}.${avisoSeguranca}`;
    cardsResumoEl.innerHTML = `
        <div class="trend-stat">
            <span class="trend-stat-label">Modelo selecionado</span>
            <span class="trend-stat-value">${escapeHtml(filtroModeloGraficoConsumo === 'todos' ? 'Todos' : filtroModeloGraficoConsumo)}</span>
        </div>
        <div class="trend-stat">
            <span class="trend-stat-label">Total de PN</span>
            <span class="trend-stat-value">${itens.length}${totalItensGraficoAntesLimite > itens.length ? ` de ${totalItensGraficoAntesLimite}` : ''}</span>
        </div>
        <div class="trend-stat">
            <span class="trend-stat-label">Consumo total</span>
            <span class="trend-stat-value">${totalConsumo.toFixed(1)} p/h</span>
        </div>
        <div class="trend-stat">
            <span class="trend-stat-label">Acaba primeiro</span>
            <span class="trend-stat-value">${escapeHtml(String(itemMaisCritico.pn))} (${escapeHtml(formatarDataHora(new Date(itemMaisCritico.dataFim)))})</span>
        </div>
    `;

    const { labels, datasets } = gerarDatasetsGraficoConsumo(itens);
    const fingerprintGrafico = gerarFingerprintGrafico({
        tipo: 'consumo',
        textColor,
        labels,
        datasets: datasets.map((dataset) => ({
            label: dataset.label,
            data: dataset.data,
            fimPrevisto: dataset.metaInfo?.fimPrevisto || ''
        }))
    });

    if (
        graficoConsumoChart
        && graficoConsumoFingerprint === fingerprintGrafico
        && graficoConsumoChart.canvas === canvas
    ) {
        configurarHoverTooltipConsumo();
        return;
    }

    if (graficoConsumoChart) {
        removerHoverTooltipConsumo();
        graficoConsumoChart.destroy();
        graficoConsumoChart = null;
    }

    graficoConsumoChart = new Chart(canvas, {
        type: 'line',
        data: {
            labels,
            datasets
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            resizeDelay: 120,
            events: ['click', 'mouseout', 'touchstart', 'touchend'],
            interaction: {
                mode: 'point',
                intersect: true
            },
            plugins: {
                legend: {
                    labels: {
                        color: textColor,
                        font: { size: 11, weight: 700 },
                        boxWidth: 12
                    }
                },
                tooltip: {
                    enabled: true,
                    mode: 'point',
                    intersect: true,
                    callbacks: {
                        label: (context) => `${context.dataset.label}: ${Number(context.parsed.y).toFixed(0)} peças`,
                        afterLabel: (context) => {
                            const meta = context.dataset.metaInfo;
                            return meta ? `Acaba em: ${meta.fimPrevisto}` : '';
                        }
                    }
                }
            },
            scales: {
                x: {
                    ticks: {
                        color: textColor,
                        maxRotation: 50,
                        minRotation: 0
                    },
                    grid: {
                        color: 'rgba(120, 140, 170, 0.1)'
                    }
                },
                y: {
                    beginAtZero: true,
                    ticks: {
                        color: textColor
                    },
                    grid: {
                        color: 'rgba(120, 140, 170, 0.18)'
                    },
                    title: {
                        display: true,
                        text: 'Estoque restante',
                        color: textColor
                    }
                }
            }
        }
    });
    graficoConsumoFingerprint = fingerprintGrafico;
    configurarHoverTooltipConsumo();
}

function configurarHoverTooltipConsumo() {
    if (!graficoConsumoChart || !graficoConsumoChart.canvas) {
        removerHoverTooltipConsumo();
        return;
    }

    removerHoverTooltipConsumo();

    const chart = graficoConsumoChart;
    const canvas = chart.canvas;

    const limparTooltipInterno = () => {
        if (!ultimoHoverTooltipConsumo) {
            return;
        }
        ultimoHoverTooltipConsumo = '';
        chart.setActiveElements([]);
        if (chart.tooltip) {
            chart.tooltip.setActiveElements([], { x: 0, y: 0 });
        }
        chart.update('none');
    };

    const atualizarTooltip = (evento) => {
        const pontoEvento = evento?.touches?.[0] || evento?.changedTouches?.[0] || evento;
        if (!pontoEvento || typeof pontoEvento.clientX !== 'number' || typeof pontoEvento.clientY !== 'number') {
            return;
        }

        const rect = canvas.getBoundingClientRect();
        if (!rect.width || !rect.height) {
            return;
        }

        const x = (pontoEvento.clientX - rect.left) * (chart.width / rect.width);
        const y = (pontoEvento.clientY - rect.top) * (chart.height / rect.height);

        let melhor = null;
        let menorDistancia = Number.POSITIVE_INFINITY;

        chart.data.datasets.forEach((dataset, datasetIndex) => {
            const meta = chart.getDatasetMeta(datasetIndex);
            if (!meta || meta.hidden) {
                return;
            }

            meta.data.forEach((ponto, index) => {
                const pos = ponto.getProps(['x', 'y', 'skip'], true);
                if (pos.skip || !Number.isFinite(pos.x) || !Number.isFinite(pos.y)) {
                    return;
                }

                const distancia = Math.hypot(pos.x - x, pos.y - y);
                if (distancia < menorDistancia) {
                    menorDistancia = distancia;
                    melhor = { datasetIndex, index, x: pos.x, y: pos.y, dataset };
                }
            });
        });

        if (!melhor) {
            limparTooltipInterno();
            return;
        }

        const raio = numeroSeguro(melhor.dataset?.pointRadius, 3);
        const distanciaMaxima = Math.max(4, raio + 1);
        if (menorDistancia > distanciaMaxima) {
            limparTooltipInterno();
            return;
        }

        const ativos = [{ datasetIndex: melhor.datasetIndex, index: melhor.index }];
        const chaveHover = `${melhor.datasetIndex}:${melhor.index}`;
        if (ultimoHoverTooltipConsumo === chaveHover) {
            return;
        }
        ultimoHoverTooltipConsumo = chaveHover;
        chart.setActiveElements(ativos);
        if (chart.tooltip) {
            chart.tooltip.setActiveElements(ativos, { x: melhor.x, y: melhor.y });
        }
        chart.update('none');
    };

    handlerHoverGraficoConsumo = {
        atualizarTooltip,
        limparTooltip: limparTooltipInterno
    };

    canvas.addEventListener('mousemove', atualizarTooltip, { passive: true });
    canvas.addEventListener('touchmove', atualizarTooltip, { passive: true });
    canvas.addEventListener('mouseleave', limparTooltip, { passive: true });
    canvas.addEventListener('touchend', limparTooltip, { passive: true });
}

function removerHoverTooltipConsumo() {
    if (!handlerHoverGraficoConsumo || !graficoConsumoChart?.canvas) {
        ultimoHoverTooltipConsumo = '';
        handlerHoverGraficoConsumo = null;
        return;
    }

    const canvas = graficoConsumoChart.canvas;
    canvas.removeEventListener('mousemove', handlerHoverGraficoConsumo.atualizarTooltip);
    canvas.removeEventListener('touchmove', handlerHoverGraficoConsumo.atualizarTooltip);
    canvas.removeEventListener('mouseleave', handlerHoverGraficoConsumo.limparTooltip);
    canvas.removeEventListener('touchend', handlerHoverGraficoConsumo.limparTooltip);
    ultimoHoverTooltipConsumo = '';
    handlerHoverGraficoConsumo = null;
}

function gerarDatasetsGraficoConsumo(itens) {
    const itensOrdenados = [...itens].sort((a, b) => {
        const dataA = obterTimestampSeguro(a.dataFim) ?? Number.POSITIVE_INFINITY;
        const dataB = obterTimestampSeguro(b.dataFim) ?? Number.POSITIVE_INFINITY;
        return dataA - dataB;
    });
    const inicioComum = new Date();
    const fimMaisDistante = itensOrdenados.reduce((maior, item) => {
        const timestamp = obterTimestampSeguro(item.dataFim) ?? inicioComum.getTime();
        return Math.max(maior, timestamp);
    }, inicioComum.getTime());
    const series = itensOrdenados.map((item) => ({
        item,
        serie: gerarSerieEstoqueReal(item, {
            inicio: inicioComum,
            fim: fimMaisDistante,
            numPontos: NUM_PONTOS_GRAFICO
        })
    }));
    const labels = series[0]?.serie?.labels || [];

    const palette = [
        '#e20d2a',
        '#1894ff',
        '#29c177',
        '#f5bb2e',
        '#8b5cf6',
        '#ff7a18',
        '#14b8a6',
        '#ef4444',
        '#3b82f6',
        '#84cc16'
    ];

    const datasets = series.map(({ item, serie }, index) => {
        const cor = palette[index % palette.length];
        return {
            label: `${item.pn} (${item.modelo || '-'})`,
            data: serie.valores,
            borderColor: cor,
            backgroundColor: `${cor}22`,
            fill: false,
            tension: 0.12,
            pointRadius: 3,
            pointHoverRadius: 5,
            pointHitRadius: 0,
            hitRadius: 0,
            hoverRadius: 5,
            borderWidth: 2,
            metaInfo: {
                fimPrevisto: serie.fimPrevisto
            }
        };
    });

    return { labels, datasets };
}

function fecharModal() {
    modalDetalhes.classList.remove('mostrar');
    modalContent?.classList.remove('modal-consumo');
    destruirGraficosModal();
    modalBody.innerHTML = '';
}

function destruirGraficosModal() {
    if (graficoConsumoChart) {
        removerHoverTooltipConsumo();
        graficoConsumoChart.destroy();
        graficoConsumoChart = null;
    }
    graficoConsumoFingerprint = '';
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

function gerarSerieEstoqueReal(item, opcoes = {}) {
    const agora = opcoes.inicio ? new Date(opcoes.inicio) : new Date();
    const dataFim = new Date(item.dataFim);
    const estoqueInicial = numeroSeguro(item.estoque, 0);
    const consumoHora = numeroSeguro(item.consumo, 0);
    const labels = [];
    const valores = [];
    const timestampAgora = agora.getTime();
    const timestampFimItem = Number.isNaN(dataFim.getTime())
        ? timestampAgora + (numeroSeguro(item.autonomiaHoras, 24) * 3600000)
        : dataFim.getTime();
    const timestampFimGrafico = opcoes.fim
        ? new Date(opcoes.fim).getTime()
        : timestampFimItem;
    const duracaoHoras = Math.max(0.1, (timestampFimGrafico - timestampAgora) / 3600000);
    const numPontos = opcoes.numPontos || NUM_PONTOS_GRAFICO;
    const passoHoras = duracaoHoras / (numPontos - 1);
    const horasAteFimItem = Math.max(0, (timestampFimItem - timestampAgora) / 3600000);

    for (let i = 0; i < numPontos; i += 1) {
        const horasPassadas = i * passoHoras;
        const dataHora = new Date(timestampAgora + horasPassadas * 60 * 60 * 1000);
        const label = duracaoHoras <= 48
            ? `${String(dataHora.getHours()).padStart(2, '0')}:00`
            : `${String(dataHora.getDate()).padStart(2, '0')}/${String(dataHora.getMonth() + 1).padStart(2, '0')} ${String(dataHora.getHours()).padStart(2, '0')}:00`;
        const horasConsideradas = Math.min(horasPassadas, horasAteFimItem);
        const valor = Math.max(0, estoqueInicial - (consumoHora * horasConsideradas));

        labels.push(label);
        valores.push(Number(valor.toFixed(2)));
    }

    return {
        labels,
        valores,
        estoqueInicial,
        consumoHora,
        fimPrevisto: formatarDataHora(new Date(timestampFimItem))
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
    if (buscaDebounceTimer) {
        clearTimeout(buscaDebounceTimer);
        buscaDebounceTimer = null;
    }

    filtros = {
        modelo: 'todos',
        turno: 'todos',
        status: 'todos',
        ordem: 'urgencia',
        limiteCards: 'todos',
        busca: ''
    };

    buscaInput.value = '';

    [
        [filtroModeloEl, 'modelo', 'todos'],
        [filtroTurnoEl, 'turno', 'todos'],
        [filtroStatusEl, 'status', 'todos'],
        [filtroOrdemEl, 'ordem', 'urgencia'],
        [filtroLimiteCardsEl, 'limiteCards', 'todos']
    ].forEach(([container, dataAttr, valor]) => {
        container?.querySelectorAll('.filtro-btn').forEach((btn) => {
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
        carregarDados({ silencioso: true });
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

function gerarFingerprintDados(lista) {
    if (!Array.isArray(lista) || lista.length === 0) {
        return 'dados:0';
    }

    const assinatura = lista
        .map((item) => gerarAssinaturaItem(item))
        .sort((a, b) => a.localeCompare(b))
        .join('||');

    return `dados:${lista.length}:${hashStringFNV1a(assinatura)}`;
}

function gerarAssinaturaItem(item) {
    const timestampFim = obterTimestampSeguro(item?.dataFim);

    return [
        String(item?.pn ?? ''),
        String(item?.modelo ?? ''),
        String(item?.linha ?? ''),
        numeroSeguro(item?.estoque, 0).toFixed(4),
        numeroSeguro(item?.consumo, 0).toFixed(4),
        numeroSeguro(item?.autonomiaHoras, 0).toFixed(4),
        timestampFim === null ? '' : String(timestampFim),
        String(item?.turnoPrevisto ?? '')
    ].join('|');
}

function gerarFingerprintListaModelos(modelos) {
    if (!Array.isArray(modelos) || modelos.length === 0) {
        return 'modelos:0';
    }

    return `modelos:${modelos.length}:${hashStringFNV1a(modelos.map((item) => String(item)).join('|'))}`;
}

function gerarFingerprintGrafico(payload) {
    return `grafico:${hashStringFNV1a(JSON.stringify(payload))}`;
}

function hashStringFNV1a(texto) {
    let hash = 2166136261;
    for (let i = 0; i < texto.length; i += 1) {
        hash ^= texto.charCodeAt(i);
        hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16).padStart(8, '0');
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

function obterTimestampEsgotamento(item) {
    const timestampDataFim = obterTimestampSeguro(item?.dataFim);
    if (timestampDataFim !== null) {
        return timestampDataFim;
    }

    const autonomiaHoras = numeroSeguro(item?.autonomiaHoras, Number.POSITIVE_INFINITY);
    if (!Number.isFinite(autonomiaHoras)) {
        return null;
    }

    return Date.now() + (autonomiaHoras * 3600000);
}

function itemPodeSerExibido(item) {
    const estoque = numeroSeguro(item?.estoque, 0);
    const consumo = numeroSeguro(item?.consumo, 0);
    return !(estoque <= 0 && consumo <= 0);
}

function itemPodeAparecerNoGraficoConsumo(item) {
    const estoque = numeroSeguro(item?.estoque, 0);
    const consumo = numeroSeguro(item?.consumo, 0);
    return estoque > 0 && consumo > 0;
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

function mostrarLoading(ativo, limparCards = true) {
    if (ativo) {
        btnAtualizar.disabled = true;
        btnAtualizar.innerHTML = '<span class="loading"></span> Atualizando...';
        if (limparCards) {
            desconectarObserverLazyCards();
            cardsVisiveis = [];
            cursorCardsVisiveis = 0;
            cardsContainer.classList.remove('modo-tabela');
            cardsContainer.innerHTML = '<div class="loading-cards"><span class="loading"></span> Carregando dados da API...</div>';
        }
    } else {
        btnAtualizar.disabled = false;
        btnAtualizar.innerHTML = '<span>↻</span> Atualizar dados';
    }
}

function mostrarErro(mensagem) {
    desconectarObserverLazyCards();
    cardsVisiveis = [];
    cursorCardsVisiveis = 0;
    cardsContainer.classList.remove('modo-tabela');
    cardsContainer.innerHTML = `<div class="loading-cards" style="color: #d42f4a;">${escapeHtml(mensagem)}</div>`;
}




