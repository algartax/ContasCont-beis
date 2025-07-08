// ============================================
// SCRIPT1.JS - PARTE 1 - BASE E FUNCIONALIDADES CORE
// ============================================
// VERSÃO CORRIGIDA PARA RESOLVER DUPLICAÇÃO DE GRÁFICOS
// Este arquivo contém funcionalidades que raramente precisam ser alteradas:
// - Configurações e variáveis globais
// - Funções utilitárias 
// - Sistema de upload e processamento Excel
// - Gestão de modais e UI básica
// - Sistema de filtros dinâmicos
// - Funcionalidades de comparação avançada

// ============================================
// VARIÁVEIS GLOBAIS E CONFIGURAÇÕES
// ============================================

let workbook = null;
let allSheetsData = {};
let selectedSheet = '';
let rawData = [];
let filteredData = [];
let periods = [];
let charts = {};
let chartViewMode = 'monthly';

// NOVA VARIÁVEL: Controle de debounce para filtros
let filterTimeout = null;

// Variáveis para o sistema de comparações
let comparisonMode = {
    active: false,
    type: 'mensal', // mensal, trimestral, semestral, anual
    period1: null,
    period2: null,
    entityFilter: 'total',
    minVariation: null,
    maxResults: 'all'
};

// Sheet filter configurations
const SHEET_FILTERS = {
    'SINTETICA': [
        { id: 'contaSearch', label: 'Buscar Conta', field: 'CONTA_SEARCH', type: 'search' },
        { id: 'descricaoDf', label: 'Descrição DF', field: 'Descrição DF' },
        { id: 'descricaoDfGroup', label: 'Descrição DF Group', field: 'Group' }
    ],
    'ANALITICA': [
        { id: 'contaSearch', label: 'Buscar Conta', field: 'CONTA_SEARCH', type: 'search' },
        { id: 'fsGroupReport', label: 'FS Group REPORT', field: 'FS Group REPORT' },
        { id: 'fsGroup', label: 'FS Group SAP', field: 'FS Group SAP' },
        { id: 'sintetica', label: 'Sintética', field: 'SINTETICA' }
    ],
    'ABERTURA DF': [
        { id: 'contaSearch', label: 'Buscar Conta', field: 'CONTA_SEARCH', type: 'search' },
        { id: 'descricaoDf', label: 'Descrição DF', field: 'Descrição DF' }
    ],
    'SUBGRUPO': [
        { id: 'contaSearch', label: 'Buscar Conta', field: 'CONTA_SEARCH', type: 'search' },
        { id: 'descricaoDf', label: 'Descrição DF', field: 'Descrição DF' }
    ],
    'GRUPO FEC': [
        { id: 'contaSearch', label: 'Buscar Conta', field: 'CONTA_SEARCH', type: 'search' },
        { id: 'descricaoDf', label: 'Descrição DF', field: 'Descrição DF' }
    ]
};

// ============================================
// FUNÇÕES UTILITÁRIAS
// ============================================

function formatPeriod(period) {
    try {
        const [day, month, year] = period.split('/');
        const date = new Date(year, month - 1, day);
        
        const months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        return `${months[date.getMonth()]}/${year}`;
    } catch (error) {
        console.error('Error parsing period:', error);
        return period;
    }
}

function formatQuarter(period) {
    try {
        const [day, month, year] = period.split('/');
        const monthNum = parseInt(month);
        
        let quarter;
        if (monthNum >= 1 && monthNum <= 3) quarter = '1T';
        else if (monthNum >= 4 && monthNum <= 6) quarter = '2T';
        else if (monthNum >= 7 && monthNum <= 9) quarter = '3T';
        else quarter = '4T';
        
        return `${quarter} ${year}`;
    } catch (error) {
        console.error('Error parsing period for quarter:', error);
        return period;
    }
}

function formatCurrency(value) {
    return new Intl.NumberFormat('pt-BR', {
        style: 'currency',
        currency: 'BRL',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(value);
}

function formatCurrencyCompact(value) {
    if (Math.abs(value) >= 1000000000) {
        return 'R$ ' + (value / 1000000000).toFixed(1) + 'B';
    } else if (Math.abs(value) >= 1000000) {
        return 'R$ ' + (value / 1000000).toFixed(1) + 'M';
    } else if (Math.abs(value) >= 1000) {
        return 'R$ ' + (value / 1000).toFixed(1) + 'K';
    }
    return formatCurrency(value);
}

// CORRIGIDO: Extrair informações da conta com foco no NÚMERO da conta
function extractAccountInfo(item) {
    let contaSAP = '';
    let description = '';
    
    // Lista de possíveis nomes de colunas para conta SAP (prioriza número)
    const contaFields = [
        'Conta SAP', 'CONTA SAP', 'Número Conta', 'NÚMERO CONTA',
        'ContaSAP', 'CONTASAP', 'NumeroConta', 'NUMEROCONTA',
        'Conta', 'CONTA', 'Account', 'ACCOUNT'
    ];
    
    // Lista de possíveis nomes de colunas para descrição
    const descriptionFields = [
        'Descrição Conta', 'Descrição DF', 'DESCRIÇÃO CONTA', 'DESCRIÇÃO DF',
        'DescricaoConta', 'DescricaoDF', 'Description', 'DESCRIPTION',
        'Descricao', 'DESCRICAO', 'Nome', 'NOME'
    ];
    
    // Buscar conta SAP (PRIORIDADE MÁXIMA)
    for (const field of contaFields) {
        if (item.hasOwnProperty(field) && item[field] !== null && item[field] !== undefined) {
            const value = String(item[field]).trim();
            if (value && value !== '') {
                contaSAP = value;
                break;
            }
        }
    }
    
    // Buscar descrição
    for (const field of descriptionFields) {
        if (item.hasOwnProperty(field) && item[field] !== null && item[field] !== undefined) {
            const value = String(item[field]).trim();
            if (value && value !== '') {
                description = value;
                break;
            }
        }
    }
    
    // Se ainda não encontrou conta SAP, buscar por qualquer chave que contenha esses termos
    if (!contaSAP) {
        for (const key of Object.keys(item)) {
            const keyLower = key.toLowerCase().replace(/\s+/g, '');
            if (keyLower.includes('conta') || keyLower.includes('account') || keyLower.includes('numero')) {
                const value = String(item[key] || '').trim();
                if (value && value !== '') {
                    contaSAP = value;
                    break;
                }
            }
        }
    }
    
    return { contaSAP, description };
}

// Safe data access helper functions
function safeGetItemValue(item, period, entity) {
    if (!item.values || !item.values[period]) {
        return 0;
    }
    return item.values[period][entity] || 0;
}

function safeCalculateSummary(data, currentPeriod, previousPeriod, entity) {
    let currentTotal = 0;
    let previousTotal = 0;
    
    data.forEach(item => {
        currentTotal += safeGetItemValue(item, currentPeriod, entity);
        previousTotal += safeGetItemValue(item, previousPeriod, entity);
    });
    
    return {
        currentTotal,
        previousTotal,
        variation: previousTotal !== 0 ? ((currentTotal - previousTotal) / previousTotal) * 100 : 0
    };
}

// ============================================
// FUNÇÕES DE COMPARAÇÃO DE PERÍODOS
// ============================================

// Mapear períodos para diferentes tipos de agrupamento
function mapPeriodsByType(type) {
    const periodMap = {};
    
    periods.forEach(period => {
        try {
            const [day, month, year] = period.split('/');
            const date = new Date(year, month - 1, day);
            let key;
            
            switch (type) {
                case 'mensal':
                    key = `${String(month).padStart(2, '0')}/${year}`;
                    break;
                case 'trimestral':
                    const quarter = Math.ceil(parseInt(month) / 3);
                    key = `${quarter}T ${year}`;
                    break;
                case 'semestral':
                    const semester = parseInt(month) <= 6 ? 1 : 2;
                    key = `${semester}S ${year}`;
                    break;
                case 'anual':
                    key = year;
                    break;
                default:
                    key = period;
            }
            
            if (!periodMap[key]) {
                periodMap[key] = [];
            }
            periodMap[key].push(period);
        } catch (error) {
            console.error('Erro ao processar período:', period, error);
        }
    });
    
    return periodMap;
}

// Calcular total para um conjunto de períodos
function calculateTotalForPeriods(data, periods, entity) {
    let total = 0;
    
    data.forEach(item => {
        periods.forEach(period => {
            total += safeGetItemValue(item, period, entity);
        });
    });
    
    return total;
}

// Aplicar comparação de períodos
function applyPeriodComparison() {
    console.log('🔄 Aplicando comparação de períodos:', comparisonMode);
    
    if (!comparisonMode.period1 || !comparisonMode.period2) {
        alert('Selecione ambos os períodos para comparação');
        return;
    }
    
    const periodMap = mapPeriodsByType(comparisonMode.type);
    const periods1 = periodMap[comparisonMode.period1];
    const periods2 = periodMap[comparisonMode.period2];
    
    if (!periods1 || !periods2) {
        alert('Períodos inválidos para comparação');
        return;
    }
    
    console.log(`Comparando períodos:`, {
        period1: comparisonMode.period1,
        periods1,
        period2: comparisonMode.period2,
        periods2
    });
    
    // Atualizar headers da tabela
    updateTableHeaders(comparisonMode.period1, comparisonMode.period2);
    
    // Filtrar e ordenar dados baseado na comparação
    updateDetailTableWithComparison(periods1, periods2);
    
    // Ativar modo de comparação
    comparisonMode.active = true;
    
    // Fechar modal
    document.getElementById('comparisonModal').classList.remove('active');
    
    // Mostrar resumo
    showComparisonSummary(periods1, periods2);
}

// Atualizar headers da tabela para comparação
function updateTableHeaders(period1Label, period2Label) {
    const table = document.getElementById('detailsTable');
    const headers = table.querySelectorAll('thead th');
    
    if (headers.length >= 7) {
        headers[0].textContent = 'Empresa';
        headers[1].textContent = 'Conta SAP';
        headers[2].textContent = 'Descrição';
        headers[3].textContent = period2Label; // Período anterior
        headers[4].textContent = period1Label; // Período atual
        headers[5].textContent = 'Variação';
        headers[6].textContent = 'Var %';
    }
}

// CORRIGIDO: Filtro de conta por NÚMERO na comparação
function updateDetailTableWithComparison(periods1, periods2) {
    const tbody = document.querySelector('#detailsTable tbody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    // Verificar se deve mostrar formato expandido
    if (shouldShowExpandedFormat() && filteredData.length > 0) {
        // NOVO: Mostrar as 3 empresas para comparação de conta específica
        const item = filteredData[0]; // Pega a primeira conta
        const { contaSAP, description } = extractAccountInfo(item);
        
        const entities = [
            { key: 'telecom', name: 'Telecom' },
            { key: 'vogel', name: 'Vogel' },
            { key: 'consolidado', name: 'Somado' }
        ];
        
        // SEMPRE mostrar as 3 empresas quando conta específica (ignorar filtro dropdown)
        const entitiesToShow = entities;
        
        entitiesToShow.forEach(entity => {
            let value1 = 0, value2 = 0;
            
            periods1.forEach(p => {
                value1 += safeGetItemValue(item, p, entity.key);
            });
            periods2.forEach(p => {
                value2 += safeGetItemValue(item, p, entity.key);
            });
            
            const variation = value1 - value2;
            const variationPercent = value2 !== 0 ? (variation / value2) * 100 : 0;
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td><strong>${entity.name}</strong></td>
                <td><strong>${contaSAP || '-'}</strong></td>
                <td>${description || '-'}</td>
                <td>${formatCurrency(value2)}</td>
                <td>${formatCurrency(value1)}</td>
                <td class="variation-cell ${variation >= 0 ? 'positive' : 'negative'}">
                    ${formatCurrency(variation)}
                </td>
                <td class="variation-cell ${variationPercent >= 0 ? 'positive' : 'negative'}">
                    ${variationPercent >= 0 ? '+' : ''}${variationPercent.toFixed(2)}%
                </td>
            `;
            tbody.appendChild(row);
        });
        
        console.log(`✅ Tabela comparação atualizada com 3 empresas para conta: ${contaSAP}`);
        return;
    }
    
    // Calcular dados para cada conta (modo normal)
    const comparisonData = filteredData.map(item => {
        const { contaSAP, description } = extractAccountInfo(item);
        
        let value1 = 0, value2 = 0;
        
        // Calcular valores para cada conjunto de períodos
        if (comparisonMode.entityFilter === 'total') {
            periods1.forEach(p => {
                if (item.values && item.values[p]) {
                    value1 += Object.values(item.values[p]).reduce((a, b) => a + b, 0);
                }
            });
            periods2.forEach(p => {
                if (item.values && item.values[p]) {
                    value2 += Object.values(item.values[p]).reduce((a, b) => a + b, 0);
                }
            });
        } else {
            periods1.forEach(p => {
                value1 += safeGetItemValue(item, p, comparisonMode.entityFilter);
            });
            periods2.forEach(p => {
                value2 += safeGetItemValue(item, p, comparisonMode.entityFilter);
            });
        }
        
        const variation = value1 - value2;
        const variationPercent = value2 !== 0 ? (variation / value2) * 100 : 0;
        
        return {
            contaSAP,
            description,
            value1,
            value2,
            variation,
            variationPercent,
            absVariationPercent: Math.abs(variationPercent)
        };
    });
    
    // Filtrar por variação mínima se especificada
    let filteredComparison = comparisonData;
    if (comparisonMode.minVariation) {
        filteredComparison = comparisonData.filter(item => 
            item.absVariationPercent >= comparisonMode.minVariation
        );
    }
    
    // Ordenar por maior variação absoluta
    filteredComparison.sort((a, b) => b.absVariationPercent - a.absVariationPercent);
    
    // Limitar resultados
    if (comparisonMode.maxResults !== 'all') {
        filteredComparison = filteredComparison.slice(0, parseInt(comparisonMode.maxResults));
    }
    
    // CORRIGIDO: Filtrar por descrição OU número da conta se houver filtro ativo
    const descriptionFilter = document.getElementById('descriptionFilter').value.toLowerCase();
    if (descriptionFilter) {
        filteredComparison = filteredComparison.filter(item => {
            const descMatch = item.description && item.description.toLowerCase().includes(descriptionFilter);
            const contaMatch = item.contaSAP && item.contaSAP.toString().toLowerCase().includes(descriptionFilter);
            return descMatch || contaMatch;
        });
    }
    
    // Renderizar resultados
    if (filteredComparison.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="7" class="no-data">Nenhuma conta encontrada com os critérios de comparação especificados.</td>`;
        tbody.appendChild(row);
        return;
    }
    
    filteredComparison.forEach(item => {
        const entityName = comparisonMode.entityFilter === 'total' ? 'Consolidado' : 
                         comparisonMode.entityFilter === 'consolidado' ? 'Somado' :
                         comparisonMode.entityFilter.charAt(0).toUpperCase() + comparisonMode.entityFilter.slice(1);
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${entityName}</td>
            <td>${item.contaSAP || '-'}</td>
            <td>${item.description || '-'}</td>
            <td>${formatCurrency(item.value2)}</td>
            <td>${formatCurrency(item.value1)}</td>
            <td class="variation-cell ${item.variation >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(item.variation)}
            </td>
            <td class="variation-cell ${item.variationPercent >= 0 ? 'positive' : 'negative'}">
                ${item.variationPercent >= 0 ? '+' : ''}${item.variationPercent.toFixed(2)}%
            </td>
        `;
        
        tbody.appendChild(row);
    });
    
    console.log(`✅ Tabela atualizada com ${filteredComparison.length} registros de comparação`);
}

// Mostrar resumo da comparação
function showComparisonSummary(periods1, periods2) {
    const summarySection = document.getElementById('comparisonSummary');
    const summaryContent = document.getElementById('summaryContent');
    
    // Calcular totais
    const totals1 = calculateComparisonTotals(periods1);
    const totals2 = calculateComparisonTotals(periods2);
    
    const totalVariation = {
        telecom: totals1.telecom - totals2.telecom,
        vogel: totals1.vogel - totals2.vogel,
        consolidado: totals1.consolidado - totals2.consolidado
    };
    
    const totalVariationPercent = {
        telecom: totals2.telecom !== 0 ? (totalVariation.telecom / totals2.telecom) * 100 : 0,
        vogel: totals2.vogel !== 0 ? (totalVariation.vogel / totals2.vogel) * 100 : 0,
        consolidado: totals2.consolidado !== 0 ? (totalVariation.consolidado / totals2.consolidado) * 100 : 0
    };
    
    summaryContent.innerHTML = `
        <div class="summary-item">
            <span class="summary-label">Tipo de Comparação:</span>
            <span class="summary-value">${comparisonMode.type.charAt(0).toUpperCase() + comparisonMode.type.slice(1)}</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Períodos:</span>
            <span class="summary-value">${comparisonMode.period1} vs ${comparisonMode.period2}</span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Telecom:</span>
            <span class="summary-value variation-cell ${totalVariation.telecom >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(totalVariation.telecom)} (${totalVariationPercent.telecom >= 0 ? '+' : ''}${totalVariationPercent.telecom.toFixed(2)}%)
            </span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Vogel:</span>
            <span class="summary-value variation-cell ${totalVariation.vogel >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(totalVariation.vogel)} (${totalVariationPercent.vogel >= 0 ? '+' : ''}${totalVariationPercent.vogel.toFixed(2)}%)
            </span>
        </div>
        <div class="summary-item">
            <span class="summary-label">Somado:</span>
            <span class="summary-value variation-cell ${totalVariation.consolidado >= 0 ? 'positive' : 'negative'}">
                ${formatCurrency(totalVariation.consolidado)} (${totalVariationPercent.consolidado >= 0 ? '+' : ''}${totalVariationPercent.consolidado.toFixed(2)}%)
            </span>
        </div>
    `;
    
    summarySection.style.display = 'block';
}

// Calcular totais para comparação
function calculateComparisonTotals(periodsArray) {
    const totals = { telecom: 0, vogel: 0, consolidado: 0 };
    
    filteredData.forEach(item => {
        periodsArray.forEach(period => {
            totals.telecom += safeGetItemValue(item, period, 'telecom');
            totals.vogel += safeGetItemValue(item, period, 'vogel');
            totals.consolidado += safeGetItemValue(item, period, 'consolidado');
        });
    });
    
    return totals;
}

// Comparações rápidas
function applyQuickComparison(type) {
    console.log('🚀 Aplicando comparação rápida:', type);
    
    if (!periods || periods.length < 2) {
        alert('Dados insuficientes para comparação rápida');
        return;
    }
    
    const lastPeriod = periods[periods.length - 1];
    const [day, month, year] = lastPeriod.split('/');
    
    let period1, period2, compType;
    
    try {
        switch (type) {
            case 'mes-anterior':
                compType = 'mensal';
                period1 = formatPeriod(lastPeriod);
                const prevMonth = periods[periods.length - 2];
                period2 = formatPeriod(prevMonth);
                break;
                
            case 'trimestre-anterior':
                compType = 'trimestral';
                const currentQuarter = Math.ceil(parseInt(month) / 3);
                period1 = `${currentQuarter}T ${year}`;
                const prevQuarter = currentQuarter === 1 ? `4T ${parseInt(year) - 1}` : `${currentQuarter - 1}T ${year}`;
                period2 = prevQuarter;
                break;
                
            case 'mesmo-periodo-ano-passado':
                compType = 'mensal';
                period1 = formatPeriod(lastPeriod);
                period2 = `${String(month).padStart(2, '0')}/${parseInt(year) - 1}`;
                break;
                
            case 'trimestre-mesmo-ano-passado':
                compType = 'trimestral';
                const quarter = Math.ceil(parseInt(month) / 3);
                period1 = `${quarter}T ${year}`;
                period2 = `${quarter}T ${parseInt(year) - 1}`;
                break;
        }
        
        // Verificar se os períodos existem
        const periodMap = mapPeriodsByType(compType);
        if (!periodMap[period1] || !periodMap[period2]) {
            alert(`Períodos não encontrados para comparação: ${period1} ou ${period2}`);
            return;
        }
        
        // Aplicar comparação
        comparisonMode.type = compType;
        comparisonMode.period1 = period1;
        comparisonMode.period2 = period2;
        comparisonMode.entityFilter = 'total';
        comparisonMode.minVariation = null;
        comparisonMode.maxResults = 250;
        
        // Atualizar interface
        document.querySelector(`[data-type="${compType}"]`).click();
        document.getElementById('comparisonPeriod1').value = period1;
        document.getElementById('comparisonPeriod2').value = period2;
        
        // Aplicar comparação
        applyPeriodComparison();
        
    } catch (error) {
        console.error('Erro na comparação rápida:', error);
        alert('Erro ao aplicar comparação rápida');
    }
}

// ============================================
// FUNÇÕES DE INTERFACE E CONTROLE
// ============================================

function showLoading(show) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (show) {
        loadingOverlay.classList.add('active');
    } else {
        loadingOverlay.classList.remove('active');
    }
}

function showUploadModal(show) {
    const modal = document.getElementById('uploadModal');
    if (show) {
        modal.classList.add('active');
    } else {
        modal.classList.remove('active');
    }
}

function showComparisonModal(show) {
    const modal = document.getElementById('comparisonModal');
    if (show) {
        modal.classList.add('active');
        populateComparisonPeriods();
    } else {
        modal.classList.remove('active');
    }
}

function toggleSidebar() {
    const sidebarContent = document.querySelector('.sidebar-content');
    sidebarContent.classList.toggle('expanded');
}

// CORRIGIDO: toggleChartView com limpeza forçada de gráficos
function toggleChartView(mode) {
    chartViewMode = mode;
    
    const monthlyBtn = document.getElementById('monthlyBtn');
    const quarterlyBtn = document.getElementById('quarterlyBtn');
    
    if (mode === 'monthly') {
        monthlyBtn.classList.add('active');
        quarterlyBtn.classList.remove('active');
    } else {
        monthlyBtn.classList.remove('active');
        quarterlyBtn.classList.add('active');
    }
    
    // CORREÇÃO CRÍTICA: Forçar limpeza antes de atualizar
    if (typeof forceChartCleanup === 'function') {
        forceChartCleanup();
    }
    
    // Aguardar antes de renderizar novo gráfico
    setTimeout(() => {
        updateEvolutionChart();
        if (typeof updateChartTitle === 'function') {
            updateChartTitle();
        }
    }, 200);
}

function updateCurrentDate() {
    const dateDisplay = document.getElementById('dateDisplay');
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    
    dateDisplay.textContent = `Data de análise: ${day}/${month}/${year}`;
}

// CORRIGIDO: handleResize com função segura de chart resize
function handleResize() {
    const sidebar = document.querySelector('.sidebar');
    const sidebarContent = document.querySelector('.sidebar-content');
    const sidebarHeader = document.querySelector('.sidebar-header');
    
    if (window.innerWidth <= 768) {
        if (!document.querySelector('.toggle-sidebar')) {
            const toggleButton = document.createElement('button');
            toggleButton.className = 'toggle-sidebar';
            toggleButton.innerHTML = '<i class="fas fa-bars"></i>';
            toggleButton.addEventListener('click', toggleSidebar);
            sidebarHeader.appendChild(toggleButton);
        }
    } else {
        const toggleButton = document.querySelector('.toggle-sidebar');
        if (toggleButton) {
            toggleButton.remove();
        }
        
        sidebarContent.classList.remove('expanded');
        sidebarContent.style.maxHeight = '';
    }
    
    // CORREÇÃO: Usar função segura de resize
    if (typeof handleChartResize === 'function') {
        handleChartResize();
    }
}

// ============================================
// UPLOAD E PROCESSAMENTO DE EXCEL
// ============================================

// CORRIGIDO: handleFileUpload com limpeza de gráficos
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    showLoading(true);
    
    // CORREÇÃO: Limpar gráficos antes de processar novo arquivo
    if (typeof forceChartCleanup === 'function') {
        forceChartCleanup();
    }
    
    try {
        await readExcelFile(file);
        
        const sheetSelector = document.getElementById('sheetSelector');
        if (sheetSelector.options.length > 1) {
            sheetSelector.selectedIndex = 1;
            handleSheetChange();
        }
        
        showUploadModal(false);
        showLoading(false);
    } catch (error) {
        console.error('Error processing file:', error);
        alert(`Erro ao processar arquivo: ${error.message}\n\nVerifique se é um arquivo Excel válido.`);
        showLoading(false);
    }
}

async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                
                const sheetNames = workbook.SheetNames;
                console.log('Available sheets:', sheetNames);
                
                const sheetSelector = document.getElementById('sheetSelector');
                sheetSelector.innerHTML = '<option value="">Selecione uma aba...</option>';
                
                sheetNames.forEach(sheetName => {
                    if (sheetName.toLowerCase().includes('3000 ant')) {
                        return;
                    }
                    
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelector.appendChild(option);
                });
                
                sheetNames.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    allSheetsData[sheetName] = jsonData;
                });
                
                resolve();
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// CORRIGIDO: handleSheetChange com limpeza de gráficos
function handleSheetChange() {
    const sheetSelector = document.getElementById('sheetSelector');
    const sheetName = sheetSelector.value;
    
    // CORREÇÃO: Limpar gráficos ao mudar aba
    if (typeof forceChartCleanup === 'function') {
        forceChartCleanup();
    }
    
    if (!sheetName) {
        document.getElementById('dynamicFilters').innerHTML = '';
        document.getElementById('currentPeriod').innerHTML = '<option value="">Selecione...</option>';
        document.getElementById('previousPeriod').innerHTML = '<option value="">Selecione...</option>';
        
        rawData = [];
        filteredData = [];
        periods = [];
        selectedSheet = '';
        
        return;
    }
    
    selectedSheet = sheetName;
    createDynamicFilters(sheetName);
    processSheetData(sheetName);
}

function processSheetData(sheetName) {
    console.log(`🔄 processSheetData iniciado para: ${sheetName}`);
    
    const data = allSheetsData[sheetName];
    
    if (!data || data.length < 2) {
        alert('Dados insuficientes na aba selecionada');
        return;
    }
    
    const headers = data[0];
    const rows = data.slice(1);
    
    const dateColumns = {};
    headers.forEach((header, index) => {
        if (typeof header === 'string' && header.match(/\d{2}\/\d{2}\/\d{4}/)) {
            const match = header.match(/(\d{2}\/\d{2}\/\d{4})\s+(Telecom|Vogel|Consolidado|Somado)/i);
            if (match) {
                const date = match[1];
                const entity = match[2];
                
                if (!dateColumns[date]) {
                    dateColumns[date] = {};
                }
                const normalizedEntity = entity.toLowerCase() === 'somado' ? 'consolidado' : entity.toLowerCase();
                dateColumns[date][normalizedEntity] = index;
            }
        }
    });
    
    periods = Object.keys(dateColumns).sort((a, b) => {
        try {
            const parseDate = (dateStr) => {
                const [day, month, year] = dateStr.split('/');
                return new Date(year, month - 1, day);
            };
            
            const dateA = parseDate(a);
            const dateB = parseDate(b);
            
            return dateA.getTime() - dateB.getTime();
        } catch (error) {
            console.error('Error parsing dates:', error);
            return 0;
        }
    });
    
    rawData = rows.map(row => {
        const item = {
            values: {}
        };
        
        headers.forEach((header, index) => {
            if (typeof header === 'string' && !header.match(/\d{2}\/\d{2}\/\d{4}/)) {
                item[header] = row[index];
            }
        });
        
        periods.forEach(period => {
            if (!dateColumns[period]) return;
            
            item.values[period] = {
                telecom: parseFloat(row[dateColumns[period].telecom]) || 0,
                vogel: parseFloat(row[dateColumns[period].vogel]) || 0,
                consolidado: parseFloat(row[dateColumns[period].consolidado]) || 0
            };
        });
        
        return item;
    });
    
    console.log(`📊 Processados ${rawData.length} registros`);
    
    const accountSearchInput = document.getElementById('filter_contaSearch');
    if (accountSearchInput && accountSearchInput._generateAccountOptions) {
        accountSearchInput._generateAccountOptions();
    }
    
    setupAccountSearchIcon();
    
    populateFilterOptions();
    updatePeriodSelectors();
    autoSelectLastTwoPeriods();
    
    filteredData = [...rawData];
    updateDashboard();
}

function autoSelectLastTwoPeriods() {
    if (periods.length >= 2) {
        const lastTwoPeriods = periods.slice(-2);
        
        const previousPeriodSelect = document.getElementById('previousPeriod');
        const currentPeriodSelect = document.getElementById('currentPeriod');
        
        if (previousPeriodSelect && lastTwoPeriods[0]) {
            previousPeriodSelect.value = lastTwoPeriods[0];
        }
        
        if (currentPeriodSelect && lastTwoPeriods[1]) {
            currentPeriodSelect.value = lastTwoPeriods[1];
        }
    }
}

function updatePeriodSelectors() {
    const currentPeriodSelect = document.getElementById('currentPeriod');
    const previousPeriodSelect = document.getElementById('previousPeriod');
    
    if (!currentPeriodSelect || !previousPeriodSelect) return;
    
    currentPeriodSelect.innerHTML = '<option value="">Selecione...</option>';
    previousPeriodSelect.innerHTML = '<option value="">Selecione...</option>';
    
    periods.forEach(period => {
        const formattedPeriod = formatPeriod(period);
        
        const option1 = document.createElement('option');
        option1.value = period;
        option1.textContent = formattedPeriod;
        currentPeriodSelect.appendChild(option1);
        
        const option2 = document.createElement('option');
        option2.value = period;
        option2.textContent = formattedPeriod;
        previousPeriodSelect.appendChild(option2);
    });
}

// ============================================
// SISTEMA DE FILTROS DINÂMICOS
// ============================================

function createDynamicFilters(sheetName) {
    console.log(`🔧 Criando filtros dinâmicos para aba: ${sheetName}`);
    
    const dynamicFiltersContainer = document.getElementById('dynamicFilters');
    dynamicFiltersContainer.innerHTML = '';
    
    const filterConfig = SHEET_FILTERS[sheetName] || [
        { id: 'contaSearch', label: 'Buscar Conta', field: 'CONTA_SEARCH', type: 'search' },
        { id: 'fsGroupReport', label: 'FS Group REPORT', field: 'FS Group REPORT' },
        { id: 'fsGroup', label: 'FS Group', field: 'FS Group' },
        { id: 'sintetica', label: 'SINTÉTICA', field: 'SINTETICA' }
    ];
    
    filterConfig.forEach(config => {
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-group';
        
        if (config.type === 'search') {
            filterGroup.innerHTML = `
                <label for="filter_${config.id}">${config.label}</label>
                <div class="search-container">
                    <div class="search-input-wrapper">
                        <i class="fas fa-search"></i>
                        <input type="text" 
                               id="filter_${config.id}" 
                               data-field="${config.field}"
                               placeholder="Digite número da conta ou descrição..."
                               autocomplete="off">
                        <button type="button" class="clear-search" id="clear_${config.id}" style="display: none;">
                            <i class="fas fa-times"></i>
                        </button>
                    </div>
                    <div class="search-suggestions" id="suggestions_${config.id}"></div>
                </div>
            `;
        } else {
            filterGroup.innerHTML = `
                <label for="filter_${config.id}">${config.label}</label>
                <div class="select-wrapper">
                    <select id="filter_${config.id}" data-field="${config.field}">
                        <option value="">Todos</option>
                    </select>
                    <i class="fas fa-chevron-down"></i>
                </div>
            `;
        }
        
        dynamicFiltersContainer.appendChild(filterGroup);
    });
    
    // Adicionar event listeners
    filterConfig.forEach(config => {
        if (config.type === 'search') {
            setTimeout(() => {
                setupAccountSearch(config.id);
            }, 100);
        } else {
            setTimeout(() => {
                const select = document.getElementById(`filter_${config.id}`);
                if (select) {
                    select.addEventListener('change', () => updateCascadingFilters());
                }
            }, 100);
        }
    });
}

// NOVA FUNÇÃO: Detectar se deve mostrar formato expandido
function shouldShowExpandedFormat() {
    const accountInput = document.getElementById('filter_contaSearch');
    const hasSpecificAccount = accountInput && (
        accountInput.dataset.selectedConta || 
        (accountInput.value && accountInput.value.trim().length > 0)
    );
    
    return hasSpecificAccount;
}

// CORRIGIDO: Sistema de busca por conta otimizado para NÚMERO
function setupAccountSearch(filterId) {
    const input = document.getElementById(`filter_${filterId}`);
    const clearBtn = document.getElementById(`clear_${filterId}`);
    const suggestions = document.getElementById(`suggestions_${filterId}`);
    
    if (!input || !clearBtn || !suggestions) return;
    
    let accountOptions = [];
    
    function generateAccountOptions() {
        accountOptions = [];
        const seen = new Set();
        
        console.log('🔍 === GERANDO OPÇÕES DE CONTA ===');
        
        rawData.forEach((item, index) => {
            const { contaSAP, description } = extractAccountInfo(item);
            
            if (contaSAP || description) {
                // PRIORIDADE: Se tem número da conta, usar ele como chave principal
                const primaryKey = contaSAP || description;
                const displayText = contaSAP && description ? 
                    `${contaSAP} - ${description}` : 
                    contaSAP || description;
                
                const searchKey = `${contaSAP}|${description}`.toLowerCase();
                
                if (!seen.has(searchKey)) {
                    seen.add(searchKey);
                    accountOptions.push({
                        contaSAP: contaSAP,
                        description: description,
                        displayText: displayText,
                        searchText: searchKey,
                        primaryKey: primaryKey // Para busca prioritária
                    });
                }
            }
            
            // Log dos primeiros registros para debug
            if (index < 3) {
                console.log(`🔍 Item ${index}: ContaSAP="${contaSAP}", Desc="${description?.substring(0, 30)}..."`);
            }
        });
        
        // ORDENAÇÃO: Priorizar por número da conta
        accountOptions.sort((a, b) => {
            // Se ambos têm conta SAP, ordenar numericamente
            if (a.contaSAP && b.contaSAP) {
                const contaA = String(a.contaSAP);
                const contaB = String(b.contaSAP);
                const numA = parseInt(contaA);
                const numB = parseInt(contaB);
                if (!isNaN(numA) && !isNaN(numB)) {
                    return numA - numB;
                }
                return contaA.localeCompare(contaB);
            }
            // Se apenas um tem conta SAP, priorizá-lo
            if (a.contaSAP && !b.contaSAP) return -1;
            if (!a.contaSAP && b.contaSAP) return 1;
            // Se nenhum tem conta SAP, ordenar por descrição
            return a.displayText.localeCompare(b.displayText);
        });
        
        console.log(`✅ ${accountOptions.length} opções de conta geradas`);
        if (accountOptions.length > 0) {
            console.log(`🔍 Primeira opção: ${accountOptions[0].displayText}`);
            console.log(`🔍 Última opção: ${accountOptions[accountOptions.length - 1].displayText}`);
        }
    }
    
    function showSuggestions(query) {
        if (!query || query.length < 1) {
            suggestions.innerHTML = '';
            suggestions.style.display = 'none';
            return;
        }
        
        const queryLower = query.toLowerCase();
        
        // Verificar se a busca é numérica (conta) ou texto (descrição)
        const isNumericSearch = /^\d+/.test(query);
        
        let filtered;
        
        if (isNumericSearch) {
            // BUSCA NUMÉRICA: Apenas no número da conta
            filtered = accountOptions.filter(option => {
                return option.contaSAP && option.contaSAP.toString().toLowerCase().includes(queryLower);
            }).sort((a, b) => {
                // Priorizar matches que começam com a query
                const aStartsWith = a.contaSAP && a.contaSAP.toString().toLowerCase().startsWith(queryLower);
                const bStartsWith = b.contaSAP && b.contaSAP.toString().toLowerCase().startsWith(queryLower);
                
                if (aStartsWith && !bStartsWith) return -1;
                if (!aStartsWith && bStartsWith) return 1;
                
                // Ordenar numericamente
                if (aStartsWith && bStartsWith) {
                    const numA = parseInt(a.contaSAP);
                    const numB = parseInt(b.contaSAP);
                    if (!isNaN(numA) && !isNaN(numB)) {
                        return numA - numB;
                    }
                }
                
                return 0;
            });
        } else {
            // BUSCA POR TEXTO: Na descrição e conta
            filtered = accountOptions.filter(option => {
                const contaMatch = option.contaSAP && option.contaSAP.toString().toLowerCase().includes(queryLower);
                const descMatch = option.description && option.description.toLowerCase().includes(queryLower);
                return contaMatch || descMatch;
            }).sort((a, b) => {
                // Priorizar matches na conta primeiro
                const aContaMatch = a.contaSAP && a.contaSAP.toString().toLowerCase().includes(queryLower);
                const bContaMatch = b.contaSAP && b.contaSAP.toString().toLowerCase().includes(queryLower);
                
                if (aContaMatch && !bContaMatch) return -1;
                if (!aContaMatch && bContaMatch) return 1;
                
                return 0;
            });
        }
        
        filtered = filtered.slice(0, 15);
        
        if (filtered.length === 0) {
            suggestions.innerHTML = `<div class="suggestion-item">
                <div class="suggestion-description">
                    ${isNumericSearch ? 
                        `Nenhuma conta encontrada com número "${query}"` : 
                        `Nenhuma conta encontrada com "${query}"`
                    }
                </div>
            </div>`;
            suggestions.style.display = 'block';
            return;
        }
        
        suggestions.innerHTML = filtered.map(option => `
            <div class="suggestion-item" data-conta="${option.contaSAP || ''}" data-description="${option.description || ''}">
                <div class="suggestion-conta" style="font-weight: bold; color: #0066cc;">
                    ${option.contaSAP || 'S/N'}
                </div>
                <div class="suggestion-description" style="font-size: 0.85em; color: #666;">
                    ${option.description || 'Sem descrição'}
                </div>
            </div>
        `).join('');
        
        suggestions.style.display = 'block';
        
        // Event listeners para as sugestões
        suggestions.querySelectorAll('.suggestion-item').forEach(item => {
            item.addEventListener('click', () => {
                const conta = item.dataset.conta;
                const description = item.dataset.description;
                
                // IMPORTANTE: Usar apenas o número da conta como valor
                input.value = conta || description;
                input.dataset.selectedConta = conta;
                input.dataset.selectedDescription = description;
                
                suggestions.style.display = 'none';
                clearBtn.style.display = 'inline-block';
                
                console.log(`🎯 Conta selecionada: ${conta} (descrição: ${description})`);
                updateCascadingFilters();
                
                setTimeout(() => {
                    applyFiltersWithDebounce();
                }, 150);
            });
        });
    }
    
    input.addEventListener('input', (e) => {
        const value = e.target.value;
        
        if (value) {
            clearBtn.style.display = 'inline-block';
            showSuggestions(value);
        } else {
            clearBtn.style.display = 'none';
            suggestions.style.display = 'none';
            delete input.dataset.selectedConta;
            delete input.dataset.selectedDescription;
            updateCascadingFilters();
        }
    });
    
    input.addEventListener('focus', () => {
        if (input.value) {
            showSuggestions(input.value);
        }
    });
    
    input.addEventListener('blur', (e) => {
        setTimeout(() => {
            suggestions.style.display = 'none';
        }, 200);
    });
    
    clearBtn.addEventListener('click', () => {
        input.value = '';
        clearBtn.style.display = 'none';
        suggestions.style.display = 'none';
        delete input.dataset.selectedConta;
        delete input.dataset.selectedDescription;
        updateCascadingFilters();
        console.log('🧹 Busca de conta limpa');
    });
    
    if (rawData.length > 0) {
        generateAccountOptions();
    }
    
    input._generateAccountOptions = generateAccountOptions;
}

function populateFilterOptions() {
    const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
    
    dynamicFilters.forEach(select => {
        const field = select.dataset.field;
        
        const currentFilters = {};
        
        dynamicFilters.forEach(filterSelect => {
            if (filterSelect.value) {
                currentFilters[filterSelect.dataset.field] = filterSelect.value;
            }
        });
        
        const contaInput = document.getElementById('filter_contaSearch');
        if (contaInput && contaInput.value) {
            currentFilters['CONTA_SEARCH'] = {
                value: contaInput.value,
                conta: contaInput.dataset.selectedConta,
                description: contaInput.dataset.selectedDescription
            };
        }
        
        let filteredDataForOptions = rawData;
        
        Object.keys(currentFilters).forEach(currentField => {
            if (currentField !== field) {
                if (currentField === 'CONTA_SEARCH') {
                    const searchFilter = currentFilters[currentField];
                    filteredDataForOptions = filteredDataForOptions.filter(item => {
                        const { contaSAP, description } = extractAccountInfo(item);
                        
                        if (searchFilter.conta || searchFilter.description) {
                            return (searchFilter.conta && String(searchFilter.conta) === String(contaSAP)) || 
                                   (searchFilter.description && searchFilter.description === description);
                        }
                        
                        const searchText = searchFilter.value.toLowerCase();
                        return String(contaSAP).toLowerCase().includes(searchText) || 
                               String(description).toLowerCase().includes(searchText);
                    });
                } else {
                    filteredDataForOptions = filteredDataForOptions.filter(item => 
                        item[currentField] === currentFilters[currentField]
                    );
                }
            }
        });
        
        const uniqueValues = [...new Set(filteredDataForOptions.map(item => item[field]).filter(val => val))];
        
        const currentValue = select.value;
        
        select.innerHTML = '<option value="">Todos</option>';
        
        uniqueValues.sort().forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        
        if (currentValue && uniqueValues.includes(currentValue)) {
            select.value = currentValue;
        }
    });
}

function updateCascadingFilters() {
    populateFilterOptions();
    
    const accountSearchInput = document.getElementById('filter_contaSearch');
    if (accountSearchInput && accountSearchInput._generateAccountOptions) {
        accountSearchInput._generateAccountOptions();
    }
    
    // Atualizar título do gráfico quando filtros mudarem
    if (typeof updateChartTitle === 'function') {
        updateChartTitle();
    }
}

function setupAccountSearchIcon() {
    setTimeout(() => {
        const accountSearchInput = document.getElementById('filter_contaSearch');
        if (!accountSearchInput) {
            setTimeout(setupAccountSearchIcon, 500);
            return;
        }
        
        const searchContainer = accountSearchInput.parentElement;
        const searchIcon = searchContainer.querySelector('.fas.fa-search');
        
        if (!searchIcon) return;
        
        searchIcon.style.cursor = 'pointer';
        searchIcon.style.pointerEvents = 'auto';
        searchIcon.title = 'Clique para ver todas as contas';
        
        const newSearchIcon = searchIcon.cloneNode(true);
        searchIcon.parentNode.replaceChild(newSearchIcon, searchIcon);
        
        newSearchIcon.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            
            if (!rawData || rawData.length === 0) {
                alert('Selecione uma aba primeiro!');
                return;
            }
            
            const accountOptions = generateAccountOptionsManually();
            
            if (accountOptions.length === 0) {
                alert('Nenhuma conta encontrada nesta aba.');
                return;
            }
            
            openAccountModalWithData(accountSearchInput, accountOptions);
        });
        
        newSearchIcon.addEventListener('mouseenter', () => {
            newSearchIcon.style.color = '#0066cc';
            newSearchIcon.style.backgroundColor = '#f0f8ff';
            newSearchIcon.style.borderRadius = '4px';
        });
        
        newSearchIcon.addEventListener('mouseleave', () => {
            newSearchIcon.style.color = '#6b7280';
            newSearchIcon.style.backgroundColor = 'transparent';
        });
        
    }, 100);
}

// CORRIGIDO: Gerar opções de conta priorizando número
function generateAccountOptionsManually() {
    const accountOptions = [];
    const seen = new Set();
    
    if (!rawData || rawData.length === 0) {
        return accountOptions;
    }
    
    const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
    const currentFilters = {};
    
    dynamicFilters.forEach(filterSelect => {
        if (filterSelect.value) {
            currentFilters[filterSelect.dataset.field] = filterSelect.value;
        }
    });
    
    let filteredDataForOptions = rawData;
    
    Object.keys(currentFilters).forEach(currentField => {
        if (currentField !== 'CONTA_SEARCH') {
            filteredDataForOptions = filteredDataForOptions.filter(item => 
                item[currentField] === currentFilters[currentField]
            );
        }
    });
    
    filteredDataForOptions.forEach(item => {
        const { contaSAP, description } = extractAccountInfo(item);
        
        if (contaSAP || description) {
            const displayText = contaSAP && description ? 
                `${contaSAP} - ${description}` : 
                contaSAP || description;
            
            const searchKey = `${contaSAP}|${description}`.toLowerCase();
            
            if (!seen.has(searchKey)) {
                seen.add(searchKey);
                accountOptions.push({
                    contaSAP: contaSAP,
                    description: description,
                    displayText: displayText,
                    searchText: searchKey
                });
            }
        }
    });
    
    // ORDENAÇÃO PRIORITÁRIA por número da conta
    accountOptions.sort((a, b) => {
        if (a.contaSAP && b.contaSAP) {
            const contaA = String(a.contaSAP);
            const contaB = String(b.contaSAP);
            const numA = parseInt(contaA);
            const numB = parseInt(contaB);
            if (!isNaN(numA) && !isNaN(numB)) {
                return numA - numB;
            }
            return contaA.localeCompare(contaB);
        }
        if (a.contaSAP && !b.contaSAP) return -1;
        if (!a.contaSAP && b.contaSAP) return 1;
        return a.displayText.localeCompare(b.displayText);
    });
    
    return accountOptions;
}

// ============================================
// MODAL DE SELEÇÃO DE CONTAS
// ============================================

function openAccountModalWithData(targetInput, accountOptions) {
    console.log(`🔍 Abrindo modal com ${accountOptions.length} contas`);
    
    const modal = document.getElementById('accountModal');
    const searchInput = document.getElementById('accountSearch');
    const accountList = document.getElementById('accountList');
    
    if (!modal || !searchInput || !accountList) {
        console.error('❌ Elementos do modal não encontrados');
        return;
    }
    
    searchInput.value = '';
    
    function renderList(accounts = accountOptions) {
        accountList.innerHTML = '';
        
        if (accounts.length === 0) {
            accountList.innerHTML = '<div class="no-accounts">Nenhuma conta encontrada</div>';
            return;
        }
        
        accounts.forEach((account, index) => {
            const item = document.createElement('div');
            item.className = 'account-item';
            item.innerHTML = `
                <div class="account-number">${account.contaSAP || `Conta ${index + 1}`}</div>
                <div class="account-description">${account.description || 'Sem descrição'}</div>
            `;
            
            item.addEventListener('click', () => {
                targetInput.value = account.displayText;
                targetInput.dataset.selectedConta = account.contaSAP;
                targetInput.dataset.selectedDescription = account.description;
                
                const clearBtn = document.getElementById('clear_contaSearch');
                if (clearBtn) clearBtn.style.display = 'inline-block';
                
                modal.classList.remove('active');
                
                console.log(`🎯 Conta selecionada do modal: ${account.contaSAP} - ${account.description}`);
                updateCascadingFilters();
                
                // CORREÇÃO: Aplicar filtros com debounce após seleção
                setTimeout(() => {
                    applyFiltersWithDebounce();
                }, 100);
            });
            
            accountList.appendChild(item);
        });
    }
    
    const newSearchInput = searchInput.cloneNode(true);
    searchInput.parentNode.replaceChild(newSearchInput, searchInput);
    
    newSearchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        
        if (!query) {
            renderList();
            return;
        }
        
        // PRIORIZAR busca por número da conta
        const filtered = accountOptions.filter(account => {
            const contaMatch = account.contaSAP && account.contaSAP.toString().toLowerCase().includes(query);
            const descMatch = account.description && account.description.toLowerCase().includes(query);
            return contaMatch || descMatch;
        }).sort((a, b) => {
            // Priorizar matches que começam com a query no número da conta
            const aContaStartsWith = a.contaSAP && a.contaSAP.toString().toLowerCase().startsWith(query);
            const bContaStartsWith = b.contaSAP && b.contaSAP.toString().toLowerCase().startsWith(query);
            
            if (aContaStartsWith && !bContaStartsWith) return -1;
            if (!aContaStartsWith && bContaStartsWith) return 1;
            
            // Priorizar qualquer match no número da conta
            const aContaMatch = a.contaSAP && a.contaSAP.toString().toLowerCase().includes(query);
            const bContaMatch = b.contaSAP && b.contaSAP.toString().toLowerCase().includes(query);
            
            if (aContaMatch && !bContaMatch) return -1;
            if (!aContaMatch && bContaMatch) return 1;
            
            return 0;
        });
        
        renderList(filtered);
    });
    
    renderList();
    modal.classList.add('active');
    setTimeout(() => newSearchInput.focus(), 100);
}

function initializeAccountModal() {
    const modal = document.getElementById('accountModal');
    const closeBtn = document.getElementById('closeAccountModal');
    
    if (modal && closeBtn) {
        closeBtn.addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });
    }
}

// ============================================
// FUNÇÕES PARA MODAL DE COMPARAÇÕES
// ============================================

function populateComparisonPeriods() {
    if (!periods || periods.length === 0) return;
    
    const period1Select = document.getElementById('comparisonPeriod1');
    const period2Select = document.getElementById('comparisonPeriod2');
    
    // Limpar opções
    period1Select.innerHTML = '<option value="">Selecione o primeiro período...</option>';
    period2Select.innerHTML = '<option value="">Selecione o segundo período...</option>';
    
    // Mapear períodos baseado no tipo selecionado
    const periodMap = mapPeriodsByType(comparisonMode.type);
    const periodKeys = Object.keys(periodMap).sort();
    
    // Popular dropdowns
    periodKeys.forEach(key => {
        const option1 = document.createElement('option');
        option1.value = key;
        option1.textContent = key;
        period1Select.appendChild(option1);
        
        const option2 = document.createElement('option');
        option2.value = key;
        option2.textContent = key;
        period2Select.appendChild(option2);
    });
    
    console.log(`✅ Períodos populados para comparação ${comparisonMode.type}:`, periodKeys);
}

function clearComparisonSettings() {
    // Resetar tipo para mensal
    document.querySelectorAll('.comparison-type-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelector('[data-type="mensal"]').classList.add('active');
    comparisonMode.type = 'mensal';
    
    // Limpar seleções
    document.getElementById('comparisonPeriod1').value = '';
    document.getElementById('comparisonPeriod2').value = '';
    document.getElementById('comparisonEntityFilter').value = 'total';
    document.getElementById('comparisonMinVariation').value = '';
    document.getElementById('comparisonMaxResults').value = '250';
    
    // Resetar modo de comparação
    comparisonMode = {
        active: false,
        type: 'mensal',
        period1: null,
        period2: null,
        entityFilter: 'total',
        minVariation: null,
        maxResults: 250
    };
    
    // Esconder resumo
    document.getElementById('comparisonSummary').style.display = 'none';
    
    // Repopular períodos
    populateComparisonPeriods();
    
    // Voltar ao modo normal da tabela se estava em comparação
    if (comparisonMode.active) {
        resetTableToNormalMode();
    }
    
    console.log('🧹 Configurações de comparação limpas');
}

function resetTableToNormalMode() {
    // Resetar headers da tabela
    const table = document.getElementById('detailsTable');
    const headers = table.querySelectorAll('thead th');
    
    if (headers.length >= 7) {
        headers[0].textContent = 'Empresa';
        headers[1].textContent = 'Conta SAP';
        headers[2].textContent = 'Descrição';
        headers[3].textContent = 'Valor Mês Anterior';
        headers[4].textContent = 'Valor Mês Atual';
        headers[5].textContent = 'Variação';
        headers[6].textContent = 'Var %';
    }
    
    // Desativar modo de comparação
    comparisonMode.active = false;
    
    // Atualizar tabela com dados normais
    filterDetailTable();
}

// ============================================
// FUNÇÕES DE APLICAÇÃO DE FILTROS CORRIGIDAS
// ============================================

// NOVA FUNÇÃO: Aplicar filtros com debounce
function applyFiltersWithDebounce() {
    // Cancelar timeout anterior
    if (filterTimeout) {
        clearTimeout(filterTimeout);
    }
    
    // Agendar aplicação com debounce
    filterTimeout = setTimeout(() => {
        applyFilters();
    }, 200);
}

// FUNÇÃO MODIFICADA: Aplicar filtros com proteção DOM
function applyFilters() {
    console.log('🔄 === APLICANDO FILTROS (FILTRO EXATO POR CONTA) ===');
    
    const filterSelects = document.querySelectorAll('#dynamicFilters select');
    const filterInputs = document.querySelectorAll('#dynamicFilters input[type="text"]');
    const filters = {};
    
    filterSelects.forEach(select => {
        if (select.value) {
            filters[select.dataset.field] = select.value;
        }
    });
    
    filterInputs.forEach(input => {
        if (input.value && input.dataset.field === 'CONTA_SEARCH') {
            filters.CONTA_SEARCH = {
                value: input.value,
                conta: input.dataset.selectedConta,
                description: input.dataset.selectedDescription
            };
        }
    });
    
    console.log('🔍 Filtros ativos:', filters);
    
    // CORREÇÃO CRÍTICA: Filtro APENAS por número da conta
    filteredData = rawData.filter((item, index) => {
        const passes = Object.keys(filters).every(field => {
            if (field === 'CONTA_SEARCH') {
                const searchFilter = filters[field];
                const { contaSAP, description } = extractAccountInfo(item);
                
                // REGRA 1: Se há uma conta específica selecionada (do modal/sugestão)
                if (searchFilter.conta) {
                    const exactMatch = String(searchFilter.conta) === String(contaSAP);
                    
                    if (index < 3) {
                        console.log(`🎯 Filtro EXATO por conta: Item[${index}] ContaSAP="${contaSAP}" vs Filtro="${searchFilter.conta}" = ${exactMatch}`);
                    }
                    
                    return exactMatch;
                }
                
                // REGRA 2: Se é busca livre, PRIORIZAR APENAS o número da conta
                const searchText = searchFilter.value.toLowerCase().trim();
                
                // Verificar se o texto digitado é um número (conta)
                const isNumericSearch = /^\d+/.test(searchText);
                
                if (isNumericSearch) {
                    // Se é numérico, buscar APENAS no número da conta
                    const contaMatch = contaSAP && String(contaSAP).toLowerCase().includes(searchText);
                    
                    if (index < 3) {
                        console.log(`🔢 Busca NUMÉRICA: Item[${index}] ContaSAP="${contaSAP}" Query="${searchText}" Match=${contaMatch}`);
                    }
                    
                    return contaMatch;
                } else {
                    // Se não é numérico, buscar na descrição também
                    const contaMatch = contaSAP && String(contaSAP).toLowerCase().includes(searchText);
                    const descMatch = description && String(description).toLowerCase().includes(searchText);
                    
                    if (index < 3) {
                        console.log(`🔍 Busca TEXTO: Item[${index}] ContaSAP="${contaSAP}" Desc="${description?.substring(0, 30)}..." Query="${searchText}" ContaMatch=${contaMatch} DescMatch=${descMatch}`);
                    }
                    
                    return contaMatch || descMatch;
                }
            } else {
                // Outros filtros normais
                return item[field] === filters[field];
            }
        });
        
        return passes;
    });
    
    console.log(`✅ RESULTADO FILTROS: ${rawData.length} → ${filteredData.length} registros`);
    
    // Log detalhado dos registros filtrados
    if (filteredData.length > 0) {
        console.log('🔍 Registros filtrados (primeiros 5):');
        filteredData.slice(0, 5).forEach((item, index) => {
            const { contaSAP, description } = extractAccountInfo(item);
            console.log(`  [${index}] ContaSAP: "${contaSAP}", Desc: "${description?.substring(0, 50)}..."`);
        });
    }
    
    // Verificar se estamos em modo de comparação
    if (comparisonMode.active) {
        const periodMap = mapPeriodsByType(comparisonMode.type);
        const periods1 = periodMap[comparisonMode.period1];
        const periods2 = periodMap[comparisonMode.period2];
        
        if (periods1 && periods2) {
            updateDetailTableWithComparison(periods1, periods2);
            showComparisonSummary(periods1, periods2);
        }
    } else {
        updateDashboard();
    }
    
    updateSummaryCardsAndCharts();
    updateChartTitle();
    
    setTimeout(() => {
        populateFilterOptions();
    }, 100);
}

// FUNÇÃO MODIFICADA: clearFilters com limpeza de gráficos
function clearFilters() {
    console.log('🧹 === LIMPANDO TODOS OS FILTROS ===');
    
    // CORREÇÃO: Forçar limpeza de gráficos primeiro
    if (typeof forceChartCleanup === 'function') {
        forceChartCleanup();
    }
    
    document.getElementById('sheetSelector').value = '';
    
    const filterSelects = document.querySelectorAll('#dynamicFilters select');
    filterSelects.forEach(select => {
        select.value = '';
    });
    
    const filterInputs = document.querySelectorAll('#dynamicFilters input[type="text"]');
    filterInputs.forEach(input => {
        input.value = '';
        const clearBtn = document.getElementById(input.id.replace('filter_', 'clear_'));
        if (clearBtn) clearBtn.style.display = 'none';
        
        const suggestions = document.getElementById(input.id.replace('filter_', 'suggestions_'));
        if (suggestions) suggestions.style.display = 'none';
        
        delete input.dataset.selectedConta;
        delete input.dataset.selectedDescription;
    });
    
    document.getElementById('dynamicFilters').innerHTML = '';
    document.getElementById('currentPeriod').innerHTML = '<option value="">Selecione...</option>';
    document.getElementById('previousPeriod').innerHTML = '<option value="">Selecione...</option>';
    
    chartViewMode = 'monthly';
    const monthlyBtn = document.getElementById('monthlyBtn');
    const quarterlyBtn = document.getElementById('quarterlyBtn');
    const chartTitle = document.getElementById('chartTitle');
    monthlyBtn.classList.add('active');
    quarterlyBtn.classList.remove('active');
    chartTitle.textContent = 'Evolução Mensal';
    
    // Reset comparison mode
    comparisonMode = {
        active: false,
        type: 'mensal',
        period1: null,
        period2: null,
        entityFilter: 'total',
        minVariation: null,
        maxResults: 250
    };
    
    resetTableToNormalMode();
    
    rawData = [];
    filteredData = [];
    periods = [];
    selectedSheet = '';
    
    updateSummaryCard('telecom', 0, 0, 0);
    updateSummaryCard('vogel', 0, 0, 0);
    updateSummaryCard('consolidado', 0, 0, 0);
    
    const tbody = document.querySelector('#detailsTable tbody');
    if (tbody) {
        tbody.innerHTML = '';
    }
    
    if (charts.evolution) {
        charts.evolution.destroy();
        charts.evolution = null;
    }
    
    console.log('🧹 Todos os filtros limpos');
}

// ============================================
// INICIALIZAÇÃO
// ============================================

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    showUploadModal(true);
    updateCurrentDate();
});

// Initialize Event Listeners
function initializeEventListeners() {
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);
    
    // CORRIGIDO: Botão aplicar agora usa versão com debounce
    const applyBtn = document.getElementById('applyFilters');
    if (applyBtn) {
        applyBtn.addEventListener('click', () => {
            console.log('🔵 BOTÃO APLICAR CLICADO!');
            applyFiltersWithDebounce(); // CORREÇÃO: Usar versão com debounce
        });
    }
    
    document.getElementById('clearFilters').addEventListener('click', () => clearFilters());
    document.getElementById('currentPeriod').addEventListener('change', () => updateDashboard());
    document.getElementById('previousPeriod').addEventListener('change', () => updateDashboard());
    document.getElementById('sheetSelector').addEventListener('change', handleSheetChange);
    document.getElementById('monthlyBtn').addEventListener('click', () => toggleChartView('monthly'));
    document.getElementById('quarterlyBtn').addEventListener('click', () => toggleChartView('quarterly'));
    document.getElementById('uploadButton').addEventListener('click', () => showUploadModal(true));
    document.getElementById('closeModal').addEventListener('click', () => showUploadModal(false));
    document.getElementById('descriptionFilter').addEventListener('input', () => {
        if (comparisonMode.active) {
            // Se estamos em modo de comparação, re-aplicar comparação
            const periodMap = mapPeriodsByType(comparisonMode.type);
            const periods1 = periodMap[comparisonMode.period1];
            const periods2 = periodMap[comparisonMode.period2];
            if (periods1 && periods2) {
                updateDetailTableWithComparison(periods1, periods2);
            }
        } else {
            filterDetailTable();
        }
    });
    document.getElementById('sortFilter').addEventListener('change', () => {
        if (comparisonMode.active) {
            // Se estamos em modo de comparação, re-aplicar comparação
            const periodMap = mapPeriodsByType(comparisonMode.type);
            const periods1 = periodMap[comparisonMode.period1];
            const periods2 = periodMap[comparisonMode.period2];
            if (periods1 && periods2) {
                updateDetailTableWithComparison(periods1, periods2);
            }
        } else {
            filterDetailTable();
        }
    });
    document.getElementById('entityFilter').addEventListener('change', () => {
        if (comparisonMode.active) {
            // Se estamos em modo de comparação, re-aplicar comparação
            const periodMap = mapPeriodsByType(comparisonMode.type);
            const periods1 = periodMap[comparisonMode.period1];
            const periods2 = periodMap[comparisonMode.period2];
            if (periods1 && periods2) {
                updateDetailTableWithComparison(periods1, periods2);
            }
        } else {
            filterDetailTable();
        }
    });
    document.getElementById('exportTable').addEventListener('click', exportTableToExcel);
    document.getElementById('downloadChart').addEventListener('click', downloadChart);
    
    // Adicionar listener para o novo botão de export do gráfico (se existir)
    const exportChartBtn = document.getElementById('exportChartToExcel');
    if (exportChartBtn) {
        exportChartBtn.addEventListener('click', exportChartToExcel);
    }
    
    // ============================================
    // EVENT LISTENERS PARA COMPARAÇÕES
    // ============================================
    
    // Botão para abrir modal de comparações
    document.getElementById('openComparisonModal').addEventListener('click', () => {
        if (!selectedSheet || !rawData || rawData.length === 0) {
            alert('Selecione uma aba e carregue dados antes de fazer comparações');
            return;
        }
        showComparisonModal(true);
    });
    
    // Fechar modal de comparações
    document.getElementById('closeComparisonModal').addEventListener('click', () => {
        showComparisonModal(false);
    });
    
    // Botões de tipo de comparação
    document.querySelectorAll('.comparison-type-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.comparison-type-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            comparisonMode.type = btn.dataset.type;
            populateComparisonPeriods();
        });
    });
    
    // Seleção de períodos para comparação
    document.getElementById('comparisonPeriod1').addEventListener('change', (e) => {
        comparisonMode.period1 = e.target.value;
    });
    
    document.getElementById('comparisonPeriod2').addEventListener('change', (e) => {
        comparisonMode.period2 = e.target.value;
    });
    
    // Filtros de comparação
    document.getElementById('comparisonEntityFilter').addEventListener('change', (e) => {
        comparisonMode.entityFilter = e.target.value;
    });
    
    document.getElementById('comparisonMinVariation').addEventListener('input', (e) => {
        comparisonMode.minVariation = e.target.value ? parseFloat(e.target.value) : null;
    });
    
    document.getElementById('comparisonMaxResults').addEventListener('change', (e) => {
        comparisonMode.maxResults = e.target.value;
    });
    
    // Comparações rápidas
    document.querySelectorAll('.quick-comparison-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const quickType = btn.dataset.quick;
            applyQuickComparison(quickType);
        });
    });
    
    // Botões de ação do modal
    document.getElementById('clearComparisonBtn').addEventListener('click', () => {
        clearComparisonSettings();
    });
    
    document.getElementById('applyComparisonBtn').addEventListener('click', () => {
        applyPeriodComparison();
    });
    
    // Inicializar modais
    initializeAccountModal();
    
    window.addEventListener('click', function(e) {
        const uploadModal = document.getElementById('uploadModal');
        const comparisonModal = document.getElementById('comparisonModal');
        
        if (e.target === uploadModal) {
            showUploadModal(false);
        }
        if (e.target === comparisonModal) {
            showComparisonModal(false);
        }
    });
    
    window.addEventListener('resize', handleResize);
    handleResize();
}