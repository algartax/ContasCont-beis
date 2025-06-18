let workbook = null;
let allSheetsData = {};
let selectedSheet = '';
let rawData = [];
let filteredData = [];
let periods = [];
let charts = {};
let chartViewMode = 'monthly';

// Variáveis para o sistema de comparações
let comparisonMode = {
    active: false,
    type: 'mensal', // mensal, trimestral, semestral, anual
    period1: null,
    period2: null,
    entityFilter: 'total',
    minVariation: null,
    maxResults: 250
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
    
    if (headers.length >= 6) {
        headers[2].textContent = period2Label; // Período anterior
        headers[3].textContent = period1Label; // Período atual
        headers[4].textContent = 'Variação';
        headers[5].textContent = 'Var %';
    }
}

// Atualizar tabela com dados de comparação
function updateDetailTableWithComparison(periods1, periods2) {
    const tbody = document.querySelector('#detailsTable tbody');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    // Calcular dados para cada conta
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
    
    // Filtrar por descrição se houver filtro ativo
    const descriptionFilter = document.getElementById('descriptionFilter').value.toLowerCase();
    if (descriptionFilter) {
        filteredComparison = filteredComparison.filter(item => 
            item.description.toLowerCase().includes(descriptionFilter) || 
            item.contaSAP.toString().toLowerCase().includes(descriptionFilter)
        );
    }
    
    // Renderizar resultados
    if (filteredComparison.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="6" class="no-data">Nenhuma conta encontrada com os critérios de comparação especificados.</td>`;
        tbody.appendChild(row);
        return;
    }
    
    filteredComparison.forEach(item => {
        const row = document.createElement('tr');
        
        row.innerHTML = `
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
// FUNÇÕES UTILITÁRIAS ORIGINAIS
// ============================================

// Utility Functions
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

// Extrair informações da conta
function extractAccountInfo(item) {
    let contaSAP = '';
    let description = '';
    
    for (const key of Object.keys(item)) {
        const keyTrimmed = key.trim().replace(/\s+/g, ' ');
        const keyNoSpaces = key.replace(/\s/g, '');
        
        if (keyTrimmed === 'Conta SAP' || keyTrimmed === 'CONTA SAP' || 
            keyTrimmed === 'Número Conta' || keyTrimmed === 'NÚMERO CONTA' ||
            keyNoSpaces.toLowerCase() === 'contasap' || 
            keyNoSpaces.toLowerCase() === 'numeroconta') {
            contaSAP = String(item[key] || '');
        }
        
        if (keyTrimmed === 'Descrição Conta' || keyTrimmed === 'Descrição DF' ||
            keyTrimmed === 'DESCRIÇÃO CONTA' || keyTrimmed === 'DESCRIÇÃO DF') {
            description = String(item[key] || '');
        }
    }
    
    return { contaSAP, description };
}

// Group periods by quarter
function groupPeriodsByQuarter(periods) {
    const quarterGroups = {};
    
    periods.forEach(period => {
        const quarterKey = formatQuarter(period);
        if (!quarterGroups[quarterKey]) {
            quarterGroups[quarterKey] = [];
        }
        quarterGroups[quarterKey].push(period);
    });
    
    return quarterGroups;
}

// Calculate quarterly totals
function calculateQuarterlyData() {
    const quarterGroups = groupPeriodsByQuarter(periods);
    const quarterlyData = {};
    
    Object.keys(quarterGroups).forEach(quarter => {
        const periodsInQuarter = quarterGroups[quarter];
        quarterlyData[quarter] = {
            telecom: 0,
            vogel: 0,
            consolidado: 0
        };
        
        periodsInQuarter.forEach(period => {
            quarterlyData[quarter].telecom += calculateTotal('telecom', period);
            quarterlyData[quarter].vogel += calculateTotal('vogel', period);
            quarterlyData[quarter].consolidado += calculateTotal('consolidado', period);
        });
    });
    
    return quarterlyData;
}

// UI Control Functions
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

function toggleChartView(mode) {
    chartViewMode = mode;
    
    const monthlyBtn = document.getElementById('monthlyBtn');
    const quarterlyBtn = document.getElementById('quarterlyBtn');
    const chartTitle = document.getElementById('chartTitle');
    
    if (mode === 'monthly') {
        monthlyBtn.classList.add('active');
        quarterlyBtn.classList.remove('active');
        chartTitle.textContent = 'Evolução Mensal';
    } else {
        monthlyBtn.classList.remove('active');
        quarterlyBtn.classList.add('active');
        chartTitle.textContent = 'Evolução Trimestral';
    }
    
    updateEvolutionChart();
}

// Initialize
document.addEventListener('DOMContentLoaded', function() {
    initializeEventListeners();
    showUploadModal(true);
    updateCurrentDate();
});

// Initialize Event Listeners
function initializeEventListeners() {
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);
    
    // Botões originais
    const applyBtn = document.getElementById('applyFilters');
    if (applyBtn) {
        applyBtn.addEventListener('click', () => {
            console.log('🔵 BOTÃO APLICAR CLICADO!');
            applyFilters();
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
            // ✅ CORREÇÃO: Se estamos em modo de comparação, re-aplicar comparação
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
            // ✅ CORREÇÃO: Se estamos em modo de comparação, re-aplicar comparação
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
            // ✅ CORREÇÃO: Se estamos em modo de comparação, re-aplicar comparação
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
    
    if (headers.length >= 6) {
        headers[2].textContent = 'Valor Anterior';
        headers[3].textContent = 'Valor Atual';
        headers[4].textContent = 'Variação';
        headers[5].textContent = 'Var %';
    }
    
    // Desativar modo de comparação
    comparisonMode.active = false;
    
    // Atualizar tabela com dados normais
    filterDetailTable();
}

function updateCurrentDate() {
    const dateDisplay = document.getElementById('dateDisplay');
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    
    dateDisplay.textContent = `Data de análise: ${day}/${month}/${year}`;
}

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
    
    if (charts.evolution) {
        charts.evolution.render();
    }
}

// File Upload Handler
async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    showLoading(true);
    
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

function handleSheetChange() {
    const sheetSelector = document.getElementById('sheetSelector');
    const sheetName = sheetSelector.value;
    
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

function setupAccountSearch(filterId) {
    const input = document.getElementById(`filter_${filterId}`);
    const clearBtn = document.getElementById(`clear_${filterId}`);
    const suggestions = document.getElementById(`suggestions_${filterId}`);
    
    if (!input || !clearBtn || !suggestions) return;
    
    let accountOptions = [];
    
    function generateAccountOptions() {
        accountOptions = [];
        const seen = new Set();
        
        rawData.forEach(item => {
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
            return a.displayText.localeCompare(b.displayText);
        });
    }
    
    function showSuggestions(query) {
        if (!query || query.length < 2) {
            suggestions.innerHTML = '';
            suggestions.style.display = 'none';
            return;
        }
        
        const filtered = accountOptions.filter(option => 
            option.searchText.includes(query.toLowerCase())
        ).slice(0, 10);
        
        if (filtered.length === 0) {
            suggestions.innerHTML = '';
            suggestions.style.display = 'none';
            return;
        }
        
        suggestions.innerHTML = filtered.map(option => `
            <div class="suggestion-item" data-conta="${option.contaSAP}" data-description="${option.description}">
                <div class="suggestion-conta">${option.contaSAP}</div>
                <div class="suggestion-description">${option.description}</div>
            </div>
        `).join('');
        
        suggestions.style.display = 'block';
        
        suggestions.querySelectorAll('.suggestion-item').forEach(item => {
            item.addEventListener('click', () => {
                const conta = item.dataset.conta;
                const description = item.dataset.description;
                const displayText = conta && description ? `${conta} - ${description}` : conta || description;
                
                input.value = displayText;
                input.dataset.selectedConta = conta;
                input.dataset.selectedDescription = description;
                
                suggestions.style.display = 'none';
                clearBtn.style.display = 'inline-block';
                
                updateCascadingFilters();
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
    });
    
    if (rawData.length > 0) {
        generateAccountOptions();
    }
    
    input._generateAccountOptions = generateAccountOptions;
}

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
                
                updateCascadingFilters();
                setTimeout(() => {
                    applyFilters();
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
        
        const filtered = accountOptions.filter(account => 
            (account.searchText && account.searchText.includes(query)) ||
            (account.contaSAP && String(account.contaSAP).toLowerCase().includes(query)) ||
            (account.description && account.description.toLowerCase().includes(query))
        );
        
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
        return a.displayText.localeCompare(b.displayText);
    });
    
    return accountOptions;
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
}

// ✅ CORREÇÃO PRINCIPAL: Função applyFilters atualizada
function applyFilters() {
    console.log('🔄 === APLICANDO FILTROS ===');
    
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
    
    filteredData = rawData.filter((item, index) => {
        const passes = Object.keys(filters).every(field => {
            if (field === 'CONTA_SEARCH') {
                const searchFilter = filters[field];
                const { contaSAP, description } = extractAccountInfo(item);
                
                if (searchFilter.conta || searchFilter.description) {
                    const match = (searchFilter.conta && String(searchFilter.conta) === String(contaSAP)) || 
                                  (searchFilter.description && searchFilter.description === description);
                    return match;
                }
                
                const searchText = searchFilter.value.toLowerCase();
                const textMatch = String(contaSAP).toLowerCase().includes(searchText) || 
                                 String(description).toLowerCase().includes(searchText);
                return textMatch;
            } else {
                return item[field] === filters[field];
            }
        });
        
        return passes;
    });
    
    console.log(`✅ RESULTADO: ${rawData.length} → ${filteredData.length} registros após filtros`);
    
    // ✅ CORREÇÃO: Verificar se estamos em modo de comparação
    if (comparisonMode.active) {
        // Se estamos em modo de comparação, re-aplicar comparação com dados filtrados
        const periodMap = mapPeriodsByType(comparisonMode.type);
        const periods1 = periodMap[comparisonMode.period1];
        const periods2 = periodMap[comparisonMode.period2];
        
        if (periods1 && periods2) {
            updateDetailTableWithComparison(periods1, periods2);
            // Recalcular resumo com dados filtrados
            showComparisonSummary(periods1, periods2);
        }
    } else {
        // Modo normal: atualizar dashboard normalmente
        updateDashboard();
    }
    
    // ✅ CORREÇÃO: Sempre atualizar os cards e gráficos, independente do modo
    updateSummaryCardsAndCharts();
    
    setTimeout(() => {
        populateFilterOptions();
    }, 100);
}

// ✅ NOVA FUNÇÃO: Separar atualização de cards e gráficos
function updateSummaryCardsAndCharts() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    if (currentPeriod && previousPeriod && selectedSheet && filteredData.length > 0) {
        // Sempre atualizar cards com dados dos períodos selecionados normalmente
        calculateSummaryValues(currentPeriod, previousPeriod);
        
        // Sempre atualizar gráficos com dados filtrados
        if (periods && periods.length > 0) {
            updateCharts();
        }
    }
}

function clearFilters() {
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

function calculateSummaryValues(currentPeriod, previousPeriod) {
    const entities = ['telecom', 'vogel', 'consolidado'];
    
    entities.forEach(entity => {
        try {
            const summary = safeCalculateSummary(filteredData, currentPeriod, previousPeriod, entity);
            updateSummaryCard(entity, summary.currentTotal, summary.previousTotal, summary.variation);
        } catch (error) {
            console.error(`Error calculating summary for ${entity}:`, error);
            updateSummaryCard(entity, 0, 0, 0);
        }
    });
}

function updateDashboard() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) {
        console.log('Dados insuficientes para atualizar dashboard');
        return;
    }
    
    if (!filteredData || filteredData.length === 0) {
        console.log('Nenhum dado filtrado disponível');
        return;
    }
    
    // Calculate summary values apenas se não estamos em modo de comparação
    if (!comparisonMode.active) {
        calculateSummaryValues(currentPeriod, previousPeriod);
        filterDetailTable();
    }
    
    if (periods && periods.length > 0) {
        updateCharts();
    }
}

function updateSummaryCard(entity, currentValue, previousValue, variation) {
    const variationElement = document.getElementById(`${entity}Variation`);
    const currentElement = document.getElementById(`${entity}Current`);
    const previousElement = document.getElementById(`${entity}Previous`);
    
    if (!variationElement || !currentElement || !previousElement) {
        console.error(`Elementos para ${entity} não encontrados`);
        return;
    }
    
    currentElement.textContent = formatCurrency(currentValue);
    previousElement.textContent = formatCurrency(previousValue);
    
    const variationValueElement = variationElement.querySelector('.variation-value');
    if (variationValueElement) {
        variationValueElement.textContent = `${Math.abs(variation).toFixed(2)}%`;
    }
    
    variationElement.className = 'variation-indicator';
    if (variation > 0) {
        variationElement.classList.add('positive');
    } else if (variation < 0) {
        variationElement.classList.add('negative');
    }
}

function updateCharts() {
    if (!periods || periods.length === 0 || !filteredData || filteredData.length === 0) {
        console.log('Dados insuficientes para atualizar gráficos');
        return;
    }
    
    updateEvolutionChart();
}

function updateEvolutionChart() {
    const chartElement = document.getElementById('evolutionChart');
    if (!chartElement) {
        console.error('Elemento do gráfico não encontrado');
        return;
    }
    
    if (!periods || periods.length === 0) {
        chartElement.innerHTML = '<div class="no-data">Selecione uma aba e períodos para visualizar o gráfico</div>';
        return;
    }
    
    if (!filteredData || filteredData.length === 0) {
        chartElement.innerHTML = '<div class="no-data">Nenhum dado disponível para o gráfico</div>';
        return;
    }

    const currentPeriod = document.getElementById('currentPeriod')?.value;
    const previousPeriod = document.getElementById('previousPeriod')?.value;
    
    if (!currentPeriod || !previousPeriod) {
        chartElement.innerHTML = '<div class="no-data">Selecione os períodos atual e anterior para visualizar o gráfico</div>';
        return;
    }
    
    let seriesData, categories;
    const isQuarterly = chartViewMode === 'quarterly';
    
    try {
        if (isQuarterly) {
            const quarterlyData = calculateQuarterlyData();
            if (!quarterlyData || Object.keys(quarterlyData).length === 0) {
                chartElement.innerHTML = '<div class="no-data">Dados insuficientes para gráfico trimestral</div>';
                return;
            }
            
            const quarters = sortQuarters(Object.keys(quarterlyData));
            
            seriesData = [
                {
                    name: 'Telecom',
                    data: quarters.map(quarter => ({
                        x: quarter,
                        y: quarterlyData[quarter]?.telecom || 0
                    }))
                },
                {
                    name: 'Vogel',
                    data: quarters.map(quarter => ({
                        x: quarter,
                        y: quarterlyData[quarter]?.vogel || 0
                    }))
                },
                {
                    name: 'Somado',
                    data: quarters.map(quarter => ({
                        x: quarter,
                        y: quarterlyData[quarter]?.consolidado || 0
                    }))
                }
            ];
            categories = quarters;
        } else {
            seriesData = [
                {
                    name: 'Telecom',
                    data: periods.map(period => ({
                        x: formatPeriod(period),
                        y: calculateTotal('telecom', period)
                    }))
                },
                {
                    name: 'Vogel',
                    data: periods.map(period => ({
                        x: formatPeriod(period),
                        y: calculateTotal('vogel', period)
                    }))
                },
                {
                    name: 'Somado',
                    data: periods.map(period => ({
                        x: formatPeriod(period),
                        y: calculateTotal('consolidado', period)
                    }))
                }
            ];
            categories = periods.map(formatPeriod);
        }
        
        if (!seriesData || seriesData.length === 0 || seriesData.every(serie => !serie.data || serie.data.length === 0)) {
            chartElement.innerHTML = '<div class="no-data">Dados insuficientes para gerar o gráfico</div>';
            return;
        }
        
        if (charts.evolution) {
            try {
                charts.evolution.destroy();
                charts.evolution = null;
            } catch (error) {
                console.log('Erro ao destruir gráfico anterior:', error);
            }
        }
        
        chartElement.innerHTML = '';
        
        setTimeout(() => {
            try {
                const options = {
                    series: seriesData,
                    chart: {
                        type: 'line',
                        height: 350,
                        fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
                        toolbar: {
                            show: false
                        },
                        animations: {
                            enabled: true,
                            easing: 'easeinout',
                            speed: 800
                        }
                    },
                    stroke: {
                        curve: 'smooth',
                        width: 3
                    },
                    dataLabels: {
                        enabled: false
                    },
                    markers: {
                        size: 6,
                        strokeWidth: 2,
                        fillOpacity: 1,
                        strokeOpacity: 1,
                        hover: {
                            size: 8
                        }
                    },
                    xaxis: {
                        type: 'category',
                        categories: categories,
                        labels: {
                            rotate: -45,
                            style: {
                                colors: '#4a5568',
                                fontSize: '12px',
                                fontWeight: 500
                            }
                        }
                    },
                    yaxis: {
                        title: {
                            text: 'Valor (R$)',
                            style: {
                                color: '#4a5568',
                                fontSize: '12px',
                                fontWeight: 500
                            }
                        },
                        labels: {
                            formatter: function(val) {
                                return formatCurrencyCompact(val);
                            }
                        }
                    },
                    colors: ['#0066cc', '#ffab00', '#00c853'],
                    fill: {
                        type: 'gradient',
                        gradient: {
                            shadeIntensity: 1,
                            opacityFrom: 0.7,
                            opacityTo: 0.9,
                            stops: [0, 100]
                        }
                    },
                    // ✅ TOOLTIP RESTAURADO COM VARIAÇÕES ENTRE PERÍODOS
                    tooltip: {
                        shared: true,
                        intersect: false,
                        custom: function({ series, seriesIndex, dataPointIndex, w }) {
                            const currentLabel = categories[dataPointIndex];
                            const entities = ['telecom', 'vogel', 'consolidado'];
                            const entityNames = ['Telecom', 'Vogel', 'Somado'];
                            const entityColors = ['#0066cc', '#ffab00', '#00c853'];
                            
                            let tooltipContent = `
                                <div style="background: white; padding: 12px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border: 1px solid #e5e7eb; min-width: 250px;">
                                    <div style="font-weight: 600; margin-bottom: 8px; color: #374151; font-size: 14px;">
                                        ${currentLabel}
                                    </div>
                            `;
                            
                            // Mostrar valores atuais
                            series.forEach((serieData, index) => {
                                const currentValue = serieData[dataPointIndex];
                                const entityName = entityNames[index];
                                const color = entityColors[index];
                                
                                // Calcular variação em relação ao período anterior se existir
                                let variationText = '';
                                if (dataPointIndex > 0) {
                                    const previousValue = serieData[dataPointIndex - 1];
                                    const variation = previousValue !== 0 ? ((currentValue - previousValue) / previousValue) * 100 : 0;
                                    const variationColor = variation >= 0 ? '#00c853' : '#ff3d57';
                                    const variationSymbol = variation >= 0 ? '↗' : '↘';
                                    
                                    variationText = `<span style="color: ${variationColor}; font-size: 11px; margin-left: 4px;">
                                        ${variationSymbol} ${Math.abs(variation).toFixed(1)}%
                                    </span>`;
                                }
                                
                                tooltipContent += `
                                    <div style="display: flex; align-items: center; margin: 4px 0;">
                                        <span style="width: 12px; height: 12px; background-color: ${color}; border-radius: 50%; margin-right: 8px;"></span>
                                        <span style="color: #374151; font-size: 13px; flex: 1;">
                                            <strong>${entityName}:</strong> ${formatCurrency(currentValue)}${variationText}
                                        </span>
                                    </div>
                                `;
                            });
                            
                            // Mostrar comparação com período anterior se existir
                            if (dataPointIndex > 0) {
                                const previousLabel = categories[dataPointIndex - 1];
                                tooltipContent += `
                                    <div style="border-top: 1px solid #e5e7eb; margin-top: 8px; padding-top: 6px;">
                                        <div style="font-size: 11px; color: #6b7280; margin-bottom: 4px;">
                                            vs ${previousLabel}
                                        </div>
                                `;
                                
                                // Calcular total da variação
                                const currentTotal = series.reduce((sum, serieData) => sum + serieData[dataPointIndex], 0);
                                const previousTotal = series.reduce((sum, serieData) => sum + serieData[dataPointIndex - 1], 0);
                                const totalVariation = previousTotal !== 0 ? ((currentTotal - previousTotal) / previousTotal) * 100 : 0;
                                const totalVariationColor = totalVariation >= 0 ? '#00c853' : '#ff3d57';
                                const totalVariationSymbol = totalVariation >= 0 ? '↗' : '↘';
                                
                                tooltipContent += `
                                        <div style="font-size: 12px; color: ${totalVariationColor}; font-weight: 600;">
                                            ${totalVariationSymbol} Total: ${formatCurrency(currentTotal - previousTotal)} (${Math.abs(totalVariation).toFixed(1)}%)
                                        </div>
                                    </div>
                                `;
                            }
                            
                            tooltipContent += '</div>';
                            return tooltipContent;
                        }
                    },
                    legend: {
                        position: 'top',
                        horizontalAlign: 'right',
                        fontSize: '12px',
                        fontWeight: 500,
                        offsetY: 0,
                        offsetX: -5
                    },
                    grid: {
                        borderColor: '#e5e7eb',
                        strokeDashArray: 4,
                        xaxis: {
                            lines: {
                                show: true
                            }
                        },
                        yaxis: {
                            lines: {
                                show: true
                            }
                        }
                    },
                    theme: {
                        mode: 'light'
                    }
                };
                
                charts.evolution = new ApexCharts(chartElement, options);
                charts.evolution.render();
                console.log('Gráfico renderizado com sucesso');
                
            } catch (error) {
                console.error('Erro ao criar gráfico:', error);
                chartElement.innerHTML = '<div class="no-data">Erro ao carregar gráfico</div>';
            }
        }, 100);
        
    } catch (error) {
        console.error('Erro geral no gráfico:', error);
        chartElement.innerHTML = '<div class="no-data">Erro ao processar dados do gráfico</div>';
    }
}

function sortQuarters(quarters) {
    return quarters.sort((a, b) => {
        const [quarterA, yearA] = a.split(' ');
        const [quarterB, yearB] = b.split(' ');
        
        const yearDiff = parseInt(yearA) - parseInt(yearB);
        if (yearDiff !== 0) return yearDiff;
        
        const quarterNumA = parseInt(quarterA.replace('T', ''));
        const quarterNumB = parseInt(quarterB.replace('T', ''));
        
        return quarterNumA - quarterNumB;
    });
}

function filterDetailTable() {
    const descriptionFilter = document.getElementById('descriptionFilter').value.toLowerCase();
    const sortFilter = document.getElementById('sortFilter').value;
    const entityFilter = document.getElementById('entityFilter').value;
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) return;
    
    let filtered = [...filteredData];
    if (descriptionFilter) {
        filtered = filtered.filter(item => {
            const { contaSAP, description } = extractAccountInfo(item);
            
            return description.toLowerCase().includes(descriptionFilter) || 
                   contaSAP.toString().toLowerCase().includes(descriptionFilter);
        });
    }
    
    switch (sortFilter) {
        case 'variation_desc':
            filtered.sort((a, b) => {
                return Math.abs(getItemVariation(b, currentPeriod, previousPeriod, entityFilter)) - 
                       Math.abs(getItemVariation(a, currentPeriod, previousPeriod, entityFilter));
            });
            break;
        case 'variation_asc':
            filtered.sort((a, b) => {
                return Math.abs(getItemVariation(a, currentPeriod, previousPeriod, entityFilter)) - 
                       Math.abs(getItemVariation(b, currentPeriod, previousPeriod, entityFilter));
            });
            break;
        case 'value_desc':
            filtered.sort((a, b) => {
                const valueA = getCurrentEntityValue(a, currentPeriod, entityFilter);
                const valueB = getCurrentEntityValue(b, currentPeriod, entityFilter);
                return valueB - valueA;
            });
            break;
        case 'value_asc':
            filtered.sort((a, b) => {
                const valueA = getCurrentEntityValue(a, currentPeriod, entityFilter);
                const valueB = getCurrentEntityValue(b, currentPeriod, entityFilter);
                return valueA - valueB;
            });
            break;
    }
    
    updateDetailTableContent(filtered, currentPeriod, previousPeriod, entityFilter);
}

function getCurrentEntityValue(item, currentPeriod, entityFilter) {
    if (!item.values || !item.values[currentPeriod]) return 0;
    
    if (entityFilter === 'total') {
        return Object.values(item.values[currentPeriod]).reduce((a, b) => a + b, 0);
    } else {
        return item.values[currentPeriod][entityFilter] || 0;
    }
}

function getItemVariation(item, currentPeriod, previousPeriod, entityFilter) {
    const currentValues = item.values && item.values[currentPeriod] ? item.values[currentPeriod] : {};
    const previousValues = item.values && item.values[previousPeriod] ? item.values[previousPeriod] : {};
    
    let currentTotal, previousTotal;
    
    if (entityFilter === 'total') {
        currentTotal = Object.values(currentValues).reduce((a, b) => a + b, 0);
        previousTotal = Object.values(previousValues).reduce((a, b) => a + b, 0);
    } else {
        currentTotal = currentValues[entityFilter] || 0;
        previousTotal = previousValues[entityFilter] || 0;
    }
    
    return currentTotal - previousTotal;
}

function updateDetailTableContent(data, currentPeriod, previousPeriod, entityFilter) {
    const tbody = document.querySelector('#detailsTable tbody');
    if (!tbody) {
        console.error('Table body element not found');
        return;
    }
    
    tbody.innerHTML = '';
    
    if (!data || data.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="6" class="no-data">Não há dados para exibir. Tente ajustar os filtros.</td>`;
        tbody.appendChild(row);
        return;
    }
    
    if (!currentPeriod || !previousPeriod) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="6" class="no-data">Selecione os períodos para visualizar os dados.</td>`;
        tbody.appendChild(row);
        return;
    }
    
    data.slice(0, 2000).forEach(item => {
        try {
            const row = document.createElement('tr');
            
            const currentValues = item.values && item.values[currentPeriod] ? item.values[currentPeriod] : {};
            const previousValues = item.values && item.values[previousPeriod] ? item.values[previousPeriod] : {};
            
            let currentTotal, previousTotal;
            
            if (entityFilter === 'total') {
                currentTotal = Object.values(currentValues).reduce((a, b) => a + b, 0);
                previousTotal = Object.values(previousValues).reduce((a, b) => a + b, 0);
            } else {
                currentTotal = currentValues[entityFilter] || 0;
                previousTotal = previousValues[entityFilter] || 0;
            }
            
            const variation = currentTotal - previousTotal;
            const variationPercent = previousTotal !== 0 ? (variation / previousTotal) * 100 : 0;
            
            const { contaSAP, description } = extractAccountInfo(item);
            
            row.innerHTML = `
                <td>${contaSAP || '-'}</td>
                <td>${description || '-'}</td>
                <td>${formatCurrency(previousTotal)}</td>
                <td>${formatCurrency(currentTotal)}</td>
                <td class="variation-cell ${variation >= 0 ? 'positive' : 'negative'}">
                    ${formatCurrency(variation)}
                </td>
                <td class="variation-cell ${variationPercent >= 0 ? 'positive' : 'negative'}">
                    ${variationPercent >= 0 ? '+' : ''}${variationPercent.toFixed(2)}%
                </td>
            `;
            
            tbody.appendChild(row);
        } catch (error) {
            console.error('Error rendering table row:', error, item);
        }
    });
}

function downloadChart() {
    if (!charts.evolution) return;
    
    charts.evolution.dataURI().then(({ imgURI }) => {
        const downloadLink = document.createElement('a');
        downloadLink.href = imgURI;
        const chartType = chartViewMode === 'quarterly' ? 'trimestral' : 'mensal';
        downloadLink.download = `evolucao-${chartType}-${new Date().toISOString().split('T')[0]}.png`;
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
    });
}

function exportTableToExcel() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    const entityFilter = document.getElementById('entityFilter').value;
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) return;
    
    // Determinar headers baseado no modo de comparação
    let headers, filename;
    if (comparisonMode.active) {
        headers = ['Conta SAP', 'Descrição', comparisonMode.period2, comparisonMode.period1, 'Variação', 'Variação %'];
        filename = `comparacao-${comparisonMode.type}-${comparisonMode.period1}-vs-${comparisonMode.period2}-${new Date().toISOString().split('T')[0]}.xlsx`;
    } else {
        headers = ['Conta SAP', 'Descrição', 'Valor Anterior', 'Valor Atual', 'Variação', 'Variação %'];
        const entityName = entityFilter === 'total' ? 'consolidado-total' : entityFilter;
        filename = `detalhamento-contas-${entityName}-${new Date().toISOString().split('T')[0]}.xlsx`;
    }
    
    const ws = XLSX.utils.aoa_to_sheet([headers]);
    
    const rows = filteredData.map(item => {
        const currentValues = item.values && item.values[currentPeriod] ? item.values[currentPeriod] : {};
        const previousValues = item.values && item.values[previousPeriod] ? item.values[previousPeriod] : {};
        
        let currentTotal, previousTotal;
        
        if (entityFilter === 'total') {
            currentTotal = Object.values(currentValues).reduce((a, b) => a + b, 0);
            previousTotal = Object.values(previousValues).reduce((a, b) => a + b, 0);
        } else {
            currentTotal = currentValues[entityFilter] || 0;
            previousTotal = previousValues[entityFilter] || 0;
        }
        
        const variation = currentTotal - previousTotal;
        const variationPercent = previousTotal !== 0 ? (variation / previousTotal) * 100 : 0;
        
        const { contaSAP, description } = extractAccountInfo(item);
        
        return [
            contaSAP || '-',
            description || '-',
            previousTotal,
            currentTotal,
            variation,
            variationPercent / 100
        ];
    });
    
    XLSX.utils.sheet_add_aoa(ws, rows, { origin: 'A2' });
    
    const currencyFormat = '"R$ "#,##0.00';
    const percentFormat = '0.00%';
    
    ['C', 'D', 'E'].forEach(col => {
        for (let i = 2; i <= rows.length + 1; i++) {
            if (!ws[col + i]) continue;
            ws[col + i].z = currencyFormat;
        }
    });
    
    for (let i = 2; i <= rows.length + 1; i++) {
        if (!ws['F' + i]) continue;
        ws['F' + i].z = percentFormat;
    }
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Detalhamento');
    
    XLSX.writeFile(wb, filename);
}

function calculateTotal(entity, period) {
    return filteredData.reduce((total, item) => {
        return total + (item.values && item.values[period] ? item.values[period][entity] || 0 : 0);
    }, 0);
}