// SCRIPT2.JS - Versão Corrigida para Problemas DOM e Duplicação
// ============================================
// CORREÇÃO DEFINITIVA PARA ERROS DE DOM E MÚLTIPLOS GRÁFICOS
// ============================================

// Variáveis globais para controle rigoroso
let isChartRendering = false;
let chartRenderTimeout = null;
let chartDestroyTimeout = null;
let lastFilterHash = '';

// Função para gerar hash dos filtros ativos
function generateFilterHash() {
    const filters = [];
    
    // Filtros dinâmicos
    const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
    dynamicFilters.forEach(select => {
        if (select && select.value) {
            filters.push(`${select.dataset.field}:${select.value}`);
        }
    });
    
    // Filtro de conta
    const accountInput = document.getElementById('filter_contaSearch');
    if (accountInput && accountInput.value) {
        filters.push(`conta:${accountInput.value}`);
    }
    
    // Períodos
    const currentPeriod = document.getElementById('currentPeriod')?.value;
    const previousPeriod = document.getElementById('previousPeriod')?.value;
    filters.push(`periods:${currentPeriod}-${previousPeriod}`);
    
    return filters.join('|');
}

function generateDynamicChartTitle() {
    const baseTitle = chartViewMode === 'quarterly' ? 'Evolução Trimestral' : 'Evolução Mensal';
    const filterParts = [];
    
    const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
    let hasActiveFilters = false;
    
    dynamicFilters.forEach(select => {
        if (select && select.value) {
            hasActiveFilters = true;
            const labelElement = select.previousElementSibling;
            if (labelElement && labelElement.textContent) {
                const label = labelElement.textContent.trim();
                filterParts.push(`${label}: ${select.value}`);
            }
        }
    });
    
    const accountSearchInput = document.getElementById('filter_contaSearch');
    if (accountSearchInput && accountSearchInput.value) {
        hasActiveFilters = true;
        const contaSAP = accountSearchInput.dataset.selectedConta;
        if (contaSAP) {
            filterParts.push(`Conta: ${contaSAP}`);
        } else {
            filterParts.push(`Busca: ${accountSearchInput.value}`);
        }
    }
    
    if (!hasActiveFilters && selectedSheet) {
        filterParts.push(`Aba: ${selectedSheet}`);
    }
    
    if (filterParts.length > 0) {
        return `${baseTitle} - ${filterParts.join(' | ')}`;
    }
    
    return baseTitle;
}

function updateDashboard() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    console.log('📊 === ATUALIZANDO DASHBOARD ===');
    console.log(`🔍 Períodos: ${currentPeriod} vs ${previousPeriod}`);
    console.log(`📋 Dados filtrados: ${filteredData.length} registros`);
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) {
        console.log('❌ Dados insuficientes para atualizar dashboard');
        return;
    }
    
    if (!filteredData || filteredData.length === 0) {
        console.log('❌ Nenhum dado filtrado disponível');
        clearDashboard();
        return;
    }
    
    if (!comparisonMode.active) {
        calculateSummaryValues(currentPeriod, previousPeriod);
        filterDetailTable();
    }
    
    if (periods && periods.length > 0) {
        updateChartsWithDebounce();
    }
    
    updateChartTitle();
    console.log('✅ Dashboard atualizado');
}

function clearDashboard() {
    updateSummaryCard('telecom', 0, 0, 0);
    updateSummaryCard('vogel', 0, 0, 0);
    updateSummaryCard('consolidado', 0, 0, 0);
    
    const chartElement = document.getElementById('evolutionChart');
    if (chartElement) {
        chartElement.innerHTML = '<div class="no-data">Selecione filtros válidos para visualizar os dados</div>';
    }
    
    const tbody = document.querySelector('#detailsTable tbody');
    if (tbody) {
        tbody.innerHTML = '<tr><td colspan="7" class="no-data">Nenhum dado disponível</td></tr>';
    }
}

function calculateSummaryValues(currentPeriod, previousPeriod) {
    const entities = ['telecom', 'vogel', 'consolidado'];
    
    console.log('💳 === CALCULANDO CARDS ===');
    console.log(`📊 Usando ${filteredData.length} registros filtrados`);
    
    entities.forEach(entity => {
        try {
            const summary = safeCalculateSummary(filteredData, currentPeriod, previousPeriod, entity);
            console.log(`💳 ${entity}: Atual=${formatCurrency(summary.currentTotal)}, Anterior=${formatCurrency(summary.previousTotal)}, Var=${summary.variation.toFixed(2)}%`);
            updateSummaryCard(entity, summary.currentTotal, summary.previousTotal, summary.variation);
        } catch (error) {
            console.error(`❌ Erro calculando card ${entity}:`, error);
            updateSummaryCard(entity, 0, 0, 0);
        }
    });
}

function updateSummaryCard(entity, currentValue, previousValue, variation) {
    const variationElement = document.getElementById(`${entity}Variation`);
    const currentElement = document.getElementById(`${entity}Current`);
    const previousElement = document.getElementById(`${entity}Previous`);
    
    if (!variationElement || !currentElement || !previousElement) {
        console.error(`❌ Elementos para ${entity} não encontrados`);
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

function updateSummaryCardsAndCharts() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    console.log('🔄 === ATUALIZANDO CARDS E GRÁFICOS ===');
    
    if (currentPeriod && previousPeriod && selectedSheet && filteredData.length > 0) {
        calculateSummaryValues(currentPeriod, previousPeriod);
        
        if (periods && periods.length > 0) {
            updateChartsWithDebounce();
        }
    }
}

// FUNÇÃO PRINCIPAL CORRIGIDA - Debounce agressivo
function updateChartsWithDebounce() {
    const currentFilterHash = generateFilterHash();
    
    // Se os filtros não mudaram, não re-renderizar
    if (currentFilterHash === lastFilterHash && charts.evolution) {
        console.log('⚡ Filtros inalterados, mantendo gráfico atual');
        updateChartTitle();
        return;
    }
    
    // Cancelar timeouts anteriores
    if (chartRenderTimeout) {
        clearTimeout(chartRenderTimeout);
        chartRenderTimeout = null;
    }
    
    if (chartDestroyTimeout) {
        clearTimeout(chartDestroyTimeout);
        chartDestroyTimeout = null;
    }
    
    // Prevenir renderização se já está em andamento
    if (isChartRendering) {
        console.log('⏳ Gráfico já renderizando, ignorando chamada');
        return;
    }
    
    console.log('📈 Agendando atualização do gráfico...');
    
    // Agendar renderização com debounce maior
    chartRenderTimeout = setTimeout(() => {
        lastFilterHash = currentFilterHash;
        safeUpdateEvolutionChart();
    }, 300); // Aumentado para 300ms
}

// Função de compatibilidade
function updateCharts() {
    updateChartsWithDebounce();
}

function updateChartTitle() {
    const chartTitle = document.getElementById('chartTitle');
    if (chartTitle) {
        chartTitle.textContent = generateDynamicChartTitle();
    }
}

// FUNÇÃO SEGURA DE ATUALIZAÇÃO DO GRÁFICO
function safeUpdateEvolutionChart() {
    if (isChartRendering) {
        console.log('⏳ Já renderizando, abortando');
        return;
    }
    
    isChartRendering = true;
    
    try {
        // Primeiro destruir o gráfico existente
        safeDestroyChart();
        
        // Aguardar um pouco antes de criar novo
        setTimeout(() => {
            try {
                updateEvolutionChart();
            } catch (error) {
                console.error('❌ Erro na renderização:', error);
                isChartRendering = false;
            }
        }, 100);
        
    } catch (error) {
        console.error('❌ Erro na atualização segura:', error);
        isChartRendering = false;
    }
}

// FUNÇÃO SEGURA PARA DESTRUIR GRÁFICO
function safeDestroyChart() {
    if (charts.evolution) {
        try {
            console.log('🗑️ Destruindo gráfico anterior...');
            
            // Verificar se o gráfico ainda existe no DOM
            const chartElement = document.getElementById('evolutionChart');
            if (chartElement && charts.evolution.w && charts.evolution.w.globals) {
                charts.evolution.destroy();
            }
            
            charts.evolution = null;
            
            // Limpar o elemento
            if (chartElement) {
                chartElement.innerHTML = '';
            }
            
            console.log('✅ Gráfico destruído com segurança');
            
        } catch (error) {
            console.warn('⚠️ Erro ao destruir gráfico:', error);
            charts.evolution = null;
            
            // Força limpeza do elemento
            const chartElement = document.getElementById('evolutionChart');
            if (chartElement) {
                chartElement.innerHTML = '';
            }
        }
    }
}

function updateEvolutionChart() {
    const chartElement = document.getElementById('evolutionChart');
    if (!chartElement) {
        console.error('❌ Elemento do gráfico não encontrado');
        isChartRendering = false;
        return;
    }
    
    console.log('📈 === RENDERIZANDO GRÁFICO ===');
    
    // Validações críticas
    if (!periods || periods.length === 0) {
        chartElement.innerHTML = '<div class="no-data">Selecione uma aba e períodos para visualizar o gráfico</div>';
        isChartRendering = false;
        return;
    }
    
    if (!filteredData || filteredData.length === 0) {
        chartElement.innerHTML = '<div class="no-data">Nenhum dado disponível com os filtros aplicados</div>';
        isChartRendering = false;
        return;
    }

    const currentPeriod = document.getElementById('currentPeriod')?.value;
    const previousPeriod = document.getElementById('previousPeriod')?.value;
    
    if (!currentPeriod || !previousPeriod) {
        chartElement.innerHTML = '<div class="no-data">Selecione os períodos atual e anterior</div>';
        isChartRendering = false;
        return;
    }
    
    const isQuarterly = chartViewMode === 'quarterly';
    
    console.log(`📊 Tipo: ${isQuarterly ? 'TRIMESTRAL' : 'MENSAL'}`);
    console.log(`🔍 Registros: ${filteredData.length}`);
    
    let seriesData, categories;
    
    try {
        if (isQuarterly) {
            const quarterlyData = calculateQuarterlyDataFiltered();
            if (!quarterlyData || Object.keys(quarterlyData).length === 0) {
                chartElement.innerHTML = '<div class="no-data">Dados insuficientes para gráfico trimestral</div>';
                isChartRendering = false;
                return;
            }
            
            const quarters = sortQuarters(Object.keys(quarterlyData));
            
            seriesData = [
                {
                    name: 'Telecom',
                    data: quarters.map(quarter => quarterlyData[quarter]?.telecom || 0)
                },
                {
                    name: 'Vogel',
                    data: quarters.map(quarter => quarterlyData[quarter]?.vogel || 0)
                },
                {
                    name: 'Somado',
                    data: quarters.map(quarter => quarterlyData[quarter]?.consolidado || 0)
                }
            ];
            categories = quarters;
        } else {
            seriesData = [
                {
                    name: 'Telecom',
                    data: periods.map(period => calculateTotalFiltered('telecom', period))
                },
                {
                    name: 'Vogel',
                    data: periods.map(period => calculateTotalFiltered('vogel', period))
                },
                {
                    name: 'Somado',
                    data: periods.map(period => calculateTotalFiltered('consolidado', period))
                }
            ];
            categories = periods.map(formatPeriod);
        }
        
        // CORREÇÃO: Validação dos dados - aceitar valores negativos também
        if (!seriesData || seriesData.length === 0) {
            chartElement.innerHTML = '<div class="no-data">Erro na preparação dos dados</div>';
            isChartRendering = false;
            return;
        }
        
        // CORRIGIDO: Verificar se há dados válidos (incluindo negativos)
        const hasValidData = seriesData.some(serie => 
            serie.data && serie.data.length > 0 && serie.data.some(value => value !== 0 && !isNaN(value) && isFinite(value))
        );
        
        // Log para debug
        console.log('📊 Dados das séries:', seriesData);
        console.log('📊 Tem dados válidos?', hasValidData);
        seriesData.forEach((serie, index) => {
            console.log(`📊 Série ${serie.name}:`, serie.data.slice(0, 5), '...');
        });
        
        if (!hasValidData) {
            chartElement.innerHTML = '<div class="no-data">Nenhum valor válido encontrado para os filtros aplicados</div>';
            isChartRendering = false;
            return;
        }
        
        // Garantir que o elemento ainda existe
        if (!document.getElementById('evolutionChart')) {
            console.warn('⚠️ Elemento do gráfico removido durante processamento');
            isChartRendering = false;
            return;
        }
        
        // RENDERIZAR GRÁFICO COM TIMEOUT SEGURO
        setTimeout(() => {
            renderChartSafely(chartElement, seriesData, categories);
        }, 50);
        
    } catch (error) {
        console.error('❌ Erro na preparação:', error);
        chartElement.innerHTML = '<div class="no-data">Erro ao processar dados</div>';
        isChartRendering = false;
    }
}

function renderChartSafely(chartElement, seriesData, categories) {
    try {
        // Verificação final do DOM
        if (!chartElement || !chartElement.parentNode || !document.getElementById('evolutionChart')) {
            console.error('❌ Elemento do gráfico inválido no momento da renderização');
            isChartRendering = false;
            return;
        }
        
        const currentPeriod = document.getElementById('currentPeriod')?.value;
        const previousPeriod = document.getElementById('previousPeriod')?.value;
        
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
                    enabled: false // DESATIVAR ANIMAÇÕES para evitar conflitos DOM
                },
                redrawOnParentResize: true,
                redrawOnWindowResize: true,
                background: 'transparent'
            },
            stroke: {
                curve: 'smooth',
                width: 3,
                lineCap: 'round'
            },
            dataLabels: {
                enabled: false
            },
            markers: {
                size: 5,
                strokeWidth: 2,
                fillOpacity: 1,
                strokeOpacity: 1,
                hover: {
                    size: 7,
                    sizeOffset: 2
                }
            },
            xaxis: {
                type: 'category',
                categories: categories,
                labels: {
                    rotate: -45,
                    style: {
                        colors: '#4a5568',
                        fontSize: '11px',
                        fontWeight: 500
                    },
                    trim: true
                },
                axisBorder: {
                    show: true,
                    color: '#e5e7eb'
                },
                axisTicks: {
                    show: true,
                    color: '#e5e7eb'
                }
            },
            yaxis: {
                title: {
                    text: 'Valor (R$)',
                    style: {
                        color: '#4a5568',
                        fontSize: '12px',
                        fontWeight: 600
                    }
                },
                labels: {
                    formatter: function(val) {
                        return formatCurrencyCompact(val);
                    },
                    style: {
                        colors: '#4a5568',
                        fontSize: '11px'
                    }
                }
            },
            colors: ['#2563eb', '#f59e0b', '#10b981'],
            fill: {
                type: 'solid',
                opacity: 1
            },
            tooltip: {
                shared: true,
                intersect: false,
                theme: 'light',
                style: {
                    fontSize: '12px'
                },
                custom: function({ series, seriesIndex, dataPointIndex, w }) {
                    const period = categories[dataPointIndex];
                    const periodFormatted = chartViewMode === 'quarterly' ? period : formatPeriod(periods[dataPointIndex]);
                    
                    let tooltipContent = `
                        <div class="custom-tooltip" style="padding: 10px; background: white; border: 1px solid #e5e7eb; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
                            <div style="font-weight: 600; margin-bottom: 8px; color: #1f2937;">${periodFormatted}</div>
                    `;
                    
                    // Adicionar valores das empresas
                    series.forEach((serie, index) => {
                        const value = serie[dataPointIndex];
                        const color = w.globals.colors[index];
                        const seriesName = w.globals.seriesNames[index];
                        
                        tooltipContent += `
                            <div style="display: flex; align-items: center; margin-bottom: 4px;">
                                <div style="width: 12px; height: 12px; background-color: ${color}; border-radius: 50%; margin-right: 8px;"></div>
                                <span style="color: #374151; font-size: 12px;">
                                    <strong>${seriesName}:</strong> ${formatCurrency(value)}
                                </span>
                            </div>
                        `;
                    });
                    
                    // Adicionar comparação com período anterior se disponível
                    if (dataPointIndex > 0 && chartViewMode === 'monthly') {
                        tooltipContent += '<div style="border-top: 1px solid #e5e7eb; margin-top: 8px; padding-top: 8px;">';
                        tooltipContent += '<div style="font-size: 11px; color: #6b7280; margin-bottom: 4px;">Variação vs período anterior:</div>';
                        
                        series.forEach((serie, index) => {
                            const currentValue = serie[dataPointIndex];
                            const previousValue = serie[dataPointIndex - 1];
                            const variation = currentValue - previousValue;
                            const variationPercent = previousValue !== 0 ? ((variation / previousValue) * 100) : 0;
                            const color = w.globals.colors[index];
                            const seriesName = w.globals.seriesNames[index];
                            const variationColor = variation >= 0 ? '#10b981' : '#ef4444';
                            const variationSymbol = variation >= 0 ? '+' : '';
                            
                            tooltipContent += `
                                <div style="display: flex; align-items: center; margin-bottom: 2px;">
                                    <div style="width: 8px; height: 8px; background-color: ${color}; border-radius: 50%; margin-right: 6px;"></div>
                                    <span style="color: ${variationColor}; font-size: 11px;">
                                        <strong>${seriesName}:</strong> ${variationSymbol}${formatCurrency(variation)} (${variationSymbol}${variationPercent.toFixed(1)}%)
                                    </span>
                                </div>
                            `;
                        });
                        
                        tooltipContent += '</div>';
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
                offsetY: -5,
                offsetX: -10,
                markers: {
                    width: 8,
                    height: 8,
                    radius: 4
                }
            },
            grid: {
                borderColor: '#e5e7eb',
                strokeDashArray: 3,
                xaxis: {
                    lines: {
                        show: true
                    }
                },
                yaxis: {
                    lines: {
                        show: true
                    }
                },
                padding: {
                    top: 0,
                    right: 20,
                    bottom: 0,
                    left: 10
                }
            },
            theme: {
                mode: 'light'
            },
            responsive: [{
                breakpoint: 768,
                options: {
                    chart: {
                        height: 300
                    },
                    legend: {
                        position: 'bottom',
                        offsetY: 10
                    },
                    xaxis: {
                        labels: {
                            rotate: -90,
                            style: {
                                fontSize: '10px'
                            }
                        }
                    }
                }
            }]
        };
        
        console.log('🎨 Criando gráfico...');
        charts.evolution = new ApexCharts(chartElement, options);
        
        charts.evolution.render().then(() => {
            console.log('✅ Gráfico renderizado com sucesso');
            isChartRendering = false;
        }).catch((error) => {
            console.error('❌ Erro no render:', error);
            chartElement.innerHTML = '<div class="no-data">Erro ao renderizar gráfico</div>';
            charts.evolution = null;
            isChartRendering = false;
        });
        
    } catch (error) {
        console.error('❌ Erro na criação do gráfico:', error);
        chartElement.innerHTML = '<div class="no-data">Erro na configuração do gráfico</div>';
        charts.evolution = null;
        isChartRendering = false;
    }
}

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

function calculateQuarterlyDataFiltered() {
    const quarterGroups = groupPeriodsByQuarter(periods);
    const quarterlyData = {};
    
    console.log('📊 === CALCULANDO DADOS TRIMESTRAIS ===');
    console.log(`🔍 Usando ${filteredData.length} registros filtrados`);
    
    Object.keys(quarterGroups).forEach(quarter => {
        const periodsInQuarter = quarterGroups[quarter];
        quarterlyData[quarter] = {
            telecom: 0,
            vogel: 0,
            consolidado: 0
        };
        
        periodsInQuarter.forEach(period => {
            quarterlyData[quarter].telecom += calculateTotalFiltered('telecom', period);
            quarterlyData[quarter].vogel += calculateTotalFiltered('vogel', period);
            quarterlyData[quarter].consolidado += calculateTotalFiltered('consolidado', period);
        });
    });
    
    return quarterlyData;
}

function calculateQuarterlyData() {
    return calculateQuarterlyDataFiltered();
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

function calculateTotalFiltered(entity, period) {
    const total = filteredData.reduce((sum, item) => {
        const value = (item.values && item.values[period] ? item.values[period][entity] || 0 : 0);
        return sum + value;
    }, 0);
    
    // Debug mais detalhado para valores negativos
    if (periods.indexOf(period) < 3) {
        console.log(`🔢 ${entity} - ${formatPeriod(period)}: ${formatCurrency(total)} (${filteredData.length} registros)`);
        
        // Log dos primeiros valores para debug
        filteredData.slice(0, 2).forEach((item, idx) => {
            const value = (item.values && item.values[period] ? item.values[period][entity] || 0 : 0);
            if (value !== 0) {
                console.log(`  Item[${idx}] ${entity}: ${formatCurrency(value)}`);
            }
        });
    }
    
    return total;
}

function calculateTotal(entity, period) {
    return calculateTotalFiltered(entity, period);
}

// ============================================
// FUNÇÕES DE FILTROS E TABELA
// ============================================

function filterDetailTable() {
    const descriptionFilter = document.getElementById('descriptionFilter').value.toLowerCase();
    const sortFilter = document.getElementById('sortFilter').value;
    const entityFilter = document.getElementById('entityFilter').value;
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    console.log('🔍 === FILTRANDO TABELA ===');
    console.log(`📊 Base: ${filteredData.length} | Filtro: "${descriptionFilter}" | Entidade: ${entityFilter}`);
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) {
        return;
    }
    
    let filtered = [...filteredData];
    
    if (descriptionFilter) {
        filtered = filtered.filter(item => {
            const { contaSAP, description } = extractAccountInfo(item);
            const contaMatch = contaSAP && contaSAP.toString().toLowerCase().includes(descriptionFilter);
            const descMatch = description && description.toLowerCase().includes(descriptionFilter);
            return contaMatch || descMatch;
        });
    }
    
    switch (sortFilter) {
        case 'variation_desc':
            filtered.sort((a, b) => Math.abs(getItemVariation(b, currentPeriod, previousPeriod, entityFilter)) - Math.abs(getItemVariation(a, currentPeriod, previousPeriod, entityFilter)));
            break;
        case 'variation_asc':
            filtered.sort((a, b) => Math.abs(getItemVariation(a, currentPeriod, previousPeriod, entityFilter)) - Math.abs(getItemVariation(b, currentPeriod, previousPeriod, entityFilter)));
            break;
        case 'value_desc':
            filtered.sort((a, b) => getCurrentEntityValue(b, currentPeriod, entityFilter) - getCurrentEntityValue(a, currentPeriod, entityFilter));
            break;
        case 'value_asc':
            filtered.sort((a, b) => getCurrentEntityValue(a, currentPeriod, entityFilter) - getCurrentEntityValue(b, currentPeriod, entityFilter));
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

// CORRIGIDA: Função updateDetailTableContent com suporte para modo expandido
function updateDetailTableContent(data, currentPeriod, previousPeriod, entityFilter) {
    const tbody = document.querySelector('#detailsTable tbody');
    if (!tbody) {
        console.error('❌ Elemento tbody não encontrado');
        return;
    }
    
    // DEBUG: Verificar condições
    const accountInput = document.getElementById('filter_contaSearch');
    const shouldExpand = shouldShowExpandedFormat();
    console.log('🔍 DEBUG TABELA:');
    console.log('  - Account Input Value:', accountInput?.value);
    console.log('  - Selected Conta:', accountInput?.dataset.selectedConta);
    console.log('  - Should Show Expanded:', shouldExpand);
    console.log('  - Entity Filter:', entityFilter);
    console.log('  - Data Length:', data.length);
    
    tbody.innerHTML = '';
    
    if (!data || data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="no-data">📋 Nenhum dado encontrado</td></tr>';
        return;
    }
    
    if (!currentPeriod || !previousPeriod) {
        tbody.innerHTML = '<tr><td colspan="7" class="no-data">⚠️ Selecione os períodos</td></tr>';
        return;
    }
    
    const fragment = document.createDocumentFragment();
    
    if (shouldExpand && data.length > 0) {
        console.log('🔄 ENTRANDO NO MODO EXPANDIDO');
        
        // NOVO: Mostrar as 3 empresas em linhas separadas para conta específica
        const item = data[0]; // Pega a primeira (e provavelmente única) conta
        const { contaSAP, description } = extractAccountInfo(item);
        
        const entities = [
            { key: 'telecom', name: 'Telecom' },
            { key: 'vogel', name: 'Vogel' },
            { key: 'consolidado', name: 'Somado' }
        ];
        
        // SEMPRE mostrar as 3 empresas quando conta específica (ignorar filtro dropdown)
        const entitiesToShow = entities;
        
        console.log('  - Entities to Show:', entitiesToShow);
        console.log('  - Entity Filter:', entityFilter);
        console.log('🔄 MODO EXPANDIDO - Mostrando todas as empresas para conta:', contaSAP);
        
        entitiesToShow.forEach((entity, index) => {
            const currentValues = item.values && item.values[currentPeriod] ? item.values[currentPeriod] : {};
            const previousValues = item.values && item.values[previousPeriod] ? item.values[previousPeriod] : {};
            
            const currentTotal = currentValues[entity.key] || 0;
            const previousTotal = previousValues[entity.key] || 0;
            const variation = currentTotal - previousTotal;
            const variationPercent = previousTotal !== 0 ? (variation / previousTotal) * 100 : 0;
            
            const variationClass = variation >= 0 ? 'positive' : 'negative';
            const variationSymbol = variation >= 0 ? '+' : '';
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td><strong>${entity.name}</strong></td>
                <td><strong>${contaSAP || '-'}</strong></td>
                <td>${description || '-'}</td>
                <td>${formatCurrency(previousTotal)}</td>
                <td><strong>${formatCurrency(currentTotal)}</strong></td>
                <td class="variation-cell ${variationClass}">
                    <strong>${formatCurrency(variation)}</strong>
                </td>
                <td class="variation-cell ${variationClass}">
                    <strong>${variationSymbol}${variationPercent.toFixed(2)}%</strong>
                </td>
            `;
            
            if (index % 2 === 0) {
                row.classList.add('even-row');
            }
            
            fragment.appendChild(row);
        });
        
    } else {
        // EXISTENTE: Comportamento normal para múltiplas contas
        const maxRows = 2000;
        const rowsToRender = Math.min(data.length, maxRows);
        
        console.log(`📋 Renderizando ${rowsToRender} de ${data.length} registros`);
        
        data.slice(0, maxRows).forEach((item, index) => {
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
                
                const variationClass = variation >= 0 ? 'positive' : 'negative';
                const variationSymbol = variation >= 0 ? '+' : '';
                
                const entityName = entityFilter === 'total' ? 'Consolidado' : 
                                 entityFilter === 'consolidado' ? 'Somado' :
                                 entityFilter.charAt(0).toUpperCase() + entityFilter.slice(1);
                
                row.innerHTML = `
                    <td>${entityName}</td>
                    <td><strong>${contaSAP || '-'}</strong></td>
                    <td>${description || '-'}</td>
                    <td>${formatCurrency(previousTotal)}</td>
                    <td><strong>${formatCurrency(currentTotal)}</strong></td>
                    <td class="variation-cell ${variationClass}">
                        <strong>${formatCurrency(variation)}</strong>
                    </td>
                    <td class="variation-cell ${variationClass}">
                        <strong>${variationSymbol}${variationPercent.toFixed(2)}%</strong>
                    </td>
                `;
                
                if (index % 2 === 0) {
                    row.classList.add('even-row');
                }
                
                fragment.appendChild(row);
            } catch (error) {
                console.error('❌ Erro ao renderizar linha:', error);
            }
        });
        
        if (data.length > maxRows) {
            const infoRow = document.createElement('tr');
            infoRow.innerHTML = `<td colspan="7" class="no-data" style="background-color: #f8f9fa; font-style: italic;">📊 Mostrando ${maxRows} de ${data.length} registros</td>`;
            fragment.appendChild(infoRow);
        }
    }
    
    tbody.appendChild(fragment);
    console.log('✅ Tabela atualizada');
}

// ============================================
// FUNÇÕES DE DOWNLOAD E EXPORT
// ============================================

function downloadChart() {
    if (!charts.evolution) {
        alert('❌ Nenhum gráfico disponível para download');
        return;
    }
    
    console.log('📥 Iniciando download do gráfico...');
    
    charts.evolution.dataURI({ 
        width: 1200, 
        height: 600,
        format: 'PNG'
    }).then(({ imgURI }) => {
        const downloadLink = document.createElement('a');
        downloadLink.href = imgURI;
        
        const chartType = chartViewMode === 'quarterly' ? 'trimestral' : 'mensal';
        const filterInfo = generateFileFilterInfo();
        const timestamp = new Date().toISOString().split('T')[0];
        
        downloadLink.download = `algar-evolucao-${chartType}${filterInfo}-${timestamp}.png`;
        
        document.body.appendChild(downloadLink);
        downloadLink.click();
        document.body.removeChild(downloadLink);
        
        console.log('✅ Gráfico baixado:', downloadLink.download);
    }).catch(error => {
        console.error('❌ Erro ao baixar gráfico:', error);
        alert('❌ Erro ao baixar gráfico. Tente novamente.');
    });
}

function exportChartToExcel() {
    if (!periods || periods.length === 0 || !filteredData || filteredData.length === 0) {
        alert('❌ Nenhum dado disponível para exportar');
        return;
    }
    
    console.log('📊 === EXPORTANDO GRÁFICO PARA EXCEL ===');
    console.log(`📋 Processando ${filteredData.length} registros filtrados`);
    
    try {
        const isQuarterly = chartViewMode === 'quarterly';
        let data, headers;
        
        if (isQuarterly) {
            const quarterlyData = calculateQuarterlyData();
            const quarters = sortQuarters(Object.keys(quarterlyData));
            
            headers = ['Empresa / Período', ...quarters];
            
            const telecomRow = ['Telecom', ...quarters.map(q => quarterlyData[q].telecom)];
            const vogelRow = ['Vogel', ...quarters.map(q => quarterlyData[q].vogel)];
            const consolidadoRow = ['Somado', ...quarters.map(q => quarterlyData[q].consolidado)];
            
            const telecomVarRow = ['Telecom (Variação %)', '', ...quarters.slice(1).map((q, index) => {
                const current = quarterlyData[q].telecom;
                const previous = quarterlyData[quarters[index]].telecom;
                return previous !== 0 ? ((current - previous) / previous) : 0;
            })];
            
            const vogelVarRow = ['Vogel (Variação %)', '', ...quarters.slice(1).map((q, index) => {
                const current = quarterlyData[q].vogel;
                const previous = quarterlyData[quarters[index]].vogel;
                return previous !== 0 ? ((current - previous) / previous) : 0;
            })];
            
            const consolidadoVarRow = ['Somado (Variação %)', '', ...quarters.slice(1).map((q, index) => {
                const current = quarterlyData[q].consolidado;
                const previous = quarterlyData[quarters[index]].consolidado;
                return previous !== 0 ? ((current - previous) / previous) : 0;
            })];
            
            data = [
                headers,
                [''],
                ['VALORES ABSOLUTOS'],
                telecomRow,
                vogelRow,
                consolidadoRow,
                [''],
                ['VARIAÇÕES PERCENTUAIS'],
                telecomVarRow,
                vogelVarRow,
                consolidadoVarRow
            ];
        } else {
            const formattedPeriods = periods.map(p => formatPeriod(p));
            headers = ['Empresa / Período', ...formattedPeriods];
            
            const telecomRow = ['Telecom', ...periods.map(p => calculateTotalFiltered('telecom', p))];
            const vogelRow = ['Vogel', ...periods.map(p => calculateTotalFiltered('vogel', p))];
            const consolidadoRow = ['Somado', ...periods.map(p => calculateTotalFiltered('consolidado', p))];
            
            const telecomVarRow = ['Telecom (Variação %)', '', ...periods.slice(1).map((p, index) => {
                const current = calculateTotalFiltered('telecom', p);
                const previous = calculateTotalFiltered('telecom', periods[index]);
                return previous !== 0 ? ((current - previous) / previous) : 0;
            })];
            
            const vogelVarRow = ['Vogel (Variação %)', '', ...periods.slice(1).map((p, index) => {
                const current = calculateTotalFiltered('vogel', p);
                const previous = calculateTotalFiltered('vogel', periods[index]);
                return previous !== 0 ? ((current - previous) / previous) : 0;
            })];
            
            const consolidadoVarRow = ['Somado (Variação %)', '', ...periods.slice(1).map((p, index) => {
                const current = calculateTotalFiltered('consolidado', p);
                const previous = calculateTotalFiltered('consolidado', periods[index]);
                return previous !== 0 ? ((current - previous) / previous) : 0;
            })];
            
            data = [
                headers,
                [''],
                ['VALORES ABSOLUTOS'],
                telecomRow,
                vogelRow,
                consolidadoRow,
                [''],
                ['VARIAÇÕES PERCENTUAIS'],
                telecomVarRow,
                vogelVarRow,
                consolidadoVarRow
            ];
        }
        
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        const currencyFormat = '"R$ "#,##0.00_);("R$ "#,##0.00)';
        const percentFormat = '0.00%';
        
        for (let col = 1; col < headers.length; col++) {
            const colLetter = String.fromCharCode(65 + col);
            
            [4, 5, 6].forEach(row => {
                const cellAddr = colLetter + row;
                if (ws[cellAddr]) {
                    ws[cellAddr].z = currencyFormat;
                }
            });
            
            [9, 10, 11].forEach(row => {
                const cellAddr = colLetter + row;
                if (ws[cellAddr]) {
                    ws[cellAddr].z = percentFormat;
                }
            });
        }
        
        const colWidths = [{ wch: 20 }, ...headers.slice(1).map(() => ({ wch: 15 }))];
        ws['!cols'] = colWidths;
        
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Evolução Financeira');
        
        wb.Props = {
            Title: 'Evolução Financeira - Algar',
            Subject: `Análise ${isQuarterly ? 'Trimestral' : 'Mensal'}`,
            Author: 'Dashboard Analytics Algar',
            CreatedDate: new Date()
        };
        
        const chartType = isQuarterly ? 'trimestral' : 'mensal';
        const filterInfo = generateFileFilterInfo();
        const timestamp = new Date().toISOString().split('T')[0];
        const filename = `algar-evolucao-${chartType}${filterInfo}-${timestamp}.xlsx`;
        
        XLSX.writeFile(wb, filename);
        
        console.log('✅ Excel exportado:', filename);
        
    } catch (error) {
        console.error('❌ Erro ao exportar Excel:', error);
        alert('❌ Erro ao exportar Excel. Verifique os dados e tente novamente.');
    }
}

function generateFileFilterInfo() {
    let filterInfo = '';
    
    try {
        const activeFilters = [];
        
        const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
        dynamicFilters.forEach(select => {
            if (select && select.value) {
                const labelElement = select.previousElementSibling;
                if (labelElement && labelElement.textContent) {
                    const label = labelElement.textContent.replace(/[^a-zA-Z0-9]/g, '');
                    const value = select.value.replace(/[^a-zA-Z0-9]/g, '');
                    activeFilters.push(`${label}-${value}`);
                }
            }
        });
        
        const accountSearchInput = document.getElementById('filter_contaSearch');
        if (accountSearchInput && accountSearchInput.value) {
            const contaSAP = accountSearchInput.dataset.selectedConta;
            if (contaSAP) {
                const cleanConta = contaSAP.replace(/[^a-zA-Z0-9]/g, '');
                activeFilters.push(`Conta-${cleanConta}`);
            } else {
                const cleanValue = accountSearchInput.value.replace(/[^a-zA-Z0-9]/g, '').substring(0, 10);
                activeFilters.push(`Busca-${cleanValue}`);
            }
        }
        
        if (activeFilters.length === 0 && selectedSheet) {
            filterInfo = `-${selectedSheet.replace(/[^a-zA-Z0-9]/g, '')}`;
        } else if (activeFilters.length > 0) {
            filterInfo = `-${activeFilters.join('-')}`;
        }
        
        if (filterInfo.length > 50) {
            filterInfo = filterInfo.substring(0, 50);
        }
        
    } catch (error) {
        console.error('❌ Erro ao gerar info do filtro:', error);
        if (selectedSheet) {
            filterInfo = `-${selectedSheet.replace(/[^a-zA-Z0-9]/g, '')}`;
        }
    }
    
    return filterInfo;
}

// CORRIGIDA: Função exportTableToExcel com suporte para modo expandido
// FUNÇÃO CORRIGIDA: Exportar exatamente como mostrado na tabela
function exportTableToExcel() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    const entityFilter = document.getElementById('entityFilter').value;
    const descriptionFilter = document.getElementById('descriptionFilter').value.toLowerCase();
    const sortFilter = document.getElementById('sortFilter').value;
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) {
        alert('❌ Selecione os períodos antes de exportar');
        return;
    }
    
    console.log('📋 === EXPORTANDO TABELA EXATAMENTE COMO MOSTRADA ===');
    console.log(`📊 Processando ${filteredData.length} registros filtrados`);
    console.log(`📋 Entidade: ${entityFilter}`);
    console.log(`🔍 Filtro descrição: "${descriptionFilter}"`);
    console.log(`📊 Ordenação: ${sortFilter}`);
    
    try {
        let headers, filename, rows;
        
        // ============================================
        // MODO COMPARAÇÃO - Replicar exatamente updateDetailTableWithComparison
        // ============================================
        if (comparisonMode.active) {
            // Headers da comparação com nomes dos períodos reais
            headers = ['Empresa', 'Conta SAP', 'Descrição', comparisonMode.period2, comparisonMode.period1, 'Variação (R$)', 'Variação (%)'];
            filename = `algar-comparacao-${comparisonMode.type}-${comparisonMode.period1}-vs-${comparisonMode.period2}-${new Date().toISOString().split('T')[0]}.xlsx`;
            
            console.log('📋 Exportando modo COMPARAÇÃO');
            
            // Se está no modo expandido (conta específica) em comparação
            if (shouldShowExpandedFormat() && filteredData.length > 0) {
                console.log('  🔄 Modo expandido - 3 empresas para conta específica');
                
                const item = filteredData[0];
                const { contaSAP, description } = extractAccountInfo(item);
                
                const entities = [
                    { key: 'telecom', name: 'Telecom' },
                    { key: 'vogel', name: 'Vogel' },
                    { key: 'consolidado', name: 'Somado' }
                ];
                
                const periodMap = mapPeriodsByType(comparisonMode.type);
                const periods1 = periodMap[comparisonMode.period1];
                const periods2 = periodMap[comparisonMode.period2];
                
                rows = entities.map(entity => {
                    let value1 = 0, value2 = 0;
                    
                    periods1.forEach(p => {
                        value1 += safeGetItemValue(item, p, entity.key);
                    });
                    periods2.forEach(p => {
                        value2 += safeGetItemValue(item, p, entity.key);
                    });
                    
                    const variation = value1 - value2;
                    const variationPercent = value2 !== 0 ? (variation / value2) * 100 : 0;
                    
                    return [
                        entity.name,
                        contaSAP || '-',
                        description || '-',
                        value2,
                        value1,
                        variation,
                        variationPercent / 100
                    ];
                });
                
            } else {
                // Modo normal de comparação - múltiplas contas
                console.log('  📊 Modo normal de comparação');
                
                const periodMap = mapPeriodsByType(comparisonMode.type);
                const periods1 = periodMap[comparisonMode.period1];
                const periods2 = periodMap[comparisonMode.period2];
                
                // Calcular dados para cada conta (igual ao updateDetailTableWithComparison)
                let comparisonData = filteredData.map(item => {
                    const { contaSAP, description } = extractAccountInfo(item);
                    
                    let value1 = 0, value2 = 0;
                    
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
                
                // Aplicar filtro de variação mínima se especificado
                if (comparisonMode.minVariation) {
                    comparisonData = comparisonData.filter(item => 
                        item.absVariationPercent >= comparisonMode.minVariation
                    );
                }
                
                // Ordenar por maior variação absoluta (igual ao código original)
                comparisonData.sort((a, b) => b.absVariationPercent - a.absVariationPercent);
                
                // Limitar resultados
                if (comparisonMode.maxResults !== 'all') {
                    comparisonData = comparisonData.slice(0, parseInt(comparisonMode.maxResults));
                }
                
                // Aplicar filtro de descrição/conta se houver
                if (descriptionFilter) {
                    comparisonData = comparisonData.filter(item => {
                        const descMatch = item.description && item.description.toLowerCase().includes(descriptionFilter);
                        const contaMatch = item.contaSAP && item.contaSAP.toString().toLowerCase().includes(descriptionFilter);
                        return descMatch || contaMatch;
                    });
                }
                
                rows = comparisonData.map(item => {
                    const entityName = comparisonMode.entityFilter === 'total' ? 'Consolidado' : 
                                     comparisonMode.entityFilter === 'consolidado' ? 'Somado' :
                                     comparisonMode.entityFilter.charAt(0).toUpperCase() + comparisonMode.entityFilter.slice(1);
                    
                    return [
                        entityName,
                        item.contaSAP || '-',
                        item.description || '-',
                        item.value2,
                        item.value1,
                        item.variation,
                        item.variationPercent / 100
                    ];
                });
            }
            
        } else {
            // ============================================
            // MODO NORMAL - Replicar exatamente updateDetailTableContent
            // ============================================
            console.log('📋 Exportando modo NORMAL');
            
            headers = ['Empresa', 'Conta SAP', 'Descrição', 'Valor Mês Anterior (R$)', 'Valor Mês Atual (R$)', 'Variação (R$)', 'Variação (%)'];
            const entityName = entityFilter === 'total' ? 'consolidado-total' : entityFilter;
            const filterInfo = generateFileFilterInfo();
            const timestamp = new Date().toISOString().split('T')[0];
            filename = `algar-detalhamento-${entityName}${filterInfo}-${timestamp}.xlsx`;
            
            // Aplicar EXATAMENTE os mesmos filtros que a tabela usa
            let tableData = [...filteredData];
            
            // Filtro de descrição/conta (igual ao filterDetailTable)
            if (descriptionFilter) {
                tableData = tableData.filter(item => {
                    const { contaSAP, description } = extractAccountInfo(item);
                    const contaMatch = contaSAP && contaSAP.toString().toLowerCase().includes(descriptionFilter);
                    const descMatch = description && description.toLowerCase().includes(descriptionFilter);
                    return contaMatch || descMatch;
                });
            }
            
            // Ordenação (igual ao filterDetailTable)
            switch (sortFilter) {
                case 'variation_desc':
                    tableData.sort((a, b) => Math.abs(getItemVariation(b, currentPeriod, previousPeriod, entityFilter)) - Math.abs(getItemVariation(a, currentPeriod, previousPeriod, entityFilter)));
                    break;
                case 'variation_asc':
                    tableData.sort((a, b) => Math.abs(getItemVariation(a, currentPeriod, previousPeriod, entityFilter)) - Math.abs(getItemVariation(b, currentPeriod, previousPeriod, entityFilter)));
                    break;
                case 'value_desc':
                    tableData.sort((a, b) => getCurrentEntityValue(b, currentPeriod, entityFilter) - getCurrentEntityValue(a, currentPeriod, entityFilter));
                    break;
                case 'value_asc':
                    tableData.sort((a, b) => getCurrentEntityValue(a, currentPeriod, entityFilter) - getCurrentEntityValue(b, currentPeriod, entityFilter));
                    break;
            }
            
            // Verificar se deve usar modo expandido
            if (shouldShowExpandedFormat() && tableData.length > 0) {
                console.log('  🔄 Modo expandido - 3 empresas para conta específica');
                
                const item = tableData[0];
                const { contaSAP, description } = extractAccountInfo(item);
                
                const entities = [
                    { key: 'telecom', name: 'Telecom' },
                    { key: 'vogel', name: 'Vogel' },
                    { key: 'consolidado', name: 'Somado' }
                ];
                
                // SEMPRE mostrar as 3 empresas quando conta específica (igual ao código)
                rows = entities.map(entity => {
                    const currentValues = item.values && item.values[currentPeriod] ? item.values[currentPeriod] : {};
                    const previousValues = item.values && item.values[previousPeriod] ? item.values[previousPeriod] : {};
                    
                    const currentTotal = currentValues[entity.key] || 0;
                    const previousTotal = previousValues[entity.key] || 0;
                    const variation = currentTotal - previousTotal;
                    const variationPercent = previousTotal !== 0 ? (variation / previousTotal) * 100 : 0;
                    
                    return [
                        entity.name,
                        contaSAP || '-',
                        description || '-',
                        previousTotal,
                        currentTotal,
                        variation,
                        variationPercent / 100
                    ];
                });
                
            } else {
                console.log('  📊 Modo normal - uma linha por conta');
                
                // Limitar a 2000 registros como na tabela
                const maxRows = 2000;
                const dataToExport = tableData.slice(0, maxRows);
                
                rows = dataToExport.map(item => {
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
                    
                    const entityDisplayName = entityFilter === 'total' ? 'Consolidado' : 
                                             entityFilter === 'consolidado' ? 'Somado' :
                                             entityFilter.charAt(0).toUpperCase() + entityFilter.slice(1);
                    
                    return [
                        entityDisplayName,
                        contaSAP || '-',
                        description || '-',
                        previousTotal,
                        currentTotal,
                        variation,
                        variationPercent / 100
                    ];
                });
                
                // Adicionar nota se limitamos os registros
                if (tableData.length > maxRows) {
                    // Adicionar linha informativa
                    rows.push([
                        '',
                        '',
                        `Mostrando ${maxRows} de ${tableData.length} registros`,
                        '',
                        '',
                        '',
                        ''
                    ]);
                }
            }
        }
        
        // Se não há dados, adicionar linha informativa
        if (!rows || rows.length === 0) {
            rows = [['', '', 'Nenhum dado encontrado com os filtros aplicados', '', '', '', '']];
        }
        
        // Criar planilha
        const ws = XLSX.utils.aoa_to_sheet([headers]);
        XLSX.utils.sheet_add_aoa(ws, rows, { origin: 'A2' });
        
        // Formatação
        const currencyFormat = '"R$ "#,##0.00_);("R$ "#,##0.00)';
        const percentFormat = '0.00%';
        
        // Aplicar formato de moeda nas colunas de valores e variação
        ['D', 'E', 'F'].forEach(col => {
            for (let i = 2; i <= rows.length + 1; i++) {
                if (ws[col + i]) {
                    ws[col + i].z = currencyFormat;
                }
            }
        });
        
        // Aplicar formato de porcentagem na coluna de variação %
        for (let i = 2; i <= rows.length + 1; i++) {
            if (ws['G' + i]) {
                ws['G' + i].z = percentFormat;
            }
        }
        
        // Ajustar largura das colunas
        ws['!cols'] = [
            { wch: 12 },  // Empresa
            { wch: 15 },  // Conta SAP
            { wch: 45 },  // Descrição
            { wch: 18 },  // Valor Anterior
            { wch: 18 },  // Valor Atual
            { wch: 18 },  // Variação
            { wch: 12 }   // Var %
        ];
        
        // Criar workbook e adicionar metadados
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Detalhamento por Conta');
        
        wb.Props = {
            Title: 'Detalhamento por Conta - Algar',
            Subject: comparisonMode.active ? 'Análise Comparativa' : 'Análise Financeira Detalhada',
            Author: 'Dashboard Analytics Algar',
            CreatedDate: new Date()
        };
        
        // Gerar arquivo
        XLSX.writeFile(wb, filename);
        
        console.log('✅ Excel da tabela exportado (espelho exato):', filename);
        console.log(`   📊 Registros exportados: ${rows.length}`);
        
    } catch (error) {
        console.error('❌ Erro ao exportar tabela:', error);
        alert('❌ Erro ao exportar tabela. Verifique os dados e tente novamente.');
    }
}

// Função para ser chamada quando mudar de aba ou limpar filtros
function forceChartCleanup() {
    console.log('🧹 Limpeza forçada de gráficos...');
    
    // Cancelar todos os timeouts
    if (chartRenderTimeout) {
        clearTimeout(chartRenderTimeout);
        chartRenderTimeout = null;
    }
    
    if (chartDestroyTimeout) {
        clearTimeout(chartDestroyTimeout);
        chartDestroyTimeout = null;
    }
    
    // Resetar flags
    isChartRendering = false;
    lastFilterHash = '';
    
    // Destruir gráfico de forma segura
    safeDestroyChart();
    
    console.log('✅ Limpeza de gráficos concluída');
}

// Função para ser chamada no window.resize
function handleChartResize() {
    if (charts.evolution && !isChartRendering) {
        try {
            charts.evolution.updateOptions({
                chart: {
                    width: '100%'
                }
            });
        } catch (error) {
            console.warn('⚠️ Erro no resize do gráfico:', error);
        }
    }
}

