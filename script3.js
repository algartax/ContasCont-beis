// ============================================
// SCRIPT3.JS - SISTEMA DE EXPORT CONSOLIDADO
// ============================================
// Sistema profissional para exportar todas as contas da aba selecionada
// com agrupamento por período (mensal, bimestral, trimestral, semestral, anual)
// e cálculo automático de variações percentuais

// ============================================
// VARIÁVEIS GLOBAIS DO SISTEMA DE EXPORT
// ============================================

let exportConfig = {
    active: false,
    periodType: 'mensal', // mensal, bimestral, trimestral, semestral, anual
    startPeriod: null,
    endPeriod: null,
    includeVariations: true,
    exportFormat: 'xlsx', // xlsx, csv
    includeCharts: false,
    comparisonMode: 'sequential' // 'sequential' (consecutivo) ou 'comparative' (apenas 2 períodos)
};

// ============================================
// INICIALIZAÇÃO DO SISTEMA DE EXPORT
// ============================================

function initializeExportSystem() {
    console.log('🚀 Inicializando Sistema de Export Consolidado...');
    
    // Aguardar DOM estar pronto
    setTimeout(() => {
        addExportButtonToSidebar();
        initializeExportModal();
    }, 500);
    
    console.log('✅ Sistema de Export Consolidado inicializado');
}

// ============================================
// INTERFACE DO SISTEMA DE EXPORT
// ============================================

function addExportButtonToSidebar() {
    const uploadTrigger = document.querySelector('.upload-trigger');
    
    if (!uploadTrigger) {
        console.error('❌ Elemento .upload-trigger não encontrado');
        return;
    }
    
    // Verificar se botão já existe
    if (document.getElementById('consolidatedExportButton')) {
        console.log('⚠️ Botão já existe');
        return;
    }
    
    // Criar botão de export consolidado
    const exportButton = document.createElement('button');
    exportButton.className = 'btn btn-outline';
    exportButton.id = 'consolidatedExportButton';
    exportButton.innerHTML = `
        <i class="fas fa-file-export"></i>
        Export Total
    `;
    exportButton.title = 'Exportar todas as contas da aba selecionada';
    
    // Adicionar evento
    exportButton.addEventListener('click', () => {
        if (!selectedSheet || !rawData || rawData.length === 0) {
            alert('❌ Selecione uma aba com dados antes de exportar');
            return;
        }
        openConsolidatedExportModal();
    });
    
    // Inserir antes do botão de trocar arquivo
    uploadTrigger.appendChild(exportButton);

    
    console.log('✅ Botão de Export Total adicionado');
}

function openConsolidatedExportModal() {
    const modal = document.getElementById('consolidatedExportModal');
    if (!modal) {
        console.error('❌ Modal de export não encontrado');
        return;
    }
    
    // Resetar configurações
    resetExportConfig();
    
    // Popular períodos disponíveis
    populateAvailablePeriods();
    
    // Mostrar modal
    modal.classList.add('active');
    
    console.log('📋 Modal de Export Consolidado aberto');
}

function resetExportConfig() {
    exportConfig = {
        active: false,
        periodType: 'mensal',
        startPeriod: null,
        endPeriod: null,
        includeVariations: true, // Sempre true
        exportFormat: 'xlsx', // Sempre xlsx
        includeCharts: false, // Sempre false
        comparisonMode: 'sequential'
    };
    
    // Resetar interface
    document.querySelectorAll('.export-period-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelector('[data-period-type="mensal"]')?.classList.add('active');
    
    // Resetar modo de comparação
    document.querySelectorAll('.comparison-mode-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelector('[data-comparison-mode="sequential"]')?.classList.add('active');
    
    const startPeriod = document.getElementById('exportStartPeriod');
    const endPeriod = document.getElementById('exportEndPeriod');
    
    if (startPeriod) startPeriod.value = '';
    if (endPeriod) endPeriod.value = '';
    
    // Mostrar/esconder controles baseado no modo
    updateComparisonModeInterface();
}

function populateAvailablePeriods() {
    if (!periods || periods.length === 0) {
        console.warn('⚠️ Nenhum período disponível');
        return;
    }
    
    const startSelect = document.getElementById('exportStartPeriod');
    const endSelect = document.getElementById('exportEndPeriod');
    
    if (!startSelect || !endSelect) {
        console.error('❌ Seletores de período não encontrados');
        return;
    }
    
    console.log(`📅 Populando períodos para tipo: ${exportConfig.periodType}, modo: ${exportConfig.comparisonMode}`);
    
    // Limpar opções
    startSelect.innerHTML = '<option value="">Selecione o período inicial...</option>';
    endSelect.innerHTML = '<option value="">Selecione o período final...</option>';
    
    // Gerar grupos de períodos baseado no tipo selecionado
    const groupedPeriods = groupPeriodsByTypeForSelector(exportConfig.periodType);
    const sortedKeys = sortGroupedPeriods(groupedPeriods, exportConfig.periodType);
    
    console.log(`📊 Grupos gerados: ${sortedKeys.length}`, sortedKeys);
    
    // Popular seletores com grupos
    sortedKeys.forEach(groupKey => {
        const option1 = document.createElement('option');
        option1.value = groupKey;
        option1.textContent = groupKey;
        startSelect.appendChild(option1);
        
        const option2 = document.createElement('option');
        option2.value = groupKey;
        option2.textContent = groupKey;
        endSelect.appendChild(option2);
    });
    
    // Auto-selecionar baseado no modo
    if (sortedKeys.length > 0) {
        if (exportConfig.comparisonMode === 'comparative' && sortedKeys.length >= 2) {
            // No modo comparativo, selecionar primeiro e último automaticamente
            startSelect.value = sortedKeys[0];
            endSelect.value = sortedKeys[sortedKeys.length - 1];
            exportConfig.startPeriod = sortedKeys[0];
            exportConfig.endPeriod = sortedKeys[sortedKeys.length - 1];
        } else {
            // No modo sequencial, selecionar primeiro e último para range
            startSelect.value = sortedKeys[0];
            endSelect.value = sortedKeys[sortedKeys.length - 1];
            exportConfig.startPeriod = sortedKeys[0];
            exportConfig.endPeriod = sortedKeys[sortedKeys.length - 1];
        }
    }
    
    console.log(`✅ ${sortedKeys.length} grupos de períodos populados`);
}

// Nova função para gerar grupos apenas para o seletor
function groupPeriodsByTypeForSelector(periodType) {
    const groupedPeriods = {};
    
    periods.forEach(period => {
        try {
            const [day, month, year] = period.split('/');
            const date = new Date(year, month - 1, day);
            let groupKey;
            
            switch (periodType) {
                case 'mensal':
                    groupKey = `${String(month).padStart(2, '0')}/${year}`;
                    break;
                    
                case 'bimestral':
                    const bimester = Math.ceil(parseInt(month) / 2);
                    groupKey = `${bimester}º Bim ${year}`;
                    break;
                    
                case 'trimestral':
                    const quarter = Math.ceil(parseInt(month) / 3);
                    groupKey = `${quarter}º Trim ${year}`;
                    break;
                    
                case 'semestral':
                    const semester = parseInt(month) <= 6 ? 1 : 2;
                    groupKey = `${semester}º Sem ${year}`;
                    break;
                    
                case 'anual':
                    groupKey = year;
                    break;
                    
                default:
                    groupKey = period;
            }
            
            if (!groupedPeriods[groupKey]) {
                groupedPeriods[groupKey] = [];
            }
            groupedPeriods[groupKey].push(period);
            
        } catch (error) {
            console.error('Erro ao agrupar período:', period, error);
        }
    });
    
    return groupedPeriods;
}

// ============================================
// LÓGICA DE AGRUPAMENTO POR PERÍODO
// ============================================

function updateComparisonModeInterface() {
    // No modo comparativo, MANTER os botões de período habilitados
    const periodButtons = document.querySelector('.export-period-section');
    
    if (periodButtons) {
        // Sempre manter habilitado - tanto sequencial quanto comparativo
        periodButtons.style.opacity = '1';
        periodButtons.style.pointerEvents = 'auto';
    }
    
    // Atualizar labels dos seletores baseado no modo
    const startLabel = document.querySelector('label[for="exportStartPeriod"]');
    const endLabel = document.querySelector('label[for="exportEndPeriod"]');
    
    if (startLabel && endLabel) {
        if (exportConfig.comparisonMode === 'comparative') {
            startLabel.innerHTML = '<i class="fas fa-calendar-check"></i> Primeiro Período';
            endLabel.innerHTML = '<i class="fas fa-calendar-times"></i> Segundo Período';
        } else {
            startLabel.innerHTML = '<i class="fas fa-calendar-plus"></i> Período Inicial';
            endLabel.innerHTML = '<i class="fas fa-calendar-minus"></i> Período Final';
        }
    }
    
    // Atualizar os seletores de período baseado no tipo atual
    populateAvailablePeriods();
    
    console.log(`🔄 Interface atualizada para modo: ${exportConfig.comparisonMode}, tipo: ${exportConfig.periodType}`);
}

function groupPeriodsByType(periodType, startPeriod, endPeriod) {
    console.log(`📊 Agrupando períodos: ${periodType} de ${startPeriod} até ${endPeriod}`);
    console.log(`🔍 Modo de comparação: ${exportConfig.comparisonMode}`);
    
    // MODO COMPARATIVO: Apenas os 2 períodos selecionados
    if (exportConfig.comparisonMode === 'comparative') {
        console.log('📋 Modo COMPARATIVO - apenas 2 períodos agrupados');
        
        const groupedPeriods = {};
        
        // Usar os grupos selecionados diretamente
        const allGroupedPeriods = groupPeriodsByTypeForSelector(periodType);
        
        // Verificar se os períodos selecionados existem nos grupos
        if (allGroupedPeriods[startPeriod]) {
            groupedPeriods[startPeriod] = allGroupedPeriods[startPeriod];
        }
        
        if (allGroupedPeriods[endPeriod] && startPeriod !== endPeriod) {
            groupedPeriods[endPeriod] = allGroupedPeriods[endPeriod];
        }
        
        console.log(`✅ Modo comparativo: ${Object.keys(groupedPeriods).length} grupos de períodos`);
        return groupedPeriods;
    }
    
    // MODO SEQUENCIAL: Agrupar períodos no range
    console.log('📋 Modo SEQUENCIAL - períodos consecutivos agrupados');
    
    const allGroupedPeriods = groupPeriodsByTypeForSelector(periodType);
    const allGroupKeys = sortGroupedPeriods(allGroupedPeriods, periodType);
    
    // Encontrar índices dos grupos de início e fim
    const startIndex = allGroupKeys.indexOf(startPeriod);
    const endIndex = allGroupKeys.indexOf(endPeriod);
    
    if (startIndex === -1 || endIndex === -1) {
        console.error('❌ Períodos de início ou fim não encontrados');
        return allGroupedPeriods;
    }
    
    // Selecionar apenas os grupos no range
    const rangeGroupKeys = allGroupKeys.slice(startIndex, endIndex + 1);
    const rangeGroupedPeriods = {};
    
    rangeGroupKeys.forEach(groupKey => {
        rangeGroupedPeriods[groupKey] = allGroupedPeriods[groupKey];
    });
    
    console.log(`✅ Grupos no range: ${Object.keys(rangeGroupedPeriods).length}`);
    return rangeGroupedPeriods;
}

function sortGroupedPeriods(groupedPeriods, periodType) {
    const sortedKeys = Object.keys(groupedPeriods).sort((a, b) => {
        try {
            switch (periodType) {
                case 'mensal':
                    const [monthA, yearA] = a.split('/');
                    const [monthB, yearB] = b.split('/');
                    const dateA = new Date(yearA, monthA - 1);
                    const dateB = new Date(yearB, monthB - 1);
                    return dateA.getTime() - dateB.getTime();
                    
                case 'bimestral':
                case 'trimestral':
                case 'semestral':
                    const regex = /(\d+)º\s\w+\s(\d+)/;
                    const matchA = a.match(regex);
                    const matchB = b.match(regex);
                    if (matchA && matchB) {
                        const yearDiff = parseInt(matchA[2]) - parseInt(matchB[2]);
                        if (yearDiff !== 0) return yearDiff;
                        return parseInt(matchA[1]) - parseInt(matchB[1]);
                    }
                    return a.localeCompare(b);
                    
                case 'anual':
                    return parseInt(a) - parseInt(b);
                    
                default:
                    return a.localeCompare(b);
            }
        } catch (error) {
            console.error('Erro na ordenação:', error);
            return a.localeCompare(b);
        }
    });
    
    return sortedKeys;
}

// ============================================
// CÁLCULOS DE TOTAIS E VARIAÇÕES
// ============================================

function calculateConsolidatedTotals(groupedPeriods, sortedKeys) {
    console.log('🧮 Calculando totais consolidados...');
    
    const consolidatedData = {};
    
    // Para cada conta, calcular totais por período agrupado
    rawData.forEach(item => {
        const { contaSAP, description } = extractAccountInfo(item);
        
        if (!contaSAP && !description) return;
        
        const accountKey = contaSAP || description;
        
        if (!consolidatedData[accountKey]) {
            consolidatedData[accountKey] = {
                contaSAP: contaSAP,
                description: description,
                values: {}
            };
        }
        
        // Para cada grupo de períodos
        sortedKeys.forEach(groupKey => {
            const periodsInGroup = groupedPeriods[groupKey];
            
            consolidatedData[accountKey].values[groupKey] = {
                telecom: 0,
                vogel: 0,
                consolidado: 0
            };
            
            // Somar valores de todos os períodos do grupo
            periodsInGroup.forEach(period => {
                if (item.values && item.values[period]) {
                    consolidatedData[accountKey].values[groupKey].telecom += item.values[period].telecom || 0;
                    consolidatedData[accountKey].values[groupKey].vogel += item.values[period].vogel || 0;
                    consolidatedData[accountKey].values[groupKey].consolidado += item.values[period].consolidado || 0;
                }
            });
        });
    });
    
    console.log(`✅ Totais calculados para ${Object.keys(consolidatedData).length} contas`);
    return consolidatedData;
}

// ============================================
// PREVIEW DO EXPORT
// ============================================

function generateExportPreview() {
    console.log('👀 Gerando preview do export...');
    
    if (!exportConfig.startPeriod || !exportConfig.endPeriod) {
        console.warn('⚠️ Períodos não selecionados');
        return;
    }
    
    try {
        const groupedPeriods = groupPeriodsByType(
            exportConfig.periodType, 
            exportConfig.startPeriod, 
            exportConfig.endPeriod
        );
        
        const sortedKeys = sortGroupedPeriods(groupedPeriods, exportConfig.periodType);
        const consolidatedData = calculateConsolidatedTotals(groupedPeriods, sortedKeys);
        
        // Calcular estatísticas do preview
        const totalAccounts = Object.keys(consolidatedData).length;
        const totalPeriods = sortedKeys.length;
        
        let dateRange, modeDescription, variationInfo = '';
        
        if (exportConfig.comparisonMode === 'comparative') {
            dateRange = `${formatPeriod(exportConfig.startPeriod)} vs ${formatPeriod(exportConfig.endPeriod)}`;
            modeDescription = 'Comparação entre 2 períodos específicos';
            
            if (exportConfig.includeVariations && totalPeriods >= 2) {
                variationInfo = ` | ${totalAccounts} variações calculadas`;
            }
        } else {
            dateRange = `${formatPeriod(exportConfig.startPeriod)} a ${formatPeriod(exportConfig.endPeriod)}`;
            modeDescription = `Análise ${exportConfig.periodType} sequencial`;
            
            if (exportConfig.includeVariations && totalPeriods > 1) {
                const totalVariations = totalAccounts * (totalPeriods - 1);
                variationInfo = ` | ${totalVariations} variações calculadas`;
            }
        }
        
        // Exemplo de algumas contas para mostrar
        const sampleAccounts = Object.keys(consolidatedData).slice(0, 3);
        const sampleData = sampleAccounts.map(key => {
            const account = consolidatedData[key];
            
            if (exportConfig.comparisonMode === 'comparative' && sortedKeys.length >= 2) {
                const firstValue = account.values[sortedKeys[0]]?.consolidado || 0;
                const secondValue = account.values[sortedKeys[1]]?.consolidado || 0;
                const variation = secondValue - firstValue;
                const variationPercent = firstValue !== 0 ? (variation / firstValue) * 100 : 0;
                
                return {
                    conta: account.contaSAP || key,
                    description: account.description || '-',
                    firstValue: firstValue,
                    secondValue: secondValue,
                    variation: variation,
                    variationPercent: variationPercent
                };
            } else {
                const firstPeriodValue = account.values[sortedKeys[0]]?.consolidado || 0;
                return {
                    conta: account.contaSAP || key,
                    description: account.description || '-',
                    firstValue: firstPeriodValue
                };
            }
        });
        
        // Atualizar preview
        const previewElement = document.getElementById('exportPreview');
        const previewContent = document.getElementById('exportPreviewContent');
        
        if (!previewElement || !previewContent) {
            console.error('❌ Elementos de preview não encontrados');
            return;
        }
        
        previewContent.innerHTML = `
            <div class="preview-stats">
                <div class="preview-stat-item">
                    <i class="fas fa-list-alt"></i>
                    <div>
                        <span class="stat-label">Total de Contas</span>
                        <span class="stat-value">${totalAccounts.toLocaleString()}</span>
                    </div>
                </div>
                <div class="preview-stat-item">
                    <i class="fas fa-calendar-alt"></i>
                    <div>
                        <span class="stat-label">Períodos</span>
                        <span class="stat-value">${totalPeriods}</span>
                    </div>
                </div>
                <div class="preview-stat-item">
                    <i class="fas fa-${exportConfig.comparisonMode === 'comparative' ? 'balance-scale' : 'clock'}"></i>
                    <div>
                        <span class="stat-label">Modo</span>
                        <span class="stat-value">${exportConfig.comparisonMode === 'comparative' ? 'Comparativo' : 'Sequencial'}</span>
                    </div>
                </div>
            </div>
            
            <div class="preview-summary">
                <h4><i class="fas fa-info-circle"></i> Resumo do Export</h4>
                <p><strong>Tipo:</strong> ${modeDescription}</p>
                <p><strong>Formato:</strong> ${exportConfig.exportFormat.toUpperCase()}</p>
                <p><strong>Período:</strong> ${dateRange}</p>
                <p><strong>Dados:</strong> ${totalAccounts} contas × ${totalPeriods} períodos = ${(totalAccounts * totalPeriods).toLocaleString()} valores${variationInfo}</p>
                <p><strong>Períodos:</strong> ${sortedKeys.join(', ')}</p>
            </div>
            
            ${sampleData.length > 0 ? `
            <div class="preview-sample">
                <h4><i class="fas fa-eye"></i> Amostra dos Dados</h4>
                <div class="sample-table">
                    <table>
                        <thead>
                            <tr>
                                <th>Conta SAP</th>
                                <th>Descrição</th>
                                ${exportConfig.comparisonMode === 'comparative' && sortedKeys.length >= 2 ? `
                                    <th>${sortedKeys[0]}</th>
                                    <th>${sortedKeys[1]}</th>
                                    ${exportConfig.includeVariations ? '<th>Variação</th><th>Var %</th>' : ''}
                                ` : `
                                    <th>${sortedKeys[0]} (Somado)</th>
                                `}
                            </tr>
                        </thead>
                        <tbody>
                            ${sampleData.map(sample => `
                                <tr>
                                    <td><strong>${sample.conta}</strong></td>
                                    <td>${sample.description}</td>
                                    ${exportConfig.comparisonMode === 'comparative' && sortedKeys.length >= 2 ? `
                                        <td>${formatCurrency(sample.firstValue)}</td>
                                        <td>${formatCurrency(sample.secondValue)}</td>
                                        ${exportConfig.includeVariations ? `
                                            <td class="variation-cell ${sample.variation >= 0 ? 'positive' : 'negative'}">
                                                ${formatCurrency(sample.variation)}
                                            </td>
                                            <td class="variation-cell ${sample.variationPercent >= 0 ? 'positive' : 'negative'}">
                                                ${sample.variationPercent >= 0 ? '+' : ''}${sample.variationPercent.toFixed(2)}%
                                            </td>
                                        ` : ''}
                                    ` : `
                                        <td>${formatCurrency(sample.firstValue)}</td>
                                    `}
                                </tr>
                            `).join('')}
                            ${sampleData.length < totalAccounts ? `
                                <tr class="more-indicator">
                                    <td colspan="${exportConfig.comparisonMode === 'comparative' ? (exportConfig.includeVariations ? '6' : '4') : '3'}">
                                        <em>... e mais ${totalAccounts - sampleData.length} contas</em>
                                    </td>
                                </tr>
                            ` : ''}
                        </tbody>
                    </table>
                </div>
            </div>
            ` : ''}
        `;
        
        previewElement.style.display = 'block';
        
        console.log('✅ Preview gerado com sucesso');
        
    } catch (error) {
        console.error('❌ Erro ao gerar preview:', error);
        const previewElement = document.getElementById('exportPreview');
        if (previewElement) {
            previewElement.innerHTML = '<div class="error-message">❌ Erro ao gerar preview</div>';
            previewElement.style.display = 'block';
        }
    }
}

// ============================================
// EXECUÇÃO DO EXPORT
// ============================================

function executeConsolidatedExport() {
    console.log('🚀 === EXECUTANDO EXPORT CONSOLIDADO ===');
    
    if (!exportConfig.startPeriod || !exportConfig.endPeriod) {
        alert('❌ Selecione os períodos de início e fim');
        return;
    }
    
    if (!selectedSheet || !rawData || rawData.length === 0) {
        alert('❌ Nenhum dado disponível para exportar');
        return;
    }
    
    try {
        showLoading(true);
        
        // Aguardar um pouco para mostrar o loading
        setTimeout(() => {
            executeExportProcess();
        }, 100);
        
    } catch (error) {
        console.error('❌ Erro no export:', error);
        alert('❌ Erro ao executar export. Tente novamente.');
        showLoading(false);
    }
}

function executeExportProcess() {
    try {
        console.log('📊 Processando dados para export...');
        
        // 1. Agrupar períodos
        const groupedPeriods = groupPeriodsByType(
            exportConfig.periodType, 
            exportConfig.startPeriod, 
            exportConfig.endPeriod
        );
        
        const sortedKeys = sortGroupedPeriods(groupedPeriods, exportConfig.periodType);
        
        if (sortedKeys.length === 0) {
            throw new Error('Nenhum período válido encontrado no range selecionado');
        }
        
        // 2. Calcular totais consolidados
        const consolidatedData = calculateConsolidatedTotals(groupedPeriods, sortedKeys);
        
        if (Object.keys(consolidatedData).length === 0) {
            throw new Error('Nenhuma conta com dados encontrada');
        }
        
        // 3. Gerar Excel
        generateConsolidatedExcel(consolidatedData, sortedKeys, groupedPeriods);
        
        // 4. Fechar modal e loading
        const modal = document.getElementById('consolidatedExportModal');
        if (modal) modal.classList.remove('active');
        showLoading(false);
        
        console.log('✅ Export consolidado concluído com sucesso');
        
    } catch (error) {
        console.error('❌ Erro no processo de export:', error);
        alert(`❌ Erro no export: ${error.message}`);
        showLoading(false);
    }
}

function generateConsolidatedExcel(consolidatedData, sortedKeys, groupedPeriods) {
    console.log('📋 Gerando arquivo Excel consolidado...');
    
    const wb = XLSX.utils.book_new();
    
    // Gerar apenas as abas detalhadas (sem resumo e metadados)
    const entities = [
        { key: 'telecom', name: 'Telecom' },
        { key: 'vogel', name: 'Vogel' },
        { key: 'consolidado', name: 'Somado' }
    ];
    
    entities.forEach(entity => {
        const data = generateDetailedSheet(consolidatedData, sortedKeys, entity.key, entity.name);
        const ws = XLSX.utils.aoa_to_sheet(data);
        formatDetailedSheet(ws, sortedKeys, true); // Sempre incluir variações
        XLSX.utils.book_append_sheet(wb, ws, entity.name);
    });
    
    // Propriedades do workbook
    wb.Props = {
        Title: `Export Consolidado - ${selectedSheet}`,
        Subject: `Análise ${exportConfig.periodType} - Algar`,
        Author: 'Dashboard Analytics Algar',
        CreatedDate: new Date(),
        Category: 'Relatório Financeiro'
    };
    
    // Salvar arquivo
    const filename = generateExportFilename();
    XLSX.writeFile(wb, filename);
    
    console.log('✅ Arquivo Excel gerado:', filename);
}

function generateSummarySheet(consolidatedData, sortedKeys) {
    const data = [];
    
    // Cabeçalho
    data.push(['RESUMO EXECUTIVO - TOTAIS POR PERÍODO']);
    data.push(['']);
    data.push([
        'Período',
        'Telecom (R$)',
        'Vogel (R$)',
        'Somado (R$)',
        'Var Telecom (%)',
        'Var Vogel (%)',
        'Var Somado (%)'
    ]);
    
    // Calcular totais por período
    let previousTotals = null;
    
    sortedKeys.forEach((periodKey, index) => {
        let telecomTotal = 0;
        let vogelTotal = 0;
        let consolidadoTotal = 0;
        
        // Somar todos os valores das contas para este período
        Object.keys(consolidatedData).forEach(accountKey => {
            const account = consolidatedData[accountKey];
            if (account.values[periodKey]) {
                telecomTotal += account.values[periodKey].telecom || 0;
                vogelTotal += account.values[periodKey].vogel || 0;
                consolidadoTotal += account.values[periodKey].consolidado || 0;
            }
        });
        
        // Calcular variações
        let telecomVar = '';
        let vogelVar = '';
        let consolidadoVar = '';
        
        if (exportConfig.includeVariations && previousTotals) {
            telecomVar = previousTotals.telecom !== 0 ? 
                ((telecomTotal - previousTotals.telecom) / previousTotals.telecom) : 0;
            vogelVar = previousTotals.vogel !== 0 ? 
                ((vogelTotal - previousTotals.vogel) / previousTotals.vogel) : 0;
            consolidadoVar = previousTotals.consolidado !== 0 ? 
                ((consolidadoTotal - previousTotals.consolidado) / previousTotals.consolidado) : 0;
        }
        
        data.push([
            periodKey,
            telecomTotal,
            vogelTotal,
            consolidadoTotal,
            telecomVar,
            vogelVar,
            consolidadoVar
        ]);
        
        previousTotals = { telecom: telecomTotal, vogel: vogelTotal, consolidado: consolidadoTotal };
    });
    
    return data;
}

function generateDetailedSheet(consolidatedData, sortedKeys, entity, entityName) {
    const data = [];
    
    // Cabeçalho
    data.push([`DETALHAMENTO ${entityName.toUpperCase()} POR CONTA`]);
    data.push(['']);
    
    // Headers das colunas baseado no modo
    const headers = ['Conta SAP', 'Descrição'];
    
    if (exportConfig.comparisonMode === 'comparative') {
        // MODO COMPARATIVO: Apenas 2 períodos + 1 variação
        if (sortedKeys.length >= 2) {
            headers.push(sortedKeys[0]); // Primeiro período
            headers.push(sortedKeys[1]); // Segundo período
            if (exportConfig.includeVariations) {
                headers.push('Variação (R$)');
                headers.push('Variação (%)');
            }
        } else {
            // Fallback se só tem 1 período
            headers.push(...sortedKeys);
        }
    } else {
        // MODO SEQUENCIAL: Períodos + variações consecutivas
        if (exportConfig.includeVariations) {
            // Formato: Período1 | Período2 | Var% | Período3 | Var% | ...
            sortedKeys.forEach((period, index) => {
                headers.push(period);
                if (index > 0) {
                    headers.push('Var %');
                }
            });
        } else {
            // Formato simples: Período1 | Período2 | ...
            headers.push(...sortedKeys);
        }
    }
    
    data.push(headers);
    
    // Dados das contas
    Object.keys(consolidatedData).forEach(accountKey => {
        const account = consolidatedData[accountKey];
        const row = [
            account.contaSAP || accountKey,
            account.description || '-'
        ];
        
        if (exportConfig.comparisonMode === 'comparative') {
            // MODO COMPARATIVO
            if (sortedKeys.length >= 2) {
                const firstValue = account.values[sortedKeys[0]][entity] || 0;
                const secondValue = account.values[sortedKeys[1]][entity] || 0;
                
                row.push(firstValue);  // Primeiro período
                row.push(secondValue); // Segundo período
                
                if (exportConfig.includeVariations) {
                    const variation = secondValue - firstValue;
                    const variationPercent = firstValue !== 0 ? (variation / firstValue) : 0;
                    
                    row.push(variation);        // Variação absoluta
                    row.push(variationPercent); // Variação percentual
                }
            } else {
                // Fallback para 1 período apenas
                sortedKeys.forEach(period => {
                    row.push(account.values[period][entity] || 0);
                });
            }
            
        } else {
            // MODO SEQUENCIAL (comportamento original)
            if (exportConfig.includeVariations) {
                let previousValue = null;
                
                sortedKeys.forEach((period, index) => {
                    const currentValue = account.values[period][entity] || 0;
                    row.push(currentValue);
                    
                    if (index > 0 && previousValue !== null) {
                        const variation = previousValue !== 0 ? 
                            ((currentValue - previousValue) / previousValue) : 0;
                        row.push(variation);
                    }
                    
                    previousValue = currentValue;
                });
            } else {
                sortedKeys.forEach(period => {
                    row.push(account.values[period][entity] || 0);
                });
            }
        }
        
        data.push(row);
    });
    
    return data;
}

function formatDetailedSheet(worksheet, sortedKeys, includeVariations) {
    const currencyFormat = '"R$ "#,##0.00_);("R$ "#,##0.00)';
    const percentFormat = '0.00%';
    
    console.log(`🔧 === FORMATANDO PLANILHA ===`);
    console.log(`Modo: ${exportConfig.comparisonMode}`);
    console.log(`Períodos: ${sortedKeys.length}`);
    console.log(`Incluir variações: ${includeVariations}`);
    
    // Larguras das colunas
    const colWidths = [
        { wch: 15 }, // Conta SAP
        { wch: 40 }  // Descrição
    ];
    
    // Mapear estrutura das colunas para saber qual é qual
    const columnStructure = []; // Array que vai indicar o tipo de cada coluna
    
    if (includeVariations) {
        if (exportConfig.comparisonMode === 'comparative') {
            // Modo comparativo: Período1 | Período2 | Variação(R$) | Variação(%)
            columnStructure.push('period', 'period', 'variation_currency', 'variation_percent');
            colWidths.push({ wch: 18 }, { wch: 18 }, { wch: 15 }, { wch: 12 });
        } else {
            // Modo sequencial: Período1 | Período2 | Var% | Período3 | Var% | ...
            sortedKeys.forEach((_, index) => {
                columnStructure.push('period'); // Primeiro o período
                colWidths.push({ wch: 18 });
                
                if (index > 0) { // Depois a variação (exceto no primeiro período)
                    columnStructure.push('variation_percent');
                    colWidths.push({ wch: 12 });
                }
            });
        }
    } else {
        // Apenas períodos
        sortedKeys.forEach(() => {
            columnStructure.push('period');
            colWidths.push({ wch: 18 });
        });
    }
    
    console.log(`📊 Estrutura das colunas:`, columnStructure);
    
    worksheet['!cols'] = colWidths;
    
    // Aplicar formatação de moeda e porcentagem
    if (!worksheet['!ref']) return;
    
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    
    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            
            if (!worksheet[cellAddress]) continue;
            
            // Pular cabeçalhos (primeiras 3 linhas)
            if (R < 3) continue;
            
            // Formatação baseada na estrutura das colunas
            if (C >= 2) { // Colunas de dados (após Conta SAP e Descrição)
                const dataColumnIndex = C - 2; // Índice na estrutura de dados
                const columnType = columnStructure[dataColumnIndex];
                
                if (R === 3 && C <= 7) {
                    console.log(`  Col ${C} (dados ${dataColumnIndex}): tipo "${columnType}" = ${worksheet[cellAddress].v}`);
                }
                
                switch (columnType) {
                    case 'period':
                    case 'variation_currency':
                        worksheet[cellAddress].z = currencyFormat;
                        break;
                    case 'variation_percent':
                        worksheet[cellAddress].z = percentFormat;
                        break;
                    default:
                        console.warn(`⚠️ Tipo de coluna desconhecido: ${columnType}`);
                        worksheet[cellAddress].z = currencyFormat;
                }
            }
        }
    }
    
    console.log(`✅ Formatação aplicada com sucesso`);
}

function generateMetadataSheet(groupedPeriods, sortedKeys) {
    const data = [];
    
    data.push(['METADADOS DO EXPORT']);
    data.push(['']);
    data.push(['Configuração', 'Valor']);
    data.push(['Data do Export', new Date().toLocaleString('pt-BR')]);
    data.push(['Aba Selecionada', selectedSheet]);
    data.push(['Tipo de Período', exportConfig.periodType]);
    data.push(['Período Inicial', formatPeriod(exportConfig.startPeriod)]);
    data.push(['Período Final', formatPeriod(exportConfig.endPeriod)]);
    data.push(['Modo de Comparação', exportConfig.comparisonMode || 'sequencial']);
    data.push(['Incluir Variações', exportConfig.includeVariations ? 'Sim' : 'Não']);
    data.push(['Formato', exportConfig.exportFormat.toUpperCase()]);
    data.push(['Total de Contas', Object.keys(groupedPeriods).length]);
    data.push(['Total de Períodos', sortedKeys.length]);
    data.push(['']);
    data.push(['PERÍODOS AGRUPADOS']);
    data.push(['Grupo', 'Períodos Originais']);
    
    Object.keys(groupedPeriods).forEach(groupKey => {
        const periods = groupedPeriods[groupKey];
        data.push([groupKey, periods.map(p => formatPeriod(p)).join(', ')]);
    });
    
    return data;
}

function generateExportFilename() {
    const timestamp = new Date().toISOString().split('T')[0];
    const periodType = exportConfig.periodType;
    const sheetName = selectedSheet.replace(/[^a-zA-Z0-9]/g, '');
    const startPeriod = exportConfig.startPeriod.replace(/\//g, '');
    const endPeriod = exportConfig.endPeriod.replace(/\//g, '');
    const mode = exportConfig.comparisonMode === 'comparative' ? 'comparativo' : 'sequencial';
    
    return `algar-export-consolidado-${sheetName}-${periodType}-${mode}-${startPeriod}-${endPeriod}-${timestamp}.xlsx`;
}

// ============================================
// INICIALIZAÇÃO DOS EVENT LISTENERS
// ============================================

function initializeExportModal() {
    // Event listeners para os botões de período
    document.querySelectorAll('.export-period-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            // Agora permite mudança de período em ambos os modos
            document.querySelectorAll('.export-period-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            exportConfig.periodType = btn.dataset.periodType;
            
            console.log(`🔄 Tipo de período alterado para: ${exportConfig.periodType}`);
            
            // Atualizar interface e repopular períodos
            updateComparisonModeInterface();
        });
    });
    
    // Event listeners para os botões de modo de comparação
    document.querySelectorAll('.comparison-mode-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.comparison-mode-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            exportConfig.comparisonMode = btn.dataset.comparisonMode;
            
            // Atualizar interface baseado no modo
            updateComparisonModeInterface();
        });
    });
    
    // Event listeners para seleção de períodos
    const startPeriodEl = document.getElementById('exportStartPeriod');
    const endPeriodEl = document.getElementById('exportEndPeriod');
    const clearBtn = document.getElementById('clearExportBtn');
    const executeBtn = document.getElementById('executeExportBtn');
    const closeBtn = document.getElementById('closeConsolidatedExportModal');
    
    if (startPeriodEl) {
        startPeriodEl.addEventListener('change', (e) => {
            exportConfig.startPeriod = e.target.value;
            console.log(`📅 Período inicial selecionado: ${e.target.value}`);
        });
    }
    
    if (endPeriodEl) {
        endPeriodEl.addEventListener('change', (e) => {
            exportConfig.endPeriod = e.target.value;
            console.log(`📅 Período final selecionado: ${e.target.value}`);
        });
    }
    
    if (clearBtn) {
        clearBtn.addEventListener('click', () => {
            resetExportConfig();
        });
    }
    
    if (executeBtn) {
        executeBtn.addEventListener('click', () => {
            executeConsolidatedExport();
        });
    }
    
    if (closeBtn) {
        closeBtn.addEventListener('click', () => {
            const modal = document.getElementById('consolidatedExportModal');
            if (modal) modal.classList.remove('active');
        });
    }
    
    // Fechar modal clicando fora
    const modal = document.getElementById('consolidatedExportModal');
    if (modal) {
        modal.addEventListener('click', (e) => {
            if (e.target.id === 'consolidatedExportModal') {
                modal.classList.remove('active');
            }
        });
    }
    
    console.log('✅ Event listeners do modal de export inicializados');
}

// ============================================
// INICIALIZAÇÃO AUTOMÁTICA
// ============================================

// Verificar se o DOM já está carregado ou aguardar
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeExportSystem);
} else {
    // DOM já carregado, inicializar imediatamente
    setTimeout(() => {
        initializeExportSystem();
    }, 1000);
}

// Adicionar função global para compatibilidade
window.initializeExportSystem = initializeExportSystem;

console.log('✅ Script3.js - Sistema de Export Consolidado carregado');