let workbook = null;
let allSheetsData = {};
let selectedSheet = '';
let rawData = [];
let filteredData = [];
let periods = [];
let charts = {};
let chartViewMode = 'monthly';

// Sheet filter configurations - ATUALIZADO para incluir filtro de conta especial
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

// NOVA FUNÇÃO: Extrair informações da conta
function extractAccountInfo(item) {
    let contaSAP = '';
    let description = '';
    
    for (const key of Object.keys(item)) {
        const keyTrimmed = key.trim().replace(/\s+/g, ' ');
        const keyNoSpaces = key.replace(/\s/g, '');
        
        // Verificar se é conta SAP ou número conta
        if (keyTrimmed === 'Conta SAP' || keyTrimmed === 'CONTA SAP' || 
            keyTrimmed === 'Número Conta' || keyTrimmed === 'NÚMERO CONTA' ||
            keyNoSpaces.toLowerCase() === 'contasap' || 
            keyNoSpaces.toLowerCase() === 'numeroconta') {
            contaSAP = String(item[key] || '');
        }
        
        // Verificar descrição
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
    
    // BOTÃO APLICAR - com debug
    const applyBtn = document.getElementById('applyFilters');
    if (applyBtn) {
        applyBtn.addEventListener('click', () => {
            console.log('🔵 BOTÃO APLICAR CLICADO!');
            applyFilters();
        });
        console.log('✅ Event listener do botão Aplicar configurado');
    } else {
        console.error('❌ Botão applyFilters não encontrado!');
    }
    
    document.getElementById('clearFilters').addEventListener('click', () => clearFilters());
    document.getElementById('currentPeriod').addEventListener('change', () => updateDashboard());
    document.getElementById('previousPeriod').addEventListener('change', () => updateDashboard());
    document.getElementById('sheetSelector').addEventListener('change', handleSheetChange);
    document.getElementById('monthlyBtn').addEventListener('click', () => toggleChartView('monthly'));
    document.getElementById('quarterlyBtn').addEventListener('click', () => toggleChartView('quarterly'));
    document.getElementById('uploadButton').addEventListener('click', () => showUploadModal(true));
    document.getElementById('closeModal').addEventListener('click', () => showUploadModal(false));
    document.getElementById('descriptionFilter').addEventListener('input', filterDetailTable);
    document.getElementById('sortFilter').addEventListener('change', filterDetailTable);
    document.getElementById('entityFilter').addEventListener('change', filterDetailTable);
    document.getElementById('exportTable').addEventListener('click', exportTableToExcel);
    document.getElementById('downloadChart').addEventListener('click', downloadChart);
    
    // Inicializar modal de contas
    initializeAccountModal();
    
    window.addEventListener('click', function(e) {
        const modal = document.getElementById('uploadModal');
        if (e.target === modal) {
            showUploadModal(false);
        }
    });
    
    window.addEventListener('resize', handleResize);
    handleResize();
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
                console.log('Workbook structure:', workbook);
                
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
                    console.log(`Sheet "${sheetName}" has ${jsonData.length} rows`);
                    if (jsonData.length > 0) {
                        console.log(`Headers for "${sheetName}":`, jsonData[0]);
                    }
                });
                
                resolve();
            } catch (error) {
                console.error('Error reading Excel file:', error);
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

// FUNÇÃO ATUALIZADA: Criar filtros dinâmicos com campo de busca especial
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
    
    console.log('Configuração de filtros:', filterConfig);
    
    filterConfig.forEach(config => {
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-group';
        
        if (config.type === 'search') {
            // Campo de busca especial para contas
            console.log(`Criando campo de busca: ${config.id}`);
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
            // Select normal para outros filtros
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
    
    // Verificar se os elementos foram criados
    setTimeout(() => {
        const contaInput = document.getElementById('filter_contaSearch');
        console.log('Campo criado:', {
            existe: !!contaInput,
            id: contaInput?.id,
            dataField: contaInput?.dataset?.field,
            placeholder: contaInput?.placeholder
        });
    }, 100);
    
    // Adicionar event listeners
    filterConfig.forEach(config => {
        if (config.type === 'search') {
            setTimeout(() => {
                console.log(`Configurando busca para: ${config.id}`);
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

// NOVA FUNÇÃO: Configurar busca de conta com autocomplete
function setupAccountSearch(filterId) {
    const input = document.getElementById(`filter_${filterId}`);
    const clearBtn = document.getElementById(`clear_${filterId}`);
    const suggestions = document.getElementById(`suggestions_${filterId}`);
    
    if (!input || !clearBtn || !suggestions) return;
    
    let accountOptions = [];
    
    // Gerar lista de contas quando os dados estiverem disponíveis
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
            // Ordenar por número da conta primeiro, depois por descrição
            if (a.contaSAP && b.contaSAP) {
                const contaA = String(a.contaSAP);
                const contaB = String(b.contaSAP);
                // Ordenação numérica se ambos forem números
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
    
    // Buscar e mostrar sugestões
    function showSuggestions(query) {
        if (!query || query.length < 2) {
            suggestions.innerHTML = '';
            suggestions.style.display = 'none';
            return;
        }
        
        const filtered = accountOptions.filter(option => 
            option.searchText.includes(query.toLowerCase())
        ).slice(0, 10); // Limitar a 10 sugestões
        
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
        
        // Adicionar event listeners nas sugestões
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
    
    // Event listeners
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
        // Delay para permitir clique nas sugestões
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
    
    // Gerar opções quando disponível
    if (rawData.length > 0) {
        generateAccountOptions();
    }
    
    // Armazenar função para usar depois
    input._generateAccountOptions = generateAccountOptions;
}

// NOVA FUNÇÃO: Abrir modal de seleção de contas (SIMPLIFICADA)
function openAccountModal(targetInput, accountOptions) {
    console.log('🔍 openAccountModal chamada:', { targetInput: targetInput?.id, accountOptionsLength: accountOptions?.length });
    
    // Verificar se temos dados disponíveis
    if (!rawData || rawData.length === 0) {
        alert('Selecione uma aba primeiro para visualizar as contas disponíveis.');
        return;
    }
    
    // Se accountOptions está vazio ou não existe, gerar agora
    if (!accountOptions || accountOptions.length === 0) {
        console.log('Gerando opções de conta...');
        accountOptions = generateAccountOptionsManually();
    }
    
    if (accountOptions.length === 0) {
        alert('Nenhuma conta encontrada nesta aba. Verifique se a aba selecionada contém dados de contas.');
        return;
    }
    
    console.log(`📋 Abrindo modal com ${accountOptions.length} contas`);
    
    // Usar o modal que já existe no HTML
    const modal = document.getElementById('accountModal');
    const searchInput = document.getElementById('accountSearch');
    const accountList = document.getElementById('accountList');
    
    if (!modal || !searchInput || !accountList) {
        console.error('❌ Elementos do modal não encontrados no HTML');
        return;
    }
    
    // Limpar busca anterior
    searchInput.value = '';
    
    // Função para renderizar lista de contas
    function renderAccountList(accounts = accountOptions) {
        accountList.innerHTML = '';
        
        if (accounts.length === 0) {
            accountList.innerHTML = '<div class="no-accounts">Nenhuma conta encontrada</div>';
            return;
        }
        
        accounts.forEach(account => {
            const item = document.createElement('div');
            item.className = 'account-item';
            item.innerHTML = `
                <div class="account-number">${account.contaSAP || '-'}</div>
                <div class="account-description">${account.description || '-'}</div>
            `;
            
            item.addEventListener('click', () => {
                console.log('✅ Conta selecionada:', account);
                
                // Preencher o campo de busca original
                targetInput.value = account.displayText;
                targetInput.dataset.selectedConta = account.contaSAP;
                targetInput.dataset.selectedDescription = account.description;
                
                // Mostrar botão limpar
                const clearBtn = document.getElementById(targetInput.id.replace('filter_', 'clear_'));
                if (clearBtn) clearBtn.style.display = 'inline-block';
                
                // Fechar modal
                modal.classList.remove('active');
                
                // Atualizar filtros
                updateCascadingFilters();
            });
            
            accountList.appendChild(item);
        });
    }
    
    // Configurar busca em tempo real no modal (remover listeners anteriores)
    const newSearchInput = searchInput.cloneNode(true);
    searchInput.parentNode.replaceChild(newSearchInput, searchInput);
    
    newSearchInput.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        
        if (!query) {
            renderAccountList();
            return;
        }
        
        const filtered = accountOptions.filter(account => 
            account.searchText && account.searchText.includes(query)
        );
        
        renderAccountList(filtered);
    });
    
    // Renderizar lista inicial
    renderAccountList();
    
    // Mostrar modal
    modal.classList.add('active');
    console.log('✅ Modal ativado');
    
    // Focar no campo de busca
    setTimeout(() => {
        newSearchInput.focus();
    }, 100);
}

// Função para configurar o modal quando a página carrega
function initializeAccountModal() {
    const modal = document.getElementById('accountModal');
    const closeBtn = document.getElementById('closeAccountModal');
    
    if (modal && closeBtn) {
        // Event listener para fechar modal
        closeBtn.addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        // Fechar modal ao clicar fora
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });
        
        console.log('✅ Modal de contas inicializado');
    } else {
        console.error('❌ Modal de contas não encontrado no HTML');
    }
}

// FUNÇÃO DE TESTE: Adicionar um botão de teste temporário
function addTestButton() {
    const testBtn = document.createElement('button');
    testBtn.innerText = 'TESTE MODAL';
    testBtn.style.position = 'fixed';
    testBtn.style.top = '10px';
    testBtn.style.right = '10px';
    testBtn.style.zIndex = '9999';
    testBtn.style.padding = '10px';
    testBtn.style.backgroundColor = 'red';
    testBtn.style.color = 'white';
    testBtn.style.border = 'none';
    testBtn.style.cursor = 'pointer';
    
    testBtn.addEventListener('click', () => {
        console.log('🔴 BOTÃO TESTE CLICADO');
        const modal = document.getElementById('accountModal');
        if (modal) {
            modal.classList.add('active');
            console.log('✅ Modal ativado via teste');
        } else {
            console.error('❌ Modal não encontrado');
        }
    });
    
    document.body.appendChild(testBtn);
    console.log('🔴 Botão de teste adicionado');
}

// FUNÇÃO SEPARADA: Configurar apenas a lupa para abrir modal
function setupAccountSearchIcon() {
    console.log('🔧 Configurando lupa de conta...');
    
    setTimeout(() => {
        const accountSearchInput = document.getElementById('filter_contaSearch');
        if (!accountSearchInput) {
            console.log('❌ Campo filter_contaSearch ainda não existe, tentando novamente...');
            setTimeout(setupAccountSearchIcon, 500);
            return;
        }
        
        console.log('✅ Campo filter_contaSearch encontrado');
        
        // Encontrar a lupa específica deste campo
        const searchContainer = accountSearchInput.parentElement;
        const searchIcon = searchContainer.querySelector('.fas.fa-search');
        
        if (!searchIcon) {
            console.error('❌ Lupa não encontrada no container do campo de conta');
            return;
        }
        
        console.log('✅ Lupa de conta encontrada');
        
        // IMPORTANTE: Não mexer nos event listeners existentes do input
        // Apenas configurar a lupa separadamente
        
        // Configurar apenas a lupa (sem afetar o input)
        searchIcon.style.cursor = 'pointer';
        searchIcon.style.pointerEvents = 'auto';
        searchIcon.title = 'Clique para ver todas as contas';
        
        // Remover event listeners anteriores APENAS da lupa
        const newSearchIcon = searchIcon.cloneNode(true);
        searchIcon.parentNode.replaceChild(newSearchIcon, searchIcon);
        
        // Event listener APENAS para a lupa (modal)
        newSearchIcon.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            console.log('🔍 LUPA DE CONTA CLICADA!');
            
            // Gerar opções independentemente do autocomplete
            if (!rawData || rawData.length === 0) {
                alert('Selecione uma aba primeiro!');
                return;
            }
            
            console.log('📋 Gerando lista de contas para modal...');
            const accountOptions = generateAccountOptionsManually();
            
            if (accountOptions.length === 0) {
                alert('Nenhuma conta encontrada nesta aba.');
                return;
            }
            
            console.log(`✅ ${accountOptions.length} contas encontradas, abrindo modal...`);
            openAccountModalWithData(accountSearchInput, accountOptions);
        });
        
        // Hover effect apenas na lupa
        newSearchIcon.addEventListener('mouseenter', () => {
            newSearchIcon.style.color = '#0066cc';
            newSearchIcon.style.backgroundColor = '#f0f8ff';
            newSearchIcon.style.borderRadius = '4px';
        });
        
        newSearchIcon.addEventListener('mouseleave', () => {
            newSearchIcon.style.color = '#6b7280';
            newSearchIcon.style.backgroundColor = 'transparent';
        });
        
        console.log('✅ Lupa de conta configurada com sucesso (independente do autocomplete)!');
        
    }, 100);
}

// FUNÇÃO SIMPLIFICADA PARA ABRIR MODAL COM DADOS
function openAccountModalWithData(targetInput, accountOptions) {
    console.log(`🔍 Abrindo modal com ${accountOptions.length} contas`);
    
    const modal = document.getElementById('accountModal');
    const searchInput = document.getElementById('accountSearch');
    const accountList = document.getElementById('accountList');
    
    if (!modal || !searchInput || !accountList) {
        console.error('❌ Elementos do modal não encontrados');
        return;
    }
    
    // Limpar busca anterior
    searchInput.value = '';
    
    // Função para renderizar lista
    function renderList(accounts = accountOptions) {
        accountList.innerHTML = '';
        
        if (accounts.length === 0) {
            accountList.innerHTML = '<div class="no-accounts">Nenhuma conta encontrada</div>';
            return;
        }
        
        console.log(`📋 Renderizando ${accounts.length} contas...`);
        
        accounts.forEach((account, index) => {
            const item = document.createElement('div');
            item.className = 'account-item';
            item.innerHTML = `
                <div class="account-number">${account.contaSAP || `Conta ${index + 1}`}</div>
                <div class="account-description">${account.description || 'Sem descrição'}</div>
            `;
            
            item.addEventListener('click', () => {
                console.log('✅ Conta selecionada:', account);
                
                // Preencher campo
                targetInput.value = account.displayText;
                targetInput.dataset.selectedConta = account.contaSAP;
                targetInput.dataset.selectedDescription = account.description;
                
                // Mostrar botão limpar
                const clearBtn = document.getElementById('clear_contaSearch');
                if (clearBtn) clearBtn.style.display = 'inline-block';
                
                // Fechar modal
                modal.classList.remove('active');
                
                // APLICAR FILTROS IMEDIATAMENTE
                console.log('🔄 Conta selecionada via modal, aplicando filtros automaticamente...');
                updateCascadingFilters();
                setTimeout(() => {
                    applyFilters();
                }, 100);
            });
            
            accountList.appendChild(item);
        });
    }
    
    // Configurar busca no modal
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
    
    // Renderizar lista inicial
    renderList();
    
    // Mostrar modal
    modal.classList.add('active');
    console.log('✅ Modal aberto com sucesso!');
    
    // Focar no campo de busca
    setTimeout(() => newSearchInput.focus(), 100);
}

// NOVA FUNÇÃO: Gerar opções de conta manualmente COM FILTROS CASCATA (LÓGICA ORIGINAL)
function generateAccountOptionsManually() {
    console.log('🔄 generateAccountOptionsManually iniciada');
    const accountOptions = [];
    const seen = new Set();
    
    if (!rawData || rawData.length === 0) {
        console.log('❌ Nenhum rawData disponível');
        return accountOptions;
    }
    
    // USAR EXATAMENTE A MESMA LÓGICA DO populateFilterOptions
    const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
    const currentFilters = {};
    
    dynamicFilters.forEach(filterSelect => {
        if (filterSelect.value) {
            currentFilters[filterSelect.dataset.field] = filterSelect.value;
        }
    });
    
    // Filter data based on other selections (CÓDIGO ORIGINAL)
    let filteredDataForOptions = rawData;
    
    Object.keys(currentFilters).forEach(currentField => {
        if (currentField !== 'CONTA_SEARCH') { // Não filtrar pelo próprio campo de conta
            filteredDataForOptions = filteredDataForOptions.filter(item => 
                item[currentField] === currentFilters[currentField]
            );
        }
    });
    
    console.log(`Modal: dados após cascata ${rawData.length} → ${filteredDataForOptions.length}`);
    
    // Gerar opções apenas dos dados filtrados
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
    
    console.log(`✅ Modal: geradas ${accountOptions.length} opções com cascata original`);
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
    
    console.log('Headers encontrados:', headers);
    
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
    
    console.log('Períodos encontrados:', periods);
    
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
    
    // 1. Primeiro, gerar opções para o autocomplete original
    console.log('🔄 Atualizando autocomplete de conta...');
    const accountSearchInput = document.getElementById('filter_contaSearch');
    if (accountSearchInput && accountSearchInput._generateAccountOptions) {
        accountSearchInput._generateAccountOptions();
        console.log('✅ Autocomplete atualizado');
    }
    
    // 2. Depois, configurar a lupa para o modal (separadamente)
    console.log('🔧 Configurando lupa de conta...');
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
    console.log('🔧 populateFilterOptions iniciada');
    
    const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
    
    dynamicFilters.forEach(select => {
        const field = select.dataset.field;
        console.log(`Atualizando filtro: ${field}`);
        
        // Get current filter values INCLUDING CONTA_SEARCH
        const currentFilters = {};
        
        // Filtros de select
        dynamicFilters.forEach(filterSelect => {
            if (filterSelect.value) {
                currentFilters[filterSelect.dataset.field] = filterSelect.value;
            }
        });
        
        // INCLUIR filtro de conta se ativo
        const contaInput = document.getElementById('filter_contaSearch');
        if (contaInput && contaInput.value) {
            console.log(`Incluindo filtro de conta: ${contaInput.value}`);
            currentFilters['CONTA_SEARCH'] = {
                value: contaInput.value,
                conta: contaInput.dataset.selectedConta,
                description: contaInput.dataset.selectedDescription
            };
        }
        
        console.log('Filtros ativos para cascata:', currentFilters);
        
        // Filter data based on ALL selections (including conta)
        let filteredDataForOptions = rawData;
        
        Object.keys(currentFilters).forEach(currentField => {
            if (currentField !== field) { // Não filtrar pelo próprio campo
                if (currentField === 'CONTA_SEARCH') {
                    // Aplicar filtro de conta
                    const searchFilter = currentFilters[currentField];
                    filteredDataForOptions = filteredDataForOptions.filter(item => {
                        const { contaSAP, description } = extractAccountInfo(item);
                        
                        // Match exato se tem conta/descrição selecionada
                        if (searchFilter.conta || searchFilter.description) {
                            return (searchFilter.conta && String(searchFilter.conta) === String(contaSAP)) || 
                                   (searchFilter.description && searchFilter.description === description);
                        }
                        
                        // Senão, buscar por texto
                        const searchText = searchFilter.value.toLowerCase();
                        return String(contaSAP).toLowerCase().includes(searchText) || 
                               String(description).toLowerCase().includes(searchText);
                    });
                } else {
                    // Filtros normais
                    filteredDataForOptions = filteredDataForOptions.filter(item => 
                        item[currentField] === currentFilters[currentField]
                    );
                }
            }
        });
        
        console.log(`Dados após filtros cascata para ${field}: ${rawData.length} → ${filteredDataForOptions.length}`);
        
        // Get unique values
        const uniqueValues = [...new Set(filteredDataForOptions.map(item => item[field]).filter(val => val))];
        
        console.log(`Valores únicos para ${field}: ${uniqueValues.length} opções`);
        
        // Store current value
        const currentValue = select.value;
        
        // Clear and populate options
        select.innerHTML = '<option value="">Todos</option>';
        
        uniqueValues.sort().forEach(value => {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        });
        
        // Restore value if still exists
        if (currentValue && uniqueValues.includes(currentValue)) {
            select.value = currentValue;
        } else if (currentValue && !uniqueValues.includes(currentValue)) {
            console.log(`⚠️ Valor "${currentValue}" não existe mais nas opções de ${field}`);
        }
    });
    
    console.log('✅ populateFilterOptions concluída');
}

function updateCascadingFilters() {
    console.log('🔄 updateCascadingFilters chamada');
    
    // Atualizar selects normais
    populateFilterOptions();
    
    // IMPORTANTE: Atualizar também as opções de conta quando outros filtros mudarem
    const accountSearchInput = document.getElementById('filter_contaSearch');
    if (accountSearchInput && accountSearchInput._generateAccountOptions) {
        console.log('🔄 Atualizando opções de conta devido a mudança em filtros...');
        accountSearchInput._generateAccountOptions();
    }
}

// FUNÇÃO ATUALIZADA: Aplicar filtros incluindo busca de conta
function applyFilters() {
    console.log('🔄 === APLICANDO FILTROS ===');
    
    const filterSelects = document.querySelectorAll('#dynamicFilters select');
    const filterInputs = document.querySelectorAll('#dynamicFilters input[type="text"]');
    const filters = {};
    
    console.log(`Encontrados ${filterSelects.length} selects e ${filterInputs.length} inputs`);
    
    // Filtros de select normais
    filterSelects.forEach(select => {
        if (select.value) {
            filters[select.dataset.field] = select.value;
            console.log(`✅ Filtro select: ${select.dataset.field} = ${select.value}`);
        }
    });
    
    // Filtro de busca de conta
    filterInputs.forEach(input => {
        console.log(`Verificando input: ${input.id}, value: "${input.value}", field: ${input.dataset.field}`);
        
        if (input.value && input.dataset.field === 'CONTA_SEARCH') {
            console.log(`✅ Filtro conta encontrado: ${input.value}`, {
                selectedConta: input.dataset.selectedConta,
                selectedDescription: input.dataset.selectedDescription
            });
            
            filters.CONTA_SEARCH = {
                value: input.value,
                conta: input.dataset.selectedConta,
                description: input.dataset.selectedDescription
            };
        }
    });
    
    console.log('📋 Todos os filtros a aplicar:', filters);
    console.log(`📊 Dados antes do filtro: ${rawData.length} registros`);
    
    // Aplicar filtros
    filteredData = rawData.filter((item, index) => {
        const passes = Object.keys(filters).every(field => {
            if (field === 'CONTA_SEARCH') {
                const searchFilter = filters[field];
                const { contaSAP, description } = extractAccountInfo(item);
                
                // Debug da comparação
                if (index < 3) { // Log apenas dos primeiros 3 itens para não poluir
                    console.log(`Item ${index} - Conta: "${contaSAP}", Descrição: "${description}"`);
                    console.log(`Filtro - Conta: "${searchFilter.conta}", Descrição: "${searchFilter.description}"`);
                }
                
                // Se tem conta/descrição selecionada, fazer match exato
                if (searchFilter.conta || searchFilter.description) {
                    const match = (searchFilter.conta && String(searchFilter.conta) === String(contaSAP)) || 
                                  (searchFilter.description && searchFilter.description === description);
                    
                    if (index < 3) {
                        console.log(`Match exato: ${match}`);
                    }
                    return match;
                }
                
                // Senão, buscar por texto
                const searchText = searchFilter.value.toLowerCase();
                const textMatch = String(contaSAP).toLowerCase().includes(searchText) || 
                                 String(description).toLowerCase().includes(searchText);
                
                if (index < 3) {
                    console.log(`Match por texto "${searchText}": ${textMatch}`);
                }
                return textMatch;
            } else {
                // Filtros normais
                const match = item[field] === filters[field];
                if (index < 3) {
                    console.log(`Filtro ${field}: item="${item[field]}" === filtro="${filters[field]}" = ${match}`);
                }
                return match;
            }
        });
        
        return passes;
    });
    
    console.log(`✅ RESULTADO: ${rawData.length} → ${filteredData.length} registros após filtros`);
    
    // Atualizar dashboard primeiro
    console.log('🔄 Atualizando dashboard...');
    updateDashboard();
    
    // DEPOIS atualizar os outros filtros para mostrar apenas opções relacionadas
    console.log('🔄 Atualizando filtros cascata (incluindo conta)...');
    setTimeout(() => {
        populateFilterOptions();
    }, 100);
}

function clearFilters() {
    document.getElementById('sheetSelector').value = '';
    
    const filterSelects = document.querySelectorAll('#dynamicFilters select');
    filterSelects.forEach(select => {
        select.value = '';
    });
    
    // Limpar campos de busca
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
        console.log('Dados insuficientes para atualizar dashboard:', {
            currentPeriod: !!currentPeriod,
            previousPeriod: !!previousPeriod,
            selectedSheet: !!selectedSheet
        });
        return;
    }
    
    // Verificar se temos dados válidos antes de continuar
    if (!filteredData || filteredData.length === 0) {
        console.log('Nenhum dado filtrado disponível');
        return;
    }
    
    // Calculate summary values
    calculateSummaryValues(currentPeriod, previousPeriod);
    
    // Update charts apenas se temos períodos válidos
    if (periods && periods.length > 0) {
        updateCharts();
    } else {
        console.log('Nenhum período disponível para gráficos');
    }
    
    // Update details table
    filterDetailTable();
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
    // Verificar se temos dados suficientes antes de tentar atualizar
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
    
    // Verificações básicas
    if (!periods || periods.length === 0) {
        console.log('Nenhum período disponível para o gráfico');
        chartElement.innerHTML = '<div class="no-data">Selecione uma aba e períodos para visualizar o gráfico</div>';
        return;
    }
    
    if (!filteredData || filteredData.length === 0) {
        console.log('Nenhum dado filtrado disponível');
        chartElement.innerHTML = '<div class="no-data">Nenhum dado disponível para o gráfico</div>';
        return;
    }

    // Verificar se temos períodos selecionados
    const currentPeriod = document.getElementById('currentPeriod')?.value;
    const previousPeriod = document.getElementById('previousPeriod')?.value;
    
    if (!currentPeriod || !previousPeriod) {
        console.log('Períodos não selecionados');
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
        
        // Verificar se temos dados válidos
        if (!seriesData || seriesData.length === 0 || seriesData.every(serie => !serie.data || serie.data.length === 0)) {
            chartElement.innerHTML = '<div class="no-data">Dados insuficientes para gerar o gráfico</div>';
            return;
        }
        
        // Destroy existing chart if any
        if (charts.evolution) {
            try {
                charts.evolution.destroy();
                charts.evolution = null;
            } catch (error) {
                console.log('Erro ao destruir gráfico anterior:', error);
            }
        }
        
        // Clear the chart container
        chartElement.innerHTML = '';
        
        // Aguardar um pouco antes de criar o novo gráfico
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
                    tooltip: {
                        shared: true,
                        intersect: false,
                        custom: function({ series, seriesIndex, dataPointIndex, w }) {
                            const currentLabel = categories[dataPointIndex];
                            const entities = ['telecom', 'vogel', 'consolidado'];
                            const entityNames = ['Telecom', 'Vogel', 'Somado'];
                            const entityColors = ['#0066cc', '#ffab00', '#00c853'];
                            
                            let tooltipContent = `
                                <div style="background: white; padding: 12px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); border: 1px solid #e5e7eb; min-width: 200px;">
                                    <div style="font-weight: 600; margin-bottom: 8px; color: #374151; font-size: 14px;">
                                        ${currentLabel}
                                    </div>
                            `;
                            
                            series.forEach((serieData, index) => {
                                const value = serieData[dataPointIndex];
                                const entityName = entityNames[index];
                                const entityKey = entities[index];
                                const color = entityColors[index];
                                
                                const yoyData = calculateYoYVariation(currentLabel, entityKey, isQuarterly);
                                
                                tooltipContent += `
                                    <div style="display: flex; align-items: center; margin: 4px 0;">
                                        <span style="width: 12px; height: 12px; background-color: ${color}; border-radius: 50%; margin-right: 8px;"></span>
                                        <span style="color: #374151; font-size: 13px; flex: 1;">
                                            <strong>${entityName}:</strong> ${formatCurrency(value)}
                                        </span>
                                    </div>
                                `;
                                
                                if (yoyData.hasComparison) {
                                    const variationColor = yoyData.variation >= 0 ? '#10b981' : '#ef4444';
                                    const variationIcon = yoyData.variation >= 0 ? '↗' : '↘';
                                    
                                    tooltipContent += `
                                        <div style="margin-left: 20px; font-size: 11px; color: #6b7280; margin-top: 2px;">
                                            vs ${yoyData.previousLabel}: 
                                            <span style="color: ${variationColor}; font-weight: 600;">
                                                ${variationIcon} ${Math.abs(yoyData.variation).toFixed(1)}%
                                            </span>
                                            <br>
                                            <span style="font-size: 10px;">
                                                (${formatCurrencyCompact(yoyData.previousValue)})
                                            </span>
                                        </div>
                                    `;
                                }
                            });
                            
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

function calculateYoYVariation(currentLabel, entity, isQuarterly = false) {
    try {
        if (isQuarterly) {
            const [quarter, year] = currentLabel.split(' ');
            const previousYear = (parseInt(year) - 1).toString();
            const previousLabel = `${quarter} ${previousYear}`;
            
            const quarterlyData = calculateQuarterlyData();
            const currentValue = quarterlyData[currentLabel] ? quarterlyData[currentLabel][entity] : 0;
            const previousValue = quarterlyData[previousLabel] ? quarterlyData[previousLabel][entity] : null;
            
            if (previousValue !== null && previousValue !== 0) {
                const variation = ((currentValue - previousValue) / previousValue) * 100;
                return {
                    hasComparison: true,
                    variation: variation,
                    previousLabel: previousLabel,
                    previousValue: previousValue,
                    currentValue: currentValue
                };
            }
        } else {
            const [month, year] = currentLabel.split('/');
            const previousYear = (parseInt(year) - 1).toString();
            const previousLabel = `${month}/${previousYear}`;
            
            const currentPeriod = periods.find(p => formatPeriod(p) === currentLabel);
            const previousPeriod = periods.find(p => formatPeriod(p) === previousLabel);
            
            if (currentPeriod && previousPeriod) {
                const currentValue = calculateTotal(entity, currentPeriod);
                const previousValue = calculateTotal(entity, previousPeriod);
                
                if (previousValue !== 0) {
                    const variation = ((currentValue - previousValue) / previousValue) * 100;
                    return {
                        hasComparison: true,
                        variation: variation,
                        previousLabel: previousLabel,
                        previousValue: previousValue,
                        currentValue: currentValue
                    };
                }
            }
        }
        
        return { hasComparison: false };
    } catch (error) {
        console.error('Error calculating YoY variation:', error);
        return { hasComparison: false };
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
    
    const ws = XLSX.utils.aoa_to_sheet([
        ['Conta SAP', 'Descrição', 'Valor Anterior', 'Valor Atual', 'Variação', 'Variação %']
    ]);
    
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
    
    const numberFormat = '0.00';
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
    
    const entityName = entityFilter === 'total' ? 'consolidado-total' : entityFilter;
    const filename = `detalhamento-contas-${entityName}-${new Date().toISOString().split('T')[0]}.xlsx`;
    
    XLSX.writeFile(wb, filename);
}

function calculateTotal(entity, period) {
    return filteredData.reduce((total, item) => {
        return total + (item.values && item.values[period] ? item.values[period][entity] || 0 : 0);
    }, 0);
}