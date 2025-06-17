let workbook = null;
let allSheetsData = {};
let selectedSheet = '';
let rawData = [];
let filteredData = [];
let periods = [];
let charts = {};
let chartViewMode = 'monthly'; // 'monthly' or 'quarterly'

// Sheet filter configurations
const SHEET_FILTERS = {
    'SINTETICA': [
        { id: 'descricaoConta', label: 'Descrição Conta', field: 'Descrição Conta' },
        { id: 'descricaoDf', label: 'Descrição DF', field: 'Descrição DF' },
        { id: 'descricaoDfGroup', label: 'Descrição DF Group', field: 'Group' }
    ],
    'ANALITICA': [
        { id: 'descricaoConta', label: 'Descrição Conta', field: 'Descrição Conta' },
        { id: 'fsGroupReport', label: 'FS Group REPORT', field: 'FS Group REPORT' },
        { id: 'fsGroup', label: 'FS Group SAP', field: 'FS Group SAP' },
        { id: 'sintetica', label: 'Sintética', field: 'SINTETICA' }
    ],
    'ABERTURA DF': [
        { id: 'descricaoConta', label: 'Descrição Conta', field: 'Descrição Conta' },
        { id: 'descricaoDf', label: 'Descrição DF', field: 'Descrição DF' }
    ],
    'SUBGRUPO': [
        { id: 'descricaoConta', label: 'Descrição Conta', field: 'Descrição Conta' },
        { id: 'descricaoDf', label: 'Descrição DF', field: 'Descrição DF' }
    ],
    'GRUPO FEC': [
        { id: 'descricaoConta', label: 'Descrição Conta', field: 'Descrição Conta' },
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
        
        // Sum values for all periods in this quarter (mesmo que incompleto)
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

// Toggle chart view mode
function toggleChartView(mode) {
    chartViewMode = mode;
    
    // Update button states
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
    
    // Update chart
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
    // File input
    document.getElementById('fileInput').addEventListener('change', handleFileUpload);
    
    // Filter controls
    document.getElementById('applyFilters').addEventListener('click', () => applyFilters());
    document.getElementById('clearFilters').addEventListener('click', () => clearFilters());
    
    // Period selection
    document.getElementById('currentPeriod').addEventListener('change', () => updateDashboard());
    document.getElementById('previousPeriod').addEventListener('change', () => updateDashboard());
    
    // Sheet selector
    document.getElementById('sheetSelector').addEventListener('change', handleSheetChange);
    
    // Chart view toggle
    document.getElementById('monthlyBtn').addEventListener('click', () => toggleChartView('monthly'));
    document.getElementById('quarterlyBtn').addEventListener('click', () => toggleChartView('quarterly'));
    
    // Modal controls
    document.getElementById('uploadButton').addEventListener('click', () => showUploadModal(true));
    document.getElementById('closeModal').addEventListener('click', () => showUploadModal(false));
    
    // Table filter and sort
    document.getElementById('descriptionFilter').addEventListener('input', filterDetailTable);
    document.getElementById('sortFilter').addEventListener('change', filterDetailTable);
    document.getElementById('entityFilter').addEventListener('change', filterDetailTable);
    
    // Export and download controls
    document.getElementById('exportTable').addEventListener('click', exportTableToExcel);
    document.getElementById('downloadChart').addEventListener('click', downloadChart);
    
    // Add click outside to close modal
    window.addEventListener('click', function(e) {
        const modal = document.getElementById('uploadModal');
        if (e.target === modal) {
            showUploadModal(false);
        }
    });
    
    // Handle window resize for responsive adjustments
    window.addEventListener('resize', handleResize);
    handleResize(); // Initial call
}

// Update current date display
function updateCurrentDate() {
    const dateDisplay = document.getElementById('dateDisplay');
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    
    dateDisplay.textContent = `Data de análise: ${day}/${month}/${year}`;
}

// Handle responsive resize
function handleResize() {
    const sidebar = document.querySelector('.sidebar');
    const sidebarContent = document.querySelector('.sidebar-content');
    const sidebarHeader = document.querySelector('.sidebar-header');
    
    if (window.innerWidth <= 768) {
        // Mobile layout - verify if toggle button exists
        if (!document.querySelector('.toggle-sidebar')) {
            const toggleButton = document.createElement('button');
            toggleButton.className = 'toggle-sidebar';
            toggleButton.innerHTML = '<i class="fas fa-bars"></i>';
            toggleButton.addEventListener('click', toggleSidebar);
            sidebarHeader.appendChild(toggleButton);
        }
    } else {
        // Desktop layout - remove toggle button if exists
        const toggleButton = document.querySelector('.toggle-sidebar');
        if (toggleButton) {
            toggleButton.remove();
        }
        
        // Reset sidebar state for desktop
        sidebarContent.classList.remove('expanded');
        sidebarContent.style.maxHeight = '';
    }
    
    // Redraw chart if exists to fit new size
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
        
        // Auto-selecionar a primeira aba após o upload
        const sheetSelector = document.getElementById('sheetSelector');
        if (sheetSelector.options.length > 1) {
            sheetSelector.selectedIndex = 1; // Primeira aba real (índice 0 é "Selecione...")
            handleSheetChange(); // Disparar evento para carregar a aba selecionada
        }
        
        showUploadModal(false);
        showLoading(false);
    } catch (error) {
        console.error('Error processing file:', error);
        alert(`Erro ao processar arquivo: ${error.message}\n\nVerifique se é um arquivo Excel válido.`);
        showLoading(false);
    }
}

// Read Excel File and all sheets
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                
                // Get all sheet names
                const sheetNames = workbook.SheetNames;
                console.log('Available sheets:', sheetNames);
                
                // Populate sheet selector (excluding non-existent sheets)
                const sheetSelector = document.getElementById('sheetSelector');
                sheetSelector.innerHTML = '<option value="">Selecione uma aba...</option>';
                
                sheetNames.forEach(sheetName => {
                    // Skip if the sheet name is "3000 ANT" or similar non-standard sheets
                    if (sheetName.toLowerCase().includes('3000 ant')) {
                        return;
                    }
                    
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelector.appendChild(option);
                });
                
                // Read all sheets data
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

// Handle sheet change
function handleSheetChange() {
    const sheetSelector = document.getElementById('sheetSelector');
    const sheetName = sheetSelector.value;
    
    if (!sheetName) {
        document.getElementById('dynamicFilters').innerHTML = '';
        
        // Reset period selectors
        document.getElementById('currentPeriod').innerHTML = '<option value="">Selecione...</option>';
        document.getElementById('previousPeriod').innerHTML = '<option value="">Selecione...</option>';
        
        // Reset dashboard
        rawData = [];
        filteredData = [];
        periods = [];
        selectedSheet = '';
        
        return;
    }
    
    selectedSheet = sheetName;
    
    // Create dynamic filters for this sheet
    createDynamicFilters(sheetName);
    
    // Process sheet data
    processSheetData(sheetName);
}

// Create dynamic filters based on sheet configuration
function createDynamicFilters(sheetName) {
    const dynamicFiltersContainer = document.getElementById('dynamicFilters');
    dynamicFiltersContainer.innerHTML = '';
    
    // Get filter configuration for this sheet
    const filterConfig = SHEET_FILTERS[sheetName] || [
        { id: 'descricaoConta', label: 'Descrição Conta', field: 'Descrição Conta' },
        { id: 'fsGroupReport', label: 'FS Group REPORT', field: 'FS Group REPORT' },
        { id: 'fsGroup', label: 'FS Group', field: 'FS Group' },
        { id: 'sintetica', label: 'SINTÉTICA', field: 'SINTETICA' }
    ];
    
    // Create filter groups
    filterConfig.forEach(config => {
        const filterGroup = document.createElement('div');
        filterGroup.className = 'filter-group';
        filterGroup.innerHTML = `
            <label for="filter_${config.id}">${config.label}</label>
            <div class="select-wrapper">
                <select id="filter_${config.id}" data-field="${config.field}">
                    <option value="">Todos</option>
                </select>
                <i class="fas fa-chevron-down"></i>
            </div>
        `;
        dynamicFiltersContainer.appendChild(filterGroup);
    });
    
    // Add cascade event listeners
    filterConfig.forEach(config => {
        const select = document.getElementById(`filter_${config.id}`);
        if (select) {
            select.addEventListener('change', () => updateCascadingFilters());
        }
    });
}

// Process sheet data
function processSheetData(sheetName) {
    const data = allSheetsData[sheetName];
    
    if (!data || data.length < 2) {
        alert('Dados insuficientes na aba selecionada');
        return;
    }
    
    const headers = data[0];
    const rows = data.slice(1);
    
    // Find date columns (format: DD/MM/YYYY)
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
                // Normalizar "Somado" para "consolidado" para manter compatibilidade
                const normalizedEntity = entity.toLowerCase() === 'somado' ? 'consolidado' : entity.toLowerCase();
                dateColumns[date][normalizedEntity] = index;
            }
        }
    });
    
    // Extract periods and sort them
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
    
    // Process rows
    rawData = rows.map(row => {
        const item = {
            values: {}
        };
        
        // Add all fields from headers
        headers.forEach((header, index) => {
            if (typeof header === 'string' && !header.match(/\d{2}\/\d{2}\/\d{4}/)) {
                item[header] = row[index];
            }
        });
        
        // Extract values for each period and entity
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
    
    // Initialize filters for this sheet
    populateFilterOptions();
    
    // Update period selectors
    updatePeriodSelectors();
    
    // Select last two periods automatically
    autoSelectLastTwoPeriods();
    
    // Update dashboard with filtered data
    filteredData = [...rawData];
    updateDashboard();
}

// Function to automatically select last two periods
function autoSelectLastTwoPeriods() {
    if (periods.length >= 2) {
        // Get last two periods
        const lastTwoPeriods = periods.slice(-2);
        
        const previousPeriodSelect = document.getElementById('previousPeriod');
        const currentPeriodSelect = document.getElementById('currentPeriod');
        
        // Set previous to penultimate and current to last
        if (previousPeriodSelect && lastTwoPeriods[0]) {
            previousPeriodSelect.value = lastTwoPeriods[0];
        }
        
        if (currentPeriodSelect && lastTwoPeriods[1]) {
            currentPeriodSelect.value = lastTwoPeriods[1];
        }
    }
}

// Update period selectors
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

// Populate filter options with cascade support
function populateFilterOptions() {
    const dynamicFilters = document.querySelectorAll('#dynamicFilters select');
    
    dynamicFilters.forEach(select => {
        const field = select.dataset.field;
        
        // Get current filter values
        const currentFilters = {};
        dynamicFilters.forEach(filterSelect => {
            if (filterSelect.value) {
                currentFilters[filterSelect.dataset.field] = filterSelect.value;
            }
        });
        
        // Filter data based on other selections
        let filteredDataForOptions = rawData;
        
        Object.keys(currentFilters).forEach(currentField => {
            if (currentField !== field) {
                filteredDataForOptions = filteredDataForOptions.filter(item => 
                    item[currentField] === currentFilters[currentField]
                );
            }
        });
        
        // Get unique values
        const uniqueValues = [...new Set(filteredDataForOptions.map(item => item[field]).filter(val => val))];
        
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
        }
    });
}

// Update cascading filters
function updateCascadingFilters() {
    populateFilterOptions();
}

// Apply Filters
function applyFilters() {
    const filterSelects = document.querySelectorAll('#dynamicFilters select');
    const filters = {};
    
    filterSelects.forEach(select => {
        if (select.value) {
            filters[select.dataset.field] = select.value;
        }
    });
    
    filteredData = rawData.filter(item => {
        return Object.keys(filters).every(field => {
            return item[field] === filters[field];
        });
    });
    
    updateDashboard();
}

// Clear Filters
function clearFilters() {
    // Reset sheet selector
    document.getElementById('sheetSelector').value = '';
    
    // Clear all filters
    const filterSelects = document.querySelectorAll('#dynamicFilters select');
    filterSelects.forEach(select => {
        select.value = '';
    });
    
    // Reset dynamic filters container
    document.getElementById('dynamicFilters').innerHTML = '';
    
    // Reset period selectors
    document.getElementById('currentPeriod').innerHTML = '<option value="">Selecione...</option>';
    document.getElementById('previousPeriod').innerHTML = '<option value="">Selecione...</option>';
    
    // Reset chart view to monthly
    chartViewMode = 'monthly';
    const monthlyBtn = document.getElementById('monthlyBtn');
    const quarterlyBtn = document.getElementById('quarterlyBtn');
    const chartTitle = document.getElementById('chartTitle');
    monthlyBtn.classList.add('active');
    quarterlyBtn.classList.remove('active');
    chartTitle.textContent = 'Evolução Mensal';
    
    // Reset all data
    rawData = [];
    filteredData = [];
    periods = [];
    selectedSheet = '';
    
    // Reset cards and chart
    updateSummaryCard('telecom', 0, 0, 0);
    updateSummaryCard('vogel', 0, 0, 0);
    updateSummaryCard('consolidado', 0, 0, 0);
    
    // Clear table
    const tbody = document.querySelector('#detailsTable tbody');
    if (tbody) {
        tbody.innerHTML = '';
    }
    
    // Clear chart
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

// Calculate Summary Values
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

// Update Dashboard
function updateDashboard() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) return;
    
    // Calculate summary values
    calculateSummaryValues(currentPeriod, previousPeriod);
    
    // Update charts
    updateCharts();
    
    // Update details table
    filterDetailTable();
}

// Update Summary Card
function updateSummaryCard(entity, currentValue, previousValue, variation) {
    const variationElement = document.getElementById(`${entity}Variation`);
    const currentElement = document.getElementById(`${entity}Current`);
    const previousElement = document.getElementById(`${entity}Previous`);
    
    if (!variationElement || !currentElement || !previousElement) {
        console.error(`Elementos para ${entity} não encontrados`);
        return;
    }
    
    // Update values
    currentElement.textContent = formatCurrency(currentValue);
    previousElement.textContent = formatCurrency(previousValue);
    
    // Update variation
    const variationValueElement = variationElement.querySelector('.variation-value');
    if (variationValueElement) {
        variationValueElement.textContent = `${Math.abs(variation).toFixed(2)}%`;
    }
    
    // Update styles
    variationElement.className = 'variation-indicator';
    if (variation > 0) {
        variationElement.classList.add('positive');
    } else if (variation < 0) {
        variationElement.classList.add('negative');
    }
}

// Update Charts
function updateCharts() {
    updateEvolutionChart();
}

// Evolution Chart Function
function updateEvolutionChart() {
    const chartElement = document.getElementById('evolutionChart');
    if (!chartElement) {
        console.error('Elemento do gráfico não encontrado');
        return;
    }
    
    // Verificar se temos dados para mostrar
    if (!periods || periods.length === 0) {
        console.log('Nenhum período disponível para o gráfico');
        return;
    }
    
    let seriesData, categories;
    const isQuarterly = chartViewMode === 'quarterly';
    
    if (isQuarterly) {
        // Prepare quarterly data
        const quarterlyData = calculateQuarterlyData();
        const quarters = sortQuarters(Object.keys(quarterlyData));
        
        seriesData = [
            {
                name: 'Telecom',
                data: quarters.map(quarter => ({
                    x: quarter,
                    y: quarterlyData[quarter].telecom
                }))
            },
            {
                name: 'Vogel',
                data: quarters.map(quarter => ({
                    x: quarter,
                    y: quarterlyData[quarter].vogel
                }))
            },
            {
                name: 'Consolidado',
                data: quarters.map(quarter => ({
                    x: quarter,
                    y: quarterlyData[quarter].consolidado
                }))
            }
        ];
        categories = quarters;
    } else {
        // Prepare monthly data
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
                name: 'Consolidado',
                data: periods.map(period => ({
                    x: formatPeriod(period),
                    y: calculateTotal('consolidado', period)
                }))
            }
        ];
        categories = periods.map(formatPeriod);
    }
    
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
                const entityNames = ['Telecom', 'Vogel', 'Consolidado'];
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
                    
                    // Calcular variação YoY
                    const yoyData = calculateYoYVariation(currentLabel, entityKey, isQuarterly);
                    
                    tooltipContent += `
                        <div style="display: flex; align-items: center; margin: 4px 0;">
                            <span style="width: 12px; height: 12px; background-color: ${color}; border-radius: 50%; margin-right: 8px;"></span>
                            <span style="color: #374151; font-size: 13px; flex: 1;">
                                <strong>${entityName}:</strong> ${formatCurrency(value)}
                            </span>
                        </div>
                    `;
                    
                    // Adicionar comparação YoY se disponível
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
    
    // Destroy existing chart if any
    if (charts.evolution) {
        try {
            charts.evolution.destroy();
        } catch (error) {
            console.log('Erro ao destruir gráfico anterior:', error);
        }
    }
    
    // Clear the chart container
    chartElement.innerHTML = '';
    
    // Create new chart
    try {
        charts.evolution = new ApexCharts(chartElement, options);
        charts.evolution.render();
    } catch (error) {
        console.error('Erro ao criar gráfico:', error);
        chartElement.innerHTML = '<div class="no-data">Erro ao carregar gráfico</div>';
    }
}

function calculateYoYVariation(currentLabel, entity, isQuarterly = false) {
    try {
        if (isQuarterly) {
            // Para trimestres: "1T 2025" → "1T 2024"
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
            // Para meses: "Jan/2025" → "Jan/2024"
            const [month, year] = currentLabel.split('/');
            const previousYear = (parseInt(year) - 1).toString();
            const previousLabel = `${month}/${previousYear}`;
            
            // Encontrar o período correspondente ao ano anterior
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
// Filter and sort details table
function filterDetailTable() {
    const descriptionFilter = document.getElementById('descriptionFilter').value.toLowerCase();
    const sortFilter = document.getElementById('sortFilter').value;
    const entityFilter = document.getElementById('entityFilter').value;
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) return;
    
    // Filter data by description or account
    let filtered = [...filteredData];
    if (descriptionFilter) {
        filtered = filtered.filter(item => {
            // Buscar nos mesmos campos que usamos na tabela
            let contaSAP = '-';
            let description = '-';
            
            for (const key of Object.keys(item)) {
                const keyTrimmed = key.trim().replace(/\s+/g, ' ');
                const keyNoSpaces = key.replace(/\s/g, '');
                
                if (keyTrimmed === 'Conta SAP' || keyTrimmed === 'CONTA SAP' || 
                    keyTrimmed === 'Número Conta' || keyTrimmed === 'NÚMERO CONTA' ||
                    keyNoSpaces.toLowerCase() === 'contasap' || 
                    keyNoSpaces.toLowerCase() === 'numeroconta') {
                    contaSAP = item[key] || '-';
                }
                
                if (keyTrimmed === 'Descrição Conta' || keyTrimmed === 'Descrição DF' ||
                    keyTrimmed === 'DESCRIÇÃO CONTA' || keyTrimmed === 'DESCRIÇÃO DF') {
                    description = item[key] || '-';
                }
            }
            
            return description.toLowerCase().includes(descriptionFilter) || 
                   contaSAP.toString().toLowerCase().includes(descriptionFilter);
        });
    }
    
    // Sort data based on selected option
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
    
    // Update the table
    updateDetailTableContent(filtered, currentPeriod, previousPeriod, entityFilter);
}

// Helper function to get current entity value
function getCurrentEntityValue(item, currentPeriod, entityFilter) {
    if (!item.values || !item.values[currentPeriod]) return 0;
    
    if (entityFilter === 'total') {
        return Object.values(item.values[currentPeriod]).reduce((a, b) => a + b, 0);
    } else {
        return item.values[currentPeriod][entityFilter] || 0;
    }
}

// Helper function to get entity variation
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

// Update Detail Table Content
function updateDetailTableContent(data, currentPeriod, previousPeriod, entityFilter) {
    const tbody = document.querySelector('#detailsTable tbody');
    if (!tbody) {
        console.error('Table body element not found');
        return;
    }
    
    tbody.innerHTML = '';
    
    // Check if we have data to display
    if (!data || data.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="6" class="no-data">Não há dados para exibir. Tente ajustar os filtros.</td>`;
        tbody.appendChild(row);
        return;
    }
    
    // Check if periods exist
    if (!currentPeriod || !previousPeriod) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="6" class="no-data">Selecione os períodos para visualizar os dados.</td>`;
        tbody.appendChild(row);
        return;
    }
    
    // Display top 2000 items
    data.slice(0, 2000).forEach(item => {
        try {
            const row = document.createElement('tr');
            
            // Safely get values based on entity filter
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
            
            // Get the description and account fields - buscar Conta SAP ou Número Conta
            let contaSAP = '-';
            let description = '-';
            
            // Buscar pelos campos corretos (removendo espaços e verificando variações)
            for (const key of Object.keys(item)) {
                const keyTrimmed = key.trim().replace(/\s+/g, ' '); // normalizar espaços
                const keyNoSpaces = key.replace(/\s/g, ''); // sem espaços
                
                // Verificar se é conta SAP ou número conta
                if (keyTrimmed === 'Conta SAP' || keyTrimmed === 'CONTA SAP' || 
                    keyTrimmed === 'Número Conta' || keyTrimmed === 'NÚMERO CONTA' ||
                    keyNoSpaces.toLowerCase() === 'contasap' || 
                    keyNoSpaces.toLowerCase() === 'numeroconta') {
                    contaSAP = item[key] || '-';
                }
                
                // Verificar descrição
                if (keyTrimmed === 'Descrição Conta' || keyTrimmed === 'Descrição DF' ||
                    keyTrimmed === 'DESCRIÇÃO CONTA' || keyTrimmed === 'DESCRIÇÃO DF') {
                    description = item[key] || '-';
                }
            }
            
            row.innerHTML = `
                <td>${contaSAP}</td>
                <td>${description}</td>
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

// Download Chart as image
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

// Export table to Excel
function exportTableToExcel() {
    const currentPeriod = document.getElementById('currentPeriod').value;
    const previousPeriod = document.getElementById('previousPeriod').value;
    const entityFilter = document.getElementById('entityFilter').value;
    
    if (!currentPeriod || !previousPeriod || !selectedSheet) return;
    
    // Create a worksheet from filteredData
    const ws = XLSX.utils.aoa_to_sheet([
        ['Conta SAP', 'Descrição', 'Valor Anterior', 'Valor Atual', 'Variação', 'Variação %']
    ]);
    
    // Add data rows
    const rows = filteredData.map(item => {
        // Safely get values based on entity filter
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
        
        // Buscar pelos mesmos campos que usamos na tabela
        let contaSAP = '-';
        let description = '-';
        
        for (const key of Object.keys(item)) {
            const keyTrimmed = key.trim().replace(/\s+/g, ' ');
            const keyNoSpaces = key.replace(/\s/g, '');
            
            if (keyTrimmed === 'Conta SAP' || keyTrimmed === 'CONTA SAP' || 
                keyTrimmed === 'Número Conta' || keyTrimmed === 'NÚMERO CONTA' ||
                keyNoSpaces.toLowerCase() === 'contasap' || 
                keyNoSpaces.toLowerCase() === 'numeroconta') {
                contaSAP = item[key] || '-';
            }
            
            if (keyTrimmed === 'Descrição Conta' || keyTrimmed === 'Descrição DF' ||
                keyTrimmed === 'DESCRIÇÃO CONTA' || keyTrimmed === 'DESCRIÇÃO DF') {
                description = item[key] || '-';
            }
        }
        
        return [
            contaSAP,
            description,
            previousTotal,
            currentTotal,
            variation,
            variationPercent / 100  // Excel espera valores percentuais como decimais
        ];
    });
    
    // Append rows to worksheet
    XLSX.utils.sheet_add_aoa(ws, rows, { origin: 'A2' });
    
    // Format the columns
    const numberFormat = '0.00';
    const currencyFormat = '"R$ "#,##0.00';
    const percentFormat = '0.00%';
    
    // Format columns C, D, E as currency and F as percentage
    ['C', 'D', 'E'].forEach(col => {
        for (let i = 2; i <= rows.length + 1; i++) {
            if (!ws[col + i]) continue;
            ws[col + i].z = currencyFormat;
        }
    });
    
    // Format column F as percentage
    for (let i = 2; i <= rows.length + 1; i++) {
        if (!ws['F' + i]) continue;
        ws['F' + i].z = percentFormat;
    }
    
    // Create a workbook and add the worksheet
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Detalhamento');
    
    // Generate filename with entity filter info
    const entityName = entityFilter === 'total' ? 'consolidado-total' : entityFilter;
    const filename = `detalhamento-contas-${entityName}-${new Date().toISOString().split('T')[0]}.xlsx`;
    
    // Write and download
    XLSX.writeFile(wb, filename);
}

// Utility Functions
function calculateTotal(entity, period) {
    return filteredData.reduce((total, item) => {
        return total + (item.values && item.values[period] ? item.values[period][entity] || 0 : 0);
    }, 0);
}