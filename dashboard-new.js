// Global variables
let rawData = [];
let processedData = [];
let charts = {};
let columnConfig = {};

const DEFAULT_COLUMNS = {
    'From Station': { label: 'From Station', description: 'Origin station name', enabled: true },
    'From Meso Zone': { label: 'From Meso Zone', description: 'Origin zone/region', enabled: true },
    'To Meso Zone': { label: 'To Meso Zone', description: 'Destination zone/region', enabled: true },
    'Trips': { label: 'Trips', description: 'Number of trips (required)', enabled: true, required: true }
};

// DOM Elements
const preUploadConfigPanel = document.getElementById('preUploadConfigPanel');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const columnConfigPanel = document.getElementById('columnConfigPanel');
const columnConfigList = document.getElementById('columnConfigList');
const proceedButton = document.getElementById('proceedButton');
const dataPreviewPanel = document.getElementById('dataPreviewPanel');
const previewTable = document.getElementById('previewTable');
const dashboardSection = document.getElementById('dashboardSection');
const loadingSpinner = document.getElementById('loadingSpinner');
const rowCount = document.getElementById('rowCount');
const colCount = document.getElementById('colCount');
const fileStatus = document.getElementById('fileStatus');
const resetButton = document.getElementById('resetButton');

// Pre-upload configuration elements
const defaultColumnsContainer = document.getElementById('defaultColumnsContainer');
const customColumnsConfigContainer = document.getElementById('customColumnsConfigContainer');
const addConfigColumnBtn = document.getElementById('addConfigColumnBtn');
const resetConfigBtn = document.getElementById('resetConfigBtn');
const confirmConfigBtn = document.getElementById('confirmConfigBtn');

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    initializeColumnConfig();
    setupEventListeners();
});

// Initialize column configuration UI
function initializeColumnConfig() {
    columnConfig = JSON.parse(JSON.stringify(DEFAULT_COLUMNS));
    renderDefaultColumnsUI();
    renderCustomColumnsUI();
}

// Render default columns in pre-upload UI
function renderDefaultColumnsUI() {
    defaultColumnsContainer.innerHTML = '';
    
    Object.entries(DEFAULT_COLUMNS).forEach(([key, col]) => {
        const isEnabled = columnConfig[key]?.enabled ?? col.enabled;
        const item = document.createElement('div');
        item.style.cssText = 'display: flex; gap: 10px; align-items: flex-start; padding: 10px; background: white; border-radius: 6px; border: 1px solid var(--border-color);';
        
        item.innerHTML = `
            <input type="checkbox" id="config-${key}" ${isEnabled ? 'checked' : ''} style="margin-top: 4px;" ${col.required ? 'disabled' : ''}>
            <div style="flex: 1;">
                <label for="config-${key}" style="font-weight: 600; color: var(--primary-color); cursor: pointer; margin-bottom: 4px; display: block;">
                    ${col.label}${col.required ? ' <span style="color: red;">*</span>' : ''}
                </label>
                <p style="font-size: 12px; color: #999; margin: 0;">${col.description}</p>
            </div>
        `;
        
        const checkbox = item.querySelector('input');
        checkbox.addEventListener('change', (e) => {
            columnConfig[key].enabled = e.target.checked;
        });
        
        defaultColumnsContainer.appendChild(item);
    });
}

// Render custom columns in pre-upload UI
function renderCustomColumnsUI() {
    customColumnsConfigContainer.innerHTML = '';
    
    const customCols = Object.entries(columnConfig).filter(([key]) => !DEFAULT_COLUMNS[key]);
    
    customCols.forEach(([key, col]) => {
        const item = document.createElement('div');
        item.style.cssText = 'display: flex; gap: 10px; align-items: center; padding: 10px; background: white; border-radius: 6px; border: 1px solid var(--border-color);';
        
        item.innerHTML = `
            <input type="text" value="${col.label}" placeholder="Column name" style="flex: 1; padding: 6px; border: 1px solid var(--border-color); border-radius: 4px; font-size: 13px;">
            <button class="btn btn-danger btn-sm" style="padding: 4px 10px; font-size: 12px;">Remove</button>
        `;
        
        const input = item.querySelector('input');
        const removeBtn = item.querySelector('button');
        
        input.addEventListener('change', (e) => {
            col.label = e.target.value;
        });
        
        removeBtn.addEventListener('click', () => {
            delete columnConfig[key];
            renderCustomColumnsUI();
        });
        
        customColumnsConfigContainer.appendChild(item);
    });
}

// Add custom column
function addCustomColumn() {
    const id = 'custom_' + Date.now();
    columnConfig[id] = { label: 'Custom Column', description: 'Additional analysis column', enabled: true };
    renderCustomColumnsUI();
}

// Reset configuration to defaults
function resetConfiguration() {
    columnConfig = JSON.parse(JSON.stringify(DEFAULT_COLUMNS));
    renderDefaultColumnsUI();
    renderCustomColumnsUI();
}

// Setup event listeners
function setupEventListeners() {
    fileInput.addEventListener('change', handleFileUpload);
    proceedButton.addEventListener('click', proceedToDashboard);
    resetButton.addEventListener('click', resetDashboard);
    addConfigColumnBtn.addEventListener('click', addCustomColumn);
    resetConfigBtn.addEventListener('click', resetConfiguration);
    confirmConfigBtn.addEventListener('click', () => {
        alert('Column configuration saved! Now upload your file.');
    });
}

// Parse file (CSV or Excel)
async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    loadingSpinner.style.display = 'flex';

    try {
        if (file.name.endsWith('.csv')) {
            parseCSV(file);
        } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            parseExcel(file);
        } else {
            alert('Unsupported file format. Please use CSV or Excel files.');
            loadingSpinner.style.display = 'none';
            return;
        }
    } catch (error) {
        console.error('Error parsing file:', error);
        alert('Error reading file: ' + error.message);
        loadingSpinner.style.display = 'none';
    }
}

// Parse CSV file
function parseCSV(file) {
    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
            rawData = results.data.filter(row => Object.values(row).some(v => v));
            processFileData(Object.keys(results.data[0] || {}));
        },
        error: (error) => {
            alert('Error parsing CSV: ' + error.message);
            loadingSpinner.style.display = 'none';
        }
    });
}

// Parse Excel file
function parseExcel(file) {
    if (typeof XLSX === 'undefined') {
        alert('Excel library not loaded. Please refresh the page and try again.');
        loadingSpinner.style.display = 'none';
        return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Get raw data with proper header detection
            const rawSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (rawSheet.length === 0) {
                alert('Sheet is empty');
                loadingSpinner.style.display = 'none';
                return;
            }
            
            // Use first row as headers and trim whitespace
            const headers = rawSheet[0].map(h => String(h).trim());
            const dataRows = rawSheet.slice(1);
            
            // Convert to array of objects with headers as keys
            rawData = dataRows.map(row => {
                const obj = {};
                headers.forEach((header, index) => {
                    const value = row[index];
                    obj[header] = value !== undefined ? value : '';
                });
                return obj;
            }).filter(row => Object.values(row).some(v => v));
            
            console.log('Detected headers:', headers);
            console.log('Total rows:', rawData.length);
            console.log('First data row:', rawData[0]);
            
            processFileData(headers);
        } catch (error) {
            alert('Error parsing Excel: ' + error.message);
            console.error(error);
            loadingSpinner.style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

// Process file data and show column mapping UI
function processFileData(availableColumns) {
    if (rawData.length === 0) {
        alert('File is empty');
        loadingSpinner.style.display = 'none';
        return;
    }

    // Update stats
    rowCount.textContent = rawData.length;
    colCount.textContent = availableColumns.length;
    fileStatus.textContent = '✓ Ready';
    fileInfo.classList.remove('hidden');

    // Show preview
    showDataPreview(availableColumns.slice(0, 5));

    // Show column mapping panel (to match file columns with config)
    showColumnMapping(availableColumns);
    columnConfigPanel.classList.remove('hidden');
    dataPreviewPanel.classList.remove('hidden');
    loadingSpinner.style.display = 'none';
}

// Show column mapping (match user config to actual file columns)
function showColumnMapping(availableColumns) {
    columnConfigList.innerHTML = '';

    // Get enabled columns from config
    const enabledConfig = Object.entries(columnConfig).filter(([_, col]) => col.enabled);

    enabledConfig.forEach(([configKey, configCol]) => {
        const item = document.createElement('div');
        item.className = 'column-config-item';
        
        item.innerHTML = `
            <label class="form-label mb-0" style="min-width: 180px;">
                <strong>${configCol.label}</strong>
            </label>
            <select class="form-select form-select-sm column-input" data-config-key="${configKey}">
                <option value="">-- Select from file --</option>
                ${availableColumns.map(col => `
                    <option value="${col}">${col}</option>
                `).join('')}
            </select>
        `;
        
        columnConfigList.appendChild(item);
    });

    // Add event listeners
    document.querySelectorAll('.column-input').forEach(select => {
        select.addEventListener('change', (e) => {
            const configKey = e.target.dataset.configKey;
            // Store the mapping
            if (!columnConfig[configKey]) {
                columnConfig[configKey] = { label: '', enabled: true };
            }
            columnConfig[configKey].fileColumn = e.target.value;
        });
    });
}

// Show data preview
function showDataPreview(columns) {
    const thead = previewTable.querySelector('thead');
    const tbody = previewTable.querySelector('tbody');
    
    thead.innerHTML = `<tr>${columns.map(col => `<th>${col}</th>`).join('')}</tr>`;
    tbody.innerHTML = '';
    
    rawData.slice(0, 10).forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = columns.map(col => `<td>${row[col] || '-'}</td>`).join('');
        tbody.appendChild(tr);
    });
}

// Proceed to dashboard
function proceedToDashboard() {
    const enabledConfig = Object.entries(columnConfig).filter(([_, col]) => col.enabled);
    
    // Validate that Trips column is mapped
    const tripsConfig = enabledConfig.find(([_, col]) => col.label === 'Trips');
    if (!tripsConfig || !tripsConfig[1].fileColumn) {
        alert('Please map the "Trips" column from your file');
        return;
    }

    // Process data using the mappings
    processedData = rawData.map(row => {
        const processed = {};
        enabledConfig.forEach(([configKey, configCol]) => {
            const fileColumn = configCol.fileColumn;
            processed[configCol.label] = row[fileColumn] || '';
        });
        return processed;
    }).filter(row => row['Trips']); // Filter out empty trips

    // Hide configuration, show dashboard
    columnConfigPanel.classList.add('hidden');
    dataPreviewPanel.classList.add('hidden');
    dashboardSection.classList.remove('hidden');

    // Generate dashboard
    generateDashboard();
    
    // Scroll to dashboard
    dashboardSection.scrollIntoView({ behavior: 'smooth' });
}

// Generate dashboard with charts and statistics
function generateDashboard() {
    calculateStatistics();
    generateTopRoutesChart();
    generateOriginStationsChart();
    generateOriginMesoChart();
    generateDestMesoChart();
    generateDataTable();
}

// Calculate statistics
function calculateStatistics() {
    const totalTrips = processedData.reduce((sum, row) => {
        const trips = parseFloat(row['Trips']) || 0;
        return sum + trips;
    }, 0);

    const avgTrips = totalTrips / processedData.length;
    const uniqueStations = new Set(processedData.map(row => row['From Station'])).size;
    const uniqueMesoZones = new Set([
        ...processedData.map(row => row['From Meso Zone']),
        ...processedData.map(row => row['To Meso Zone'])
    ]).size;

    document.getElementById('statTotalTrips').textContent = totalTrips.toLocaleString('en-US', {
        maximumFractionDigits: 0
    });
    document.getElementById('statAvgTrips').textContent = avgTrips.toLocaleString('en-US', {
        maximumFractionDigits: 0
    });
    document.getElementById('statUniqueStations').textContent = uniqueStations;
    document.getElementById('statUniqueMesoZones').textContent = uniqueMesoZones;
}

// Generate top routes chart
function generateTopRoutesChart() {
    const routeData = {};
    
    processedData.forEach(row => {
        const route = `${row['From Station']} → ${row['To Meso Zone']}`;
        const trips = parseFloat(row['Trips']) || 0;
        routeData[route] = (routeData[route] || 0) + trips;
    });

    const sorted = Object.entries(routeData)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const ctx = document.getElementById('topRoutesChart').getContext('2d');
    
    if (charts.topRoutes) {
        charts.topRoutes.destroy();
    }

    charts.topRoutes = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(item => item[0].substring(0, 30) + '...'),
            datasets: [{
                label: 'Trips',
                data: sorted.map(item => item[1]),
                backgroundColor: 'rgba(102, 126, 234, 0.6)',
                borderColor: 'rgba(102, 126, 234, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: {
                legend: { display: false }
            }
        }
    });
}

// Generate origin stations chart
function generateOriginStationsChart() {
    const stationData = {};
    
    processedData.forEach(row => {
        const station = row['From Station'];
        const trips = parseFloat(row['Trips']) || 0;
        stationData[station] = (stationData[station] || 0) + trips;
    });

    const sorted = Object.entries(stationData)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    const ctx = document.getElementById('originStationsChart').getContext('2d');
    
    if (charts.originStations) {
        charts.originStations.destroy();
    }

    charts.originStations = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: sorted.map(item => item[0].substring(0, 20)),
            datasets: [{
                data: sorted.map(item => item[1]),
                backgroundColor: [
                    'rgba(102, 126, 234, 0.8)',
                    'rgba(118, 75, 162, 0.8)',
                    'rgba(237, 100, 166, 0.8)',
                    'rgba(255, 154, 158, 0.8)',
                    'rgba(250, 208, 196, 0.8)',
                    'rgba(255, 235, 59, 0.8)',
                    'rgba(102, 187, 106, 0.8)',
                    'rgba(66, 133, 244, 0.8)',
                    'rgba(156, 39, 176, 0.8)',
                    'rgba(0, 188, 212, 0.8)'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom' }
            }
        }
    });
}

// Generate origin meso zone chart
function generateOriginMesoChart() {
    const mesoData = {};
    
    processedData.forEach(row => {
        const meso = row['From Meso Zone'];
        const trips = parseFloat(row['Trips']) || 0;
        mesoData[meso] = (mesoData[meso] || 0) + trips;
    });

    const sorted = Object.entries(mesoData)
        .sort((a, b) => b[1] - a[1]);

    const ctx = document.getElementById('originMesoChart').getContext('2d');
    
    if (charts.originMeso) {
        charts.originMeso.destroy();
    }

    charts.originMeso = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(item => item[0]),
            datasets: [{
                label: 'Trips',
                data: sorted.map(item => item[1]),
                backgroundColor: 'rgba(118, 75, 162, 0.6)',
                borderColor: 'rgba(118, 75, 162, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            }
        }
    });
}

// Generate destination meso zone chart
function generateDestMesoChart() {
    const mesoData = {};
    
    processedData.forEach(row => {
        const meso = row['To Meso Zone'];
        const trips = parseFloat(row['Trips']) || 0;
        mesoData[meso] = (mesoData[meso] || 0) + trips;
    });

    const sorted = Object.entries(mesoData)
        .sort((a, b) => b[1] - a[1]);

    const ctx = document.getElementById('destMesoChart').getContext('2d');
    
    if (charts.destMeso) {
        charts.destMeso.destroy();
    }

    charts.destMeso = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: sorted.map(item => item[0]),
            datasets: [{
                label: 'Trips',
                data: sorted.map(item => item[1]),
                backgroundColor: 'rgba(237, 100, 166, 0.6)',
                borderColor: 'rgba(237, 100, 166, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            }
        }
    });
}

// Generate data table
function generateDataTable() {
    const thead = document.getElementById('dataTableHead');
    const tbody = document.getElementById('dataTableBody');
    
    // Get all columns to display
    const allColumns = Object.keys(processedData[0] || {});
    
    // Create header
    thead.innerHTML = `<tr>${allColumns.map(col => `<th>${col}</th>`).join('')}</tr>`;
    
    // Clear body
    tbody.innerHTML = '';
    
    // Add rows (limit to 100 for performance)
    processedData.slice(0, 100).forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = allColumns.map(col => {
            const value = row[col] || '-';
            return `<td>${value}</td>`;
        }).join('');
        tbody.appendChild(tr);
    });

    // Add note if more than 100 rows
    if (processedData.length > 100) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="${allColumns.length}" class="text-center text-muted">
            Showing 100 of ${processedData.length} rows
        </td>`;
        tbody.appendChild(tr);
    }
}

// Reset dashboard
function resetDashboard() {
    rawData = [];
    processedData = [];
    
    fileInput.value = '';
    
    dashboardSection.classList.add('hidden');
    columnConfigPanel.classList.add('hidden');
    dataPreviewPanel.classList.add('hidden');
    fileInfo.classList.add('hidden');
    
    rowCount.textContent = '0';
    colCount.textContent = '0';
    fileStatus.textContent = '-';
    
    Object.values(charts).forEach(chart => {
        if (chart) chart.destroy();
    });
    charts = {};
    
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Format large numbers with thousand separators
function formatNumber(num) {
    return num.toLocaleString('en-US');
}
