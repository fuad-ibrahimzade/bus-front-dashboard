// Global variables
let rawData = [];
let processedData = [];
let charts = {};

const DEFAULT_COLUMNS = ['From Station', 'From Meso Zone', 'To Meso Zone', 'Trips'];

// DOM Elements
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const dashboardSection = document.getElementById('dashboardSection');
const loadingSpinner = document.getElementById('loadingSpinner');
const rowCount = document.getElementById('rowCount');
const colCount = document.getElementById('colCount');
const fileStatus = document.getElementById('fileStatus');
const resetButton = document.getElementById('resetButton');

// Setup event listeners
document.addEventListener('DOMContentLoaded', () => {
    fileInput.addEventListener('change', handleFileUpload);
    resetButton.addEventListener('click', resetDashboard);
});

// Parse file (CSV or Excel)
function handleFileUpload(e) {
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
            processFileData();
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
            processFileData();
        } catch (error) {
            alert('Error parsing Excel: ' + error.message);
            console.error(error);
            loadingSpinner.style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

// Process file data and auto-detect default columns
function processFileData() {
    if (rawData.length === 0) {
        alert('File is empty');
        loadingSpinner.style.display = 'none';
        return;
    }

    // Get all available columns
    const availableColumns = Object.keys(rawData[0]);
    
    // Update stats
    rowCount.textContent = rawData.length;
    colCount.textContent = availableColumns.length;
    fileStatus.textContent = '✓ Ready';
    fileInfo.classList.remove('hidden');

    // Auto-detect default columns (exact match)
    const detectedColumns = {};
    DEFAULT_COLUMNS.forEach(col => {
        const found = availableColumns.find(c => c === col);
        if (found) {
            detectedColumns[col] = found;
        }
    });

    if (DEFAULT_COLUMNS.length > 0 && Object.keys(detectedColumns).length === 0) {
        alert('Could not find the default columns in your file. Expected: ' + DEFAULT_COLUMNS.join(', '));
        loadingSpinner.style.display = 'none';
        return;
    }

    // Process data using the detected columns
    processedData = rawData.map(row => {
        const processed = {};
        DEFAULT_COLUMNS.forEach(col => {
            if (detectedColumns[col]) {
                processed[col] = row[detectedColumns[col]] || '';
            }
        });
        return processed;
    }).filter(row => row['Trips']); // Filter out empty trips

    loadingSpinner.style.display = 'none';

    // Show dashboard
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
