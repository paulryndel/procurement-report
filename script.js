// --- JAVASCRIPT LOGIC ---

// --- Live Clock Logic ---
function updateDateTime() {
    const clockElement = document.getElementById('datetime-clock');
    if (clockElement) {
        const now = new Date();
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit', second: '2-digit' };
        clockElement.textContent = now.toLocaleString('en-US', options);
    }
}
updateDateTime();
setInterval(updateDateTime, 1000);

// --- Tab Switching Logic ---
document.addEventListener('DOMContentLoaded', () => {
    const tabs = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');
    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const target = tab.getAttribute('data-tab');
            tabs.forEach(t => t.classList.remove('active'));
            tabContents.forEach(c => c.style.display = 'none');
            tab.classList.add('active');
            document.getElementById(target).style.display = 'block';
        });
    });
    addSortEventListeners();
    // Register the datalabels plugin globally
    Chart.register(ChartDataLabels);
});

// --- Global variables for data, filters, sorting, and graphing ---
let rawData = [];
let mrpData = {};
let allProductCodes = [];
let allProducts = {};
let displayedCount = 0;
const rowsPerLoad = 200;
let uniquePartsAndNames = [];
let uniqueVendors = [];
let selectedParts = [];
let selectedVendors = [];
let currentSort = { key: 'code', direction: 'asc' };
let selectedForGraph = [];
let chartInstance = null;
let priceChartInstance = null; // New chart instance for price graph

// --- Event Listeners ---
document.getElementById('excel-upload').addEventListener('change', handleProcurementFile);
document.getElementById('mrp-upload').addEventListener('change', handleMrpFile);
document.getElementById('load-more-btn').addEventListener('click', loadMoreData);
document.getElementById('part-filter-btn').addEventListener('click', () => toggleDropdown('part-filter-dropdown'));
document.getElementById('vendor-filter-btn').addEventListener('click', () => toggleDropdown('vendor-filter-dropdown'));
document.getElementById('part-search-input').addEventListener('input', () => filterList('part-search-input', 'part-list'));
document.getElementById('vendor-search-input').addEventListener('input', () => filterList('vendor-search-input', 'vendor-list'));
document.getElementById('need-stock-filter').addEventListener('change', applyFiltersAndRender);
document.getElementById('clear-filters-btn').addEventListener('click', clearFiltersAndSort);
document.getElementById('print-btn').addEventListener('click', printReport);

// --- Main File Handling ---
function handleProcurementFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    const loadingOverlay = document.getElementById('loading-overlay');
    const statusIcon = document.getElementById('procurement-status-icon');
    loadingOverlay.style.display = 'flex';
    statusIcon.classList.remove('success');

    const reader = new FileReader();
    reader.onload = (e) => {
        setTimeout(() => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const worksheet = workbook.Sheets['Query1'];
            if (!worksheet) {
                alert("Error: Sheet named 'Query1' not found.");
                loadingOverlay.style.display = 'none';
                return;
            }
            rawData = XLSX.utils.sheet_to_json(worksheet);
            
            populateFilters();
            applyFiltersAndRender();
            statusIcon.classList.add('success');
            loadingOverlay.style.display = 'none';
        }, 50);
    };
    reader.onerror = () => {
        alert("Error reading the file.");
        loadingOverlay.style.display = 'none';
    };
    reader.readAsArrayBuffer(file);
}

function handleMrpFile(event) {
    const file = event.target.files[0];
    if (!file) return;
    const loadingOverlay = document.getElementById('loading-overlay');
    const statusIcon = document.getElementById('mrp-status-icon');
    loadingOverlay.style.display = 'flex';
    statusIcon.classList.remove('success');

    const reader = new FileReader();
    reader.onload = (e) => {
        setTimeout(() => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const worksheet = workbook.Sheets['MRP'];
            if (!worksheet) {
                alert("Error: Sheet named 'MRP' not found.");
                loadingOverlay.style.display = 'none';
                return;
            }
            const mrpRawData = XLSX.utils.sheet_to_json(worksheet);
            mrpData = processMrpData(mrpRawData);
            
            applyFiltersAndRender();
            statusIcon.classList.add('success');
            loadingOverlay.style.display = 'none';
        }, 50);
    };
    reader.onerror = () => {
        alert("Error reading the MRP file.");
        loadingOverlay.style.display = 'none';
    };
    reader.readAsArrayBuffer(file);
}

// --- Data Processing, Filtering, and Sorting ---
function applyFiltersAndRender() {
    let filteredRawData = rawData;

    if (selectedVendors.length > 0) {
        filteredRawData = filteredRawData.filter(row => selectedVendors.includes(row['Vendor Name']));
    }

    if (selectedParts.length > 0) {
        filteredRawData = filteredRawData.filter(row => {
            const partCode = String(row['Product Code'] || '').trim();
            const partName = String(row['Product Name'] || '').trim();
            return selectedParts.includes(partCode) || selectedParts.includes(partName);
        });
    }

    allProducts = processProcurementData(filteredRawData);
    
    const needStockFilter = document.getElementById('need-stock-filter').value;
    if (needStockFilter !== 'all') {
        const filteredCodes = Object.keys(allProducts).filter(code => {
            const p = allProducts[code];
            const productMrp = mrpData[code] || { storeStock: 0 };
            const needStock = (productMrp.storeStock - p.lowLimit <= 0) ? "YES" : "NO";
            return needStock === needStockFilter;
        });
        const tempFilteredProducts = {};
        filteredCodes.forEach(code => tempFilteredProducts[code] = allProducts[code]);
        allProducts = tempFilteredProducts;
    }

    allProductCodes = Object.keys(allProducts);
    sortData();
    updateSortIcons();
    updateSummaryBoxes();

    document.getElementById('product-table-body').innerHTML = '';
    displayedCount = 0;
    loadMoreData();
    
    const totalProducts = allProductCodes.length;
    document.getElementById('loading-message').textContent = `Showing ${Math.min(displayedCount, totalProducts)} of ${totalProducts} products.`;
    
    selectedForGraph = [];
    updateGraphAndDetails();
}

/**
 * NEW: Robustly parses a value that might be a number or a formatted string (e.g., currency).
 * @param {*} value The value to parse.
 * @returns {number} The parsed number, or 0 if parsing fails.
 */
function parseNumericValue(value) {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string') return 0;
    // Remove currency symbols, commas, and whitespace, then parse
    const cleanedValue = value.replace(/[^0-9.-]+/g, "");
    return parseFloat(cleanedValue) || 0;
}


function processProcurementData(data) {
    const products = {};
    data.forEach(row => {
        const trimmedRow = {};
        for (const key in row) trimmedRow[key.trim()] = row[key];

        const originalCode = String(trimmedRow['Product Code'] || '').trim();
        const date = trimmedRow['Date'];

        if (!originalCode || !(date instanceof Date) || isNaN(date.getTime())) return;
        
        let normalizedCode = originalCode;
        if (normalizedCode.toUpperCase().startsWith('MI')) {
            normalizedCode = 'ML' + normalizedCode.substring(2);
        }

        const year = date.getFullYear();
        // UPDATED: Use the robust parsing function
        const quantity = parseNumericValue(trimmedRow['Quantity']);
        const unitPrice = parseNumericValue(trimmedRow['Unit Price']);

        if (!products[normalizedCode]) {
            products[normalizedCode] = {
                name: trimmedRow['Product Name'],
                vendor: trimmedRow['Vendor Name'],
                firstOrderDate: date,
                lastOrderDate: date,
                latestUnitPrice: unitPrice,
                latestQuantity: quantity,
                years: {
                    2020: { qty: 0, priceSum: 0, count: 0 }, 2021: { qty: 0, priceSum: 0, count: 0 },
                    2022: { qty: 0, priceSum: 0, count: 0 }, 2023: { qty: 0, priceSum: 0, count: 0 },
                    2024: { qty: 0, priceSum: 0, count: 0 }, 2025: { qty: 0, priceSum: 0, count: 0 }
                },
                total: 0, aveQty: 0, safeStock: 0, lowLimit: 0, orderCount: 0
            };
        }

        if (date < products[normalizedCode].firstOrderDate) {
            products[normalizedCode].firstOrderDate = date;
            products[normalizedCode].name = trimmedRow['Product Name'];
            products[normalizedCode].vendor = trimmedRow['Vendor Name'];
        }
        if (date >= products[normalizedCode].lastOrderDate) {
            products[normalizedCode].lastOrderDate = date;
            products[normalizedCode].latestUnitPrice = unitPrice;
            products[normalizedCode].latestQuantity = quantity;
        }

        if (year >= 2020 && year <= 2025) {
            const yearData = products[normalizedCode].years[year];
            yearData.qty += quantity;
            if (unitPrice > 0) { // Only include orders with a price in average calculation
                yearData.priceSum += unitPrice;
                yearData.count++;
            }
        }
        products[normalizedCode].orderCount++;
    });

    const currentYear = new Date().getFullYear();
    for (const code in products) {
        const p = products[code];
        p.total = Object.values(p.years).reduce((sum, yearData) => sum + yearData.qty, 0);
        const firstOrderYear = p.firstOrderDate.getFullYear();
        if (p.total > 0) {
            const numYears = (currentYear - firstOrderYear) + 1;
            if (numYears > 0) p.aveQty = p.total / numYears;
        }
        p.safeStock = Math.round(p.aveQty * 0.4);
        p.lowLimit = Math.round(p.aveQty * 0.2);
    }
    return products;
}


function processMrpData(data) {
    const tempMrpData = {};
    data.forEach(row => {
        const originalCode = String(row['Products'] || '').trim();
        if (!originalCode) return;

        let normalizedCode = originalCode;
        if (normalizedCode.toUpperCase().startsWith('MI')) {
            normalizedCode = 'ML' + normalizedCode.substring(2);
        }

        if (!tempMrpData[normalizedCode]) tempMrpData[normalizedCode] = [];
        tempMrpData[normalizedCode].push(row);
    });

    const finalMrpData = {};
    for (const productCode in tempMrpData) {
        let minRow = tempMrpData[productCode].reduce((prev, curr) => 
            (parseFloat(prev['ThisTimeBalance']) || Infinity) < (parseFloat(curr['ThisTimeBalance']) || Infinity) ? prev : curr
        );
        finalMrpData[productCode] = {
            mrpBalance: parseFloat(minRow['MRPBalance']) || 0,
            storeStock: parseFloat(minRow['StockOnHand']) || 0,
            woPo: (parseFloat(minRow['AllWO']) || 0) + (parseFloat(minRow['AllPO']) || 0)
        };
    }
    return finalMrpData;
}

function sortData() {
    const { key, direction } = currentSort;
    const modifier = direction === 'asc' ? 1 : -1;

    allProductCodes.sort((a, b) => {
        const productA = allProducts[a];
        const productB = allProducts[b];
        
        let valA, valB;

        if (key === 'code') { valA = a; valB = b; }
        else if (key === 'name') { valA = productA.name; valB = productB.name; }
        else if (key.startsWith('y20')) { 
            const year = key.substring(1);
            valA = productA.years[year] ? productA.years[year].qty : 0;
            valB = productB.years[year] ? productB.years[year].qty : 0;
        }
        else if (key === 'needStock' || key === 'pcsNeeded') {
            const mrpA = mrpData[a] || { storeStock: 0 };
            const mrpB = mrpData[b] || { storeStock: 0 };
            const needA = (mrpA.storeStock - productA.lowLimit <= 0) ? "YES" : "NO";
            const needB = (mrpB.storeStock - productB.lowLimit <= 0) ? "YES" : "NO";
            if (key === 'needStock') { valA = needA; valB = needB; }
            else {
                valA = (needA === "YES") ? productA.lowLimit - mrpA.storeStock : 0;
                valB = (needB === "YES") ? productB.lowLimit - mrpB.storeStock : 0;
            }
        } else {
            const mrpA = mrpData[a] || { mrpBalance: 0, storeStock: 0, woPo: 0 };
            const mrpB = mrpData[b] || { mrpBalance: 0, storeStock: 0, woPo: 0 };
            const combinedA = { ...productA, ...mrpA };
            const combinedB = { ...productB, ...mrpB };
            valA = combinedA[key];
            valB = combinedB[key];
        }

        if (typeof valA === 'string' && typeof valB === 'string') {
            return valA.localeCompare(valB) * modifier;
        }
        return (valA - valB) * modifier;
    });
}

// --- UI and Helper Functions ---
function loadMoreData() {
    const tableBody = document.getElementById('product-table-body');
    const start = displayedCount;
    const end = Math.min(start + rowsPerLoad, allProductCodes.length);
    const codesToRender = allProductCodes.slice(start, end);
    appendRowsToTable(codesToRender);
    displayedCount = end;
    document.getElementById('load-more-btn').style.display = displayedCount >= allProductCodes.length ? 'none' : 'block';
}

function appendRowsToTable(codesToRender) {
    const tableBody = document.getElementById('product-table-body');
    codesToRender.forEach(code => {
        const p = allProducts[code];
        const row = document.createElement('tr');
        row.dataset.code = code;
        row.addEventListener('click', handleRowClick);
        
        if (selectedForGraph.includes(code)) {
            row.classList.add('selected-row');
        }

        const fullName = p.name || 'N/A';
        const truncatedName = fullName.length > 15 ? fullName.substring(0, 15) + '...' : fullName;
        
        const productMrp = mrpData[code] || { mrpBalance: 0, storeStock: 0, woPo: 0 };

        const storeStock = productMrp.storeStock;
        const lowLimit = p.lowLimit;
        const needStock = (storeStock - lowLimit <= 0) ? "YES" : "NO";
        const pcsNeeded = (needStock === "YES") ? lowLimit - storeStock : 0;
        
        const highlightClass = (needStock === "YES") ? "highlight-red" : "";

        row.innerHTML = `
            <td class="font-semibold sticky-col-1">${code}</td>
            <td title="${fullName}" class="sticky-col-2">${truncatedName}</td>
            <td class="text-center">${p.years[2020].qty}</td><td class="text-center">${p.years[2021].qty}</td>
            <td class="text-center">${p.years[2022].qty}</td><td class="text-center">${p.years[2023].qty}</td>
            <td class="text-center">${p.years[2024].qty}</td><td class="text-center">${p.years[2025].qty}</td>
            <td class="text-center font-bold bg-yellow-100">${p.total}</td>
            <td class="text-center">${p.aveQty.toFixed(2)}</td><td class="text-center">${p.safeStock}</td>
            <td class="text-center ${highlightClass}">${lowLimit}</td>
            <td class="text-center">${productMrp.mrpBalance}</td>
            <td class="text-center">${storeStock}</td>
            <td class="text-center">${productMrp.woPo}</td>
            <td class="text-center ${highlightClass}">${needStock}</td>
            <td class="text-center">${pcsNeeded}</td>
        `;
        tableBody.appendChild(row);
    });
}

function updateSummaryBoxes() {
    let itemsNeedStock = 0;
    let piecesNeedStock = 0;
    const vendorsToOrder = new Set();

    allProductCodes.forEach(code => {
        const p = allProducts[code];
        const productMrp = mrpData[code] || { storeStock: 0 };
        const needStock = (productMrp.storeStock - p.lowLimit <= 0);
        if (needStock) {
            itemsNeedStock++;
            piecesNeedStock += (p.lowLimit - productMrp.storeStock);
            if (p.vendor) {
                vendorsToOrder.add(p.vendor);
            }
        }
    });

    document.getElementById('items-to-stock').textContent = Math.round(itemsNeedStock);
    document.getElementById('pieces-to-stock').textContent = Math.round(piecesNeedStock);
    document.getElementById('order-supplier').textContent = vendorsToOrder.size;
}

function addSortEventListeners() {
    document.querySelectorAll('.sortable').forEach(header => {
        const sortKey = header.dataset.sortKey;
        header.innerHTML = header.textContent + `<span class="sort-icon" data-sort-key="${sortKey}"></span>`;
        header.addEventListener('click', () => {
            if (currentSort.key === sortKey) {
                currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                currentSort.key = sortKey;
                currentSort.direction = 'asc';
            }
            applyFiltersAndRender();
        });
    });
}

function updateSortIcons() {
    document.querySelectorAll('.sort-icon').forEach(icon => {
        icon.classList.remove('active');
        icon.innerHTML = '&#8693;';
        if (icon.dataset.sortKey === currentSort.key) {
            icon.classList.add('active');
            icon.innerHTML = currentSort.direction === 'asc' ? '&#8593;' : '&#8595;';
        }
    });
}

function populateFilters() {
    const partCodes = rawData.map(row => String(row['Product Code'] || '').trim());
    const partNames = rawData.map(row => String(row['Product Name'] || '').trim());
    uniquePartsAndNames = [...new Set([...partCodes, ...partNames])].filter(Boolean).sort();
    
    uniqueVendors = [...new Set(rawData.map(row => row['Vendor Name']).filter(Boolean))].sort();

    populateCheckboxList('part-list', uniquePartsAndNames, 'handlePartSelection(this)');
    populateCheckboxList('vendor-list', uniqueVendors, 'handleVendorSelection(this)');
}

function populateCheckboxList(listId, items, onchangeAction) {
    const list = document.getElementById(listId);
    list.innerHTML = '';
    items.forEach(item => {
        const label = document.createElement('label');
        label.innerHTML = `<input type="checkbox" value="${item}" onchange="${onchangeAction}"> ${item}`;
        list.appendChild(label);
    });
}

function handlePartSelection(checkbox) {
    handleSelection(checkbox, selectedParts, 'part');
}

function handleVendorSelection(checkbox) {
    handleSelection(checkbox, selectedVendors, 'vendor');
}

function handleSelection(checkbox, selectedArray, type) {
    const value = checkbox.value;
    if (checkbox.checked) {
        if (!selectedArray.includes(value)) selectedArray.push(value);
    } else {
        const index = selectedArray.indexOf(value);
        if (index > -1) selectedArray.splice(index, 1);
    }
    updateSelectedTags();
    applyFiltersAndRender();
}

function updateSelectedTags() {
    const container = document.getElementById('selected-filters-container');
    container.innerHTML = '';
    selectedParts.forEach(item => createTag(item, 'part', container));
    selectedVendors.forEach(item => createTag(item, 'vendor', container));
}

function createTag(item, type, container) {
    const tag = document.createElement('div');
    tag.className = `filter-tag ${type}`;
    tag.innerHTML = `<span>${item}</span><button onclick="removeTag('${item.replace(/'/g, "\\'")}', '${type}')">&times;</button>`;
    container.appendChild(tag);
}

function removeTag(item, type) {
    let selectedArray = type === 'part' ? selectedParts : selectedVendors;
    const index = selectedArray.indexOf(item);
    if (index > -1) selectedArray.splice(index, 1);

    const listId = type === 'part' ? 'part-list' : 'vendor-list';
    const checkbox = document.querySelector(`#${listId} input[value="${item.replace(/"/g, '\\"')}"]`);
    if (checkbox) checkbox.checked = false;

    updateSelectedTags();
    applyFiltersAndRender();
}

function filterList(inputId, listId) {
    const searchTerm = document.getElementById(inputId).value.toLowerCase();
    const labels = document.querySelectorAll(`#${listId} label`);
    labels.forEach(label => {
        const text = label.textContent.trim().toLowerCase();
        label.style.display = text.includes(searchTerm) ? 'block' : 'none';
    });
}

function toggleDropdown(dropdownId) {
    document.getElementById(dropdownId).classList.toggle('hidden');
}

function clearFiltersAndSort() {
    selectedParts = [];
    selectedVendors = [];
    document.getElementById('need-stock-filter').value = 'all';
    
    currentSort = { key: 'code', direction: 'asc' };

    document.querySelectorAll('.filter-list input[type="checkbox"]').forEach(cb => cb.checked = false);
    updateSelectedTags();
    
    applyFiltersAndRender();
}

function debounce(func, delay) {
    let timeout;
    return (...args) => {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), delay);
    };
}

// --- Graphing and Printing Functions ---
function handleRowClick(event) {
    const row = event.currentTarget;
    const code = row.dataset.code;
    
    row.classList.toggle('selected-row');

    const index = selectedForGraph.indexOf(code);
    if (index > -1) {
        selectedForGraph.splice(index, 1);
    } else {
        selectedForGraph.push(code);
    }
    
    updateGraphAndDetails();
}

function updateGraphAndDetails() {
    const detailsSection = document.getElementById('details-section');
    const qtyCtx = document.getElementById('history-chart').getContext('2d');
    const priceCtx = document.getElementById('price-history-chart').getContext('2d');

    if (selectedForGraph.length === 0) {
        detailsSection.classList.add('hidden');
        return;
    }

    detailsSection.classList.remove('hidden');

    if (chartInstance) chartInstance.destroy();
    if (priceChartInstance) priceChartInstance.destroy();
    
    const years = ['2020', '2021', '2022', '2023', '2024', '2025'];
    
    // --- Quantity Chart ---
    const qtyDatasets = selectedForGraph.map((code) => {
        const product = allProducts[code];
        const data = years.map(year => product.years[year].qty);
        const color = `rgba(${Math.floor(Math.random() * 155) + 50}, ${Math.floor(Math.random() * 155) + 50}, ${Math.floor(Math.random() * 155) + 50}, 1)`;
        return { label: code, data: data, borderColor: color, backgroundColor: color.replace('1)', '0.2)'), fill: true, tension: 0.1, pointBackgroundColor: color, pointRadius: 5, pointHoverRadius: 7 };
    });

    chartInstance = new Chart(qtyCtx, {
        type: 'line', data: { labels: years, datasets: qtyDatasets },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' }, title: { display: true, text: 'Yearly Purchase Quantity History' }, datalabels: { align: 'top', anchor: 'end', backgroundColor: (context) => context.dataset.borderColor, borderRadius: 4, color: 'white', font: { size: 8, weight: 'bold' }, formatter: (value) => value > 0 ? value : '', padding: 4 } }, scales: { y: { beginAtZero: true, title: { display: true, text: 'Quantity' } }, x: { title: { display: true, text: 'Year' } } } }
    });

    // --- Price Chart ---
    const priceDatasets = selectedForGraph.map((code) => {
        const product = allProducts[code];
        const data = years.map(year => {
            const yearData = product.years[year];
            return yearData.count > 0 ? (yearData.priceSum / yearData.count) : null;
        });
        const color = qtyDatasets.find(d => d.label === code).borderColor;
        return { label: code, data: data, borderColor: color, backgroundColor: color.replace('1)', '0.2)'), fill: true, tension: 0.1, pointBackgroundColor: color, pointRadius: 5, pointHoverRadius: 7 };
    });
    
    // Add horizontal line for latest price IF only one product is selected
    if (selectedForGraph.length === 1) {
        const product = allProducts[selectedForGraph[0]];
        if (product && product.latestUnitPrice > 0) {
            priceDatasets.push({
                label: `Latest Price (${product.latestUnitPrice.toFixed(2)})`,
                data: Array(years.length).fill(product.latestUnitPrice),
                borderColor: '#ef4444',
                borderWidth: 2,
                borderDash: [5, 5],
                type: 'line',
                fill: false,
                pointRadius: 0,
                pointHoverRadius: 0,
                datalabels: { display: false }
            });
        }
    }

    priceChartInstance = new Chart(priceCtx, {
        type: 'line', data: { labels: years, datasets: priceDatasets },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' }, title: { display: true, text: 'Yearly Average Unit Price History' }, datalabels: { align: 'top', anchor: 'end', backgroundColor: (context) => context.dataset.borderColor, borderRadius: 4, color: 'white', font: { size: 8, weight: 'bold' }, formatter: (value) => value ? value.toFixed(2) : '', padding: 4 } }, scales: { y: { beginAtZero: false, title: { display: true, text: 'Average Unit Price' } }, x: { title: { display: true, text: 'Year' } } } }
    });

    updateDetailsTable(qtyDatasets);
}

function updateDetailsTable(datasets) {
    const tableBody = document.getElementById('details-table-body');
    tableBody.innerHTML = '';

    selectedForGraph.forEach((code, index) => {
        const product = allProducts[code];
        const color = datasets[index].borderColor;

        const fullProductName = product.name || 'N/A';
        const truncatedProductName = fullProductName.length > 20 ? fullProductName.substring(0, 20) + '...' : fullProductName;
        const fullVendorName = product.vendor || 'N/A';
        const truncatedVendorName = fullVendorName.length > 20 ? fullVendorName.substring(0, 20) + '...' : fullVendorName;
        
        // UPDATED: Format date to Month-Year (e.g., Jan-2025)
        let formattedDate = 'N/A';
        if (product.lastOrderDate) {
            const d = product.lastOrderDate;
            const month = d.toLocaleString('en-US', { month: 'short' });
            const year = d.getFullYear();
            formattedDate = `${month}-${year}`;
        }

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td style="background-color: ${color.replace('1)', '0.2)')}; color: ${color}; font-weight: bold;">${code}</td>
            <td title="${fullProductName}">${truncatedProductName}</td>
            <td title="${fullVendorName}">${truncatedVendorName}</td>
            <td class="text-center">${product.orderCount}</td>
            <td class="text-center">${product.total}</td>
            <td class="text-center bg-yellow-100 font-semibold">${product.latestUnitPrice.toFixed(2)}</td>
            <td class="text-center bg-yellow-100 font-semibold">${formattedDate}</td>
            <td class="text-center bg-yellow-100 font-semibold">${product.latestQuantity}</td>
        `;
        tableBody.appendChild(row);
    });
}

function printReport() {
    const captureArea = document.getElementById('capture-area');
    html2canvas(captureArea, {
        useCORS: true,
        allowTaint: true,
        scrollX: 0,
        scrollY: -window.scrollY,
        windowWidth: document.documentElement.offsetWidth,
        windowHeight: document.documentElement.offsetHeight
    }).then(canvas => {
        const link = document.createElement('a');
        link.download = 'stock-report.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    });
}