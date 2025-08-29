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
    addMrpSortEventListeners(); // Add listener for new sortable table
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
let priceChartInstance = null;
let filterNegativeMrp = false;
let mrpRawDataForReport = [];
let mrpReportFilter = 'all'; // 'all', 'check', 'x'
let mrpReportSort = { key: 'week', direction: 'asc' }; // For MRP report sorting
let mrpChartInstance = null;
let mrpPriceChartInstance = null;
let mrpSelectedForGraph = [];
let mrpSelectedCustomers = [];
let mrpSelectedProductGroups = [];


// --- Event Listeners ---
document.getElementById('excel-upload').addEventListener('change', handleProcurementFile);
document.getElementById('mrp-upload').addEventListener('change', handleMrpFile);
document.getElementById('load-more-btn').addEventListener('click', loadMoreData);
document.getElementById('part-filter-btn').addEventListener('click', () => toggleDropdown('part-filter-dropdown'));
document.getElementById('vendor-filter-btn').addEventListener('click', () => toggleDropdown('vendor-filter-dropdown'));
document.getElementById('part-search-input').addEventListener('input', () => filterList('part-search-input', 'part-list'));
document.getElementById('vendor-search-input').addEventListener('input', () => filterList('vendor-search-input', 'vendor-list'));
document.getElementById('need-stock-filter').addEventListener('change', applyFiltersAndRender);
document.getElementById('mrp-filter-btn').addEventListener('click', toggleNegativeMrpFilter);
document.getElementById('clear-filters-btn').addEventListener('click', clearFiltersAndSort);
document.getElementById('print-btn').addEventListener('click', printReport);
// MRP Tab Listeners
document.getElementById('mrp-filter-all').addEventListener('click', () => setMrpReportFilter('all'));
document.getElementById('mrp-filter-check').addEventListener('click', () => setMrpReportFilter('check'));
document.getElementById('mrp-filter-x').addEventListener('click', () => setMrpReportFilter('x'));
document.getElementById('mrp-customer-filter-btn').addEventListener('click', () => toggleDropdown('mrp-customer-filter-dropdown'));
document.getElementById('mrp-pg-filter-btn').addEventListener('click', () => toggleDropdown('mrp-pg-filter-dropdown'));
document.getElementById('mrp-customer-search-input').addEventListener('input', () => filterList('mrp-customer-search-input', 'mrp-customer-list'));
document.getElementById('mrp-pg-search-input').addEventListener('input', () => filterList('mrp-pg-search-input', 'mrp-pg-list'));
document.getElementById('mrp-clear-search-filters-btn').addEventListener('click', clearMrpSearchFilters);


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
            // Process data for both tabs
            mrpData = processMrpData(mrpRawData);
            mrpRawDataForReport = processMrpReportData(mrpRawData);
            
            applyFiltersAndRender(); // Update main tab
            renderMrpReport(); // Update MRP report tab
            populateMrpSearchFilters(); // Populate new search filters
            
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

    // Filter by selected vendors
    if (selectedVendors.length > 0) {
        filteredRawData = filteredRawData.filter(row => selectedVendors.includes(row['Vendor Name']));
    }

    // Filter by selected parts
    if (selectedParts.length > 0) {
        filteredRawData = filteredRawData.filter(row => {
            const partCode = String(row['Product Code'] || '').trim();
            const partName = String(row['Product Name'] || '').trim();
            return selectedParts.includes(partCode) || selectedParts.includes(partName);
        });
    }

    // Process the data after basic filters
    allProducts = processProcurementData(filteredRawData);
    
    // Apply Negative MRP Filter
    if (filterNegativeMrp) {
        const filteredCodes = Object.keys(allProducts).filter(code => {
            const productMrp = mrpData[code] || { mrpBalance: 0 };
            return productMrp.mrpBalance < 0;
        });
        const tempFilteredProducts = {};
        filteredCodes.forEach(code => tempFilteredProducts[code] = allProducts[code]);
        allProducts = tempFilteredProducts;
    }
    
    // Apply "Need Stock?" dropdown filter
    const needStockFilter = document.getElementById('need-stock-filter').value;
    if (needStockFilter !== 'all') {
        const filteredCodes = Object.keys(allProducts).filter(code => {
            const p = allProducts[code];
            const productMrp = mrpData[code] || { mrpBalance: 0 };
            const needStock = (productMrp.mrpBalance <= p.lowLimit) ? "YES" : "NO";
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
        }
        if (date >= products[normalizedCode].lastOrderDate) {
            products[normalizedCode].lastOrderDate = date;
            products[normalizedCode].name = trimmedRow['Product Name'];
            products[normalizedCode].vendor = trimmedRow['Vendor Name'];
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
            mrpBalance: Math.round(parseFloat(minRow['MRPBalance']) || 0),
            storeStock: Math.round(parseFloat(minRow['StockOnHand']) || 0),
            woPo: Math.round((parseFloat(minRow['AllWO']) || 0) + (parseFloat(minRow['AllPO']) || 0))
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
            const mrpA = mrpData[a] || { mrpBalance: 0, storeStock: 0 };
            const mrpB = mrpData[b] || { mrpBalance: 0, storeStock: 0 };
            
            const needA = (mrpA.mrpBalance <= productA.lowLimit) ? "YES" : "NO";
            const needB = (mrpB.mrpBalance <= productB.lowLimit) ? "YES" : "NO";
            
            if (key === 'needStock') {
                valA = needA;
                valB = needB;
            } else { // key === 'pcsNeeded'
                if (needA === "YES") {
                    valA = productA.safeStock - mrpA.mrpBalance;
                } else {
                    valA = productA.safeStock - mrpA.storeStock;
                }
                if (needB === "YES") {
                    valB = productB.safeStock - mrpB.mrpBalance;
                } else {
                    valB = productB.safeStock - mrpB.storeStock;
                }
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
function toggleNegativeMrpFilter() {
    filterNegativeMrp = !filterNegativeMrp;
    const btn = document.getElementById('mrp-filter-btn');
    btn.classList.toggle('active', filterNegativeMrp);
    applyFiltersAndRender();
}

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

        const { storeStock, mrpBalance } = productMrp;
        const { lowLimit, safeStock } = p;

        // New calculation logic
        const needStock = (mrpBalance <= lowLimit) ? "YES" : "NO";
        let pcsNeeded = 0;
        if (needStock === "YES") {
            pcsNeeded = safeStock - mrpBalance;
        } else {
            pcsNeeded = safeStock - storeStock;
        }
        
        const highlightClass = (needStock === "YES") ? "highlight-red" : "";

        row.innerHTML = `
            <td class="font-semibold sticky-col-1">${code}</td>
            <td title="${fullName}" class="sticky-col-2">${truncatedName}</td>
            <td class="text-center">${p.years[2020].qty}</td><td class="text-center">${p.years[2021].qty}</td>
            <td class="text-center">${p.years[2022].qty}</td><td class="text-center">${p.years[2023].qty}</td>
            <td class="text-center">${p.years[2024].qty}</td><td class="text-center">${p.years[2025].qty}</td>
            <td class="text-center font-bold bg-yellow-100">${p.total}</td>
            <td class="text-center">${p.aveQty.toFixed(2)}</td><td class="text-center">${safeStock}</td>
            <td class="text-center">${lowLimit}</td>
            <td class="text-center ${highlightClass}">${mrpBalance}</td>
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
        const productMrp = mrpData[code] || { mrpBalance: 0, storeStock: 0 };
        
        const needStock = (productMrp.mrpBalance <= p.lowLimit);
        
        if (needStock) {
            itemsNeedStock++;
            const pcsNeeded = p.safeStock - productMrp.mrpBalance;
            if (pcsNeeded > 0) {
                piecesNeedStock += pcsNeeded;
            }
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
    document.querySelectorAll('#product-table .sortable').forEach(header => {
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
    document.querySelectorAll('#product-table .sort-icon').forEach(icon => {
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
    
    // Reset MRP filter state
    filterNegativeMrp = false;
    document.getElementById('mrp-filter-btn').classList.remove('active');

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

// --- MRP Report Tab Functions ---
function addMrpSortEventListeners() {
    document.querySelectorAll('#mrp-left-table .sortable').forEach(header => {
        const sortKey = header.dataset.sortKey;
        header.innerHTML += `<span class="sort-icon" data-sort-key="${sortKey}"></span>`;
        header.addEventListener('click', () => {
            if (mrpReportSort.key === sortKey) {
                mrpReportSort.direction = mrpReportSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                mrpReportSort.key = sortKey;
                mrpReportSort.direction = 'asc';
            }
            renderMrpReport();
        });
    });
}

function updateMrpSortIcons() {
    document.querySelectorAll('#mrp-left-table .sort-icon').forEach(icon => {
        icon.classList.remove('active');
        icon.innerHTML = '&#8693;';
        if (icon.dataset.sortKey === mrpReportSort.key) {
            icon.classList.add('active');
            icon.innerHTML = mrpReportSort.direction === 'asc' ? '&#8593;' : '&#8595;';
        }
    });
}

function processMrpReportData(data) {
    // Clean up column names and rename customer column
    return data.map(row => {
        const cleanedRow = {};
        for (const key in row) {
            cleanedRow[key.trim()] = row[key];
        }
        // Standardize the customer column name
        cleanedRow.Customer = cleanedRow['ชื่อลูกค้า'];
        return cleanedRow;
    });
}

function setMrpReportFilter(filter) {
    mrpReportFilter = filter;
    mrpSelectedForGraph = []; // Reset selection
    // Update active button visuals
    document.querySelectorAll('.mrp-filter-btn').forEach(btn => btn.classList.remove('active'));
    document.getElementById(`mrp-filter-${filter}`).classList.add('active');
    renderMrpReport();
    // Hide the right panel when filter changes
    document.getElementById('mrp-right-table-container').style.display = 'none';
    document.getElementById('mrp-details-section').classList.add('hidden');
    clearMrpSearchFilters();
}

function renderMrpReport() {
    if (mrpRawDataForReport.length === 0) return;

    // 1. Pre-aggregate data to determine customer status per week
    const weeklyCustomerStatus = mrpRawDataForReport.reduce((acc, row) => {
        let week = row.WeekFG;
        // Normalize week capitalization (e.g., "week49" and "WEEK49" become "Week49")
        if (typeof week === 'string' && week.toLowerCase().startsWith('week')) {
            week = 'W' + week.slice(1).toLowerCase();
        }
        
        const customer = row.Customer;
        if (!week || !customer) return acc;

        if (!acc[week]) {
            acc[week] = {};
        }
        if (!acc[week][customer]) {
            // A customer starts as 'check' (green) until a 'red' item is found
            acc[week][customer] = { status: 'check' }; 
        }

        // If any item for this customer in this week needs an order, the customer is 'x' (red) for the week
        if ((row.WeekStatus || '').includes('ให้สั่งผลิตเพิ่ม')) {
            acc[week][customer].status = 'x';
        }
        
        return acc;
    }, {});

    // 2. Calculate counts for the left table based on the filter
    const weeklyCounts = {};
    for (const week in weeklyCustomerStatus) {
        const customers = weeklyCustomerStatus[week];
        let count = 0;
        if (mrpReportFilter === 'all') {
            count = Object.keys(customers).length;
        } else {
            // Count customers whose status matches the current filter ('check' or 'x')
            count = Object.values(customers).filter(c => c.status === mrpReportFilter).length;
        }

        if (count > 0) {
            weeklyCounts[week] = count;
        }
    }

    // 3. Convert to array and apply sorting
    let leftTableData = Object.entries(weeklyCounts)
        .map(([week, count]) => ({ week, count }));

    const modifier = mrpReportSort.direction === 'asc' ? 1 : -1;
    leftTableData.sort((a, b) => {
        // Sort alphabetically by week name
        return a.week.localeCompare(b.week) * modifier;
    });

    // 4. Render the left table
    const leftTableBody = document.getElementById('mrp-left-table-body');
    leftTableBody.innerHTML = '';
    leftTableData.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.week}</td>
            <td>${item.count}</td>
        `;
        row.addEventListener('click', (event) => {
            clearMrpSearchFilters();
            document.querySelectorAll('#mrp-left-table-body tr').forEach(r => r.classList.remove('selected'));
            row.classList.add('selected');
            renderMrpDetailsTable(mrpRawDataForReport.filter(row => {
                let originalWeek = row.WeekFG;
                if (typeof originalWeek === 'string') {
                    return originalWeek.toLowerCase() === item.week.toLowerCase();
                }
                return false;
            }));
             document.getElementById('details-week-title').textContent = item.week;
        });
        leftTableBody.appendChild(row);
    });

    updateMrpSortIcons();
}

function renderMrpDetailsTable(data) {
    const rightPanel = document.getElementById('mrp-right-table-container');
    
    mrpSelectedForGraph = []; // Reset selection
    document.getElementById('mrp-details-section').classList.add('hidden');

    // 1. Pre-filter the data based on the global report filter (All, ✔, ✘)
    let weekData = data;
    if (mrpReportFilter === 'check') {
        weekData = weekData.filter(row => !(row.WeekStatus || '').includes('ให้สั่งผลิตเพิ่ม'));
    } else if (mrpReportFilter === 'x') {
        weekData = weekData.filter(row => (row.WeekStatus || '').includes('ให้สั่งผลิตเพิ่ม'));
    }

    if(weekData.length === 0) {
        rightPanel.style.display = 'none';
        return;
    }
    rightPanel.style.display = 'block';

    // 2. Pre-calculate status counts for each customer for the status bar
    const customerStatusCounts = {};
    weekData.forEach(row => {
        const customer = row.Customer || 'N/A';
        if (!customerStatusCounts[customer]) {
            customerStatusCounts[customer] = { red: 0, green: 0 };
        }
        const isRed = (row.WeekStatus || '').includes('ให้สั่งผลิตเพิ่ม');
        isRed ? customerStatusCounts[customer].red++ : customerStatusCounts[customer].green++;
    });

    // 3. Group data by Customer, then Model, and aggregate identical products
    const groupedData = weekData.reduce((acc, row) => {
        const customer = row.Customer || 'N/A';
        const model = row.Model || 'N/A';
        
        if (!acc[customer]) acc[customer] = {};
        if (!acc[customer][model]) acc[customer][model] = [];
        
        const productKey = `${row.Products}|${row.ProductName}|${row.Units}`;
        let productEntry = acc[customer][model].find(p => p.key === productKey);

        if (productEntry) {
            productEntry.Qty += parseFloat(row.Qty || 0);
        } else {
             acc[customer][model].push({
                key: productKey,
                Products: row.Products,
                ProductName: row.ProductName,
                Qty: parseFloat(row.Qty || 0),
                Units: row.Units,
                WeekStatus: row.WeekStatus
            });
        }
        return acc;
    }, {});

    // 4. Render the table with rowspan
    const rightTableBody = document.getElementById('mrp-right-table-body');
    rightTableBody.innerHTML = '';
    let customerIndex = 0;

    for (const customer in groupedData) {
        customerIndex++;
        const models = groupedData[customer];
        let isFirstCustomerRow = true;
        const customerRowSpan = Object.values(models).reduce((sum, products) => sum + products.length, 0);

        // Calculate percentages for the status bar
        const counts = customerStatusCounts[customer];
        const total = counts.red + counts.green;
        const redPercent = total > 0 ? (counts.red / total) * 100 : 0;
        const greenPercent = total > 0 ? (counts.green / total) * 100 : 0;
        const statusBarHtml = `
            <div class="status-bar-container">
                <div class="status-bar-green" style="width: ${greenPercent}%;"></div>
                <div class="status-bar-red" style="width: ${redPercent}%;"></div>
                <span class="status-bar-text">${Math.round(redPercent)}%</span>
            </div>
        `;

        for (const model in models) {
            const products = models[model];
            let isFirstModelRow = true;
            const modelRowSpan = products.length;

            products.forEach(product => {
                const row = document.createElement('tr');
                let rowHtml = '';

                if (isFirstCustomerRow) {
                    const fullCustomerName = customer;
                    const truncatedCustomerName = fullCustomerName.length > 35 ? fullCustomerName.substring(0, 35) + '...' : fullCustomerName;
                    
                    rowHtml += `<td rowspan="${customerRowSpan}">${customerIndex}</td>`;
                    rowHtml += `<td rowspan="${customerRowSpan}" title="${fullCustomerName}">${truncatedCustomerName}</td>`;
                }

                if (isFirstModelRow) {
                    rowHtml += `<td rowspan="${modelRowSpan}">${model}</td>`;
                    isFirstModelRow = false;
                }

                const status = product.WeekStatus || '';
                const highlightClass = status.includes('ให้สั่งผลิตเพิ่ม') ? 'status-red' : 'status-green';
                
                const fullProductName = product.ProductName || '';
                const truncatedProductName = fullProductName.length > 30 ? fullProductName.substring(0, 30) + '...' : fullProductName;

                rowHtml += `
                    <td class="clickable-product ${highlightClass}" data-product-code="${product.Products || ''}">${product.Products || ''}</td>
                    <td title="${fullProductName}">${truncatedProductName}</td>
                    <td>${product.Qty}</td>
                    <td>${product.Units || ''}</td>
                `;
                
                if (isFirstCustomerRow) {
                    rowHtml += `<td rowspan="${customerRowSpan}">${statusBarHtml}</td>`;
                    isFirstCustomerRow = false;
                }

                row.innerHTML = rowHtml;
                rightTableBody.appendChild(row);
            });
        }
    }
    // Add event listeners to the newly created cells
    document.querySelectorAll('#mrp-right-table .clickable-product').forEach(cell => {
        cell.addEventListener('click', handleMrpProductClick);
    });
}

function handleMrpProductClick(event) {
    const cell = event.target;
    const productCode = cell.dataset.productCode;

    if (!productCode || !allProducts[productCode]) {
        alert(`Product Code "${productCode}" has no purchasing history in the uploaded Procurement file.`);
        return;
    }

    // Toggle selection in the array
    const index = mrpSelectedForGraph.indexOf(productCode);
    if (index > -1) {
        mrpSelectedForGraph.splice(index, 1); // Deselect
    } else {
        mrpSelectedForGraph.push(productCode); // Select
    }

    // Update highlighting for all visible rows
    document.querySelectorAll('#mrp-right-table .clickable-product').forEach(c => {
        const code = c.dataset.productCode;
        const row = c.closest('tr');
        const productNameCell = row.cells[4];
        if (productNameCell) {
            productNameCell.classList.toggle('highlight-product-name', mrpSelectedForGraph.includes(code));
        }
    });

    // Re-render the details section with the full list of selected products
    renderMrpProductDetails(mrpSelectedForGraph);
}

function renderMrpProductDetails(productCodes) {
    const detailsSection = document.getElementById('mrp-details-section');
    const qtyCtx = document.getElementById('mrp-history-chart').getContext('2d');
    const priceCtx = document.getElementById('mrp-price-history-chart').getContext('2d');

    if (productCodes.length === 0) {
        detailsSection.classList.add('hidden');
        return;
    }
    detailsSection.classList.remove('hidden');

    if (mrpChartInstance) mrpChartInstance.destroy();
    if (mrpPriceChartInstance) mrpPriceChartInstance.destroy();

    const years = ['2020', '2021', '2022', '2023', '2024', '2025'];
    
    // Create datasets for all selected products
    const qtyDatasets = productCodes.map(code => {
        const product = allProducts[code];
        const data = years.map(year => product.years[year].qty);
        const color = `rgba(${Math.floor(Math.random() * 155) + 50}, ${Math.floor(Math.random() * 155) + 50}, ${Math.floor(Math.random() * 155) + 50}, 1)`;
        return { label: code, data: data, borderColor: color, backgroundColor: color.replace('1)', '0.2)'), fill: true, tension: 0.1, pointRadius: 5 };
    });

    const priceDatasets = productCodes.map(code => {
        const product = allProducts[code];
        const data = years.map(year => {
            const yearData = product.years[year];
            return yearData.count > 0 ? (yearData.priceSum / yearData.count) : null;
        });
        const color = qtyDatasets.find(d => d.label === code).borderColor;
        return { label: code, data: data, borderColor: color, backgroundColor: color.replace('1)', '0.2)'), fill: true, tension: 0.1, pointRadius: 5 };
    });

    // --- Quantity Chart ---
    mrpChartInstance = new Chart(qtyCtx, {
        type: 'line',
        data: { labels: years, datasets: qtyDatasets },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' }, title: { display: true, text: `Yearly Purchase Quantity History` }, datalabels: { align: 'top', anchor: 'end', backgroundColor: (context) => context.dataset.borderColor, borderRadius: 4, color: 'white', font: { size: 8, weight: 'bold' }, formatter: (value) => value > 0 ? value : '', padding: 4 } }, scales: { y: { beginAtZero: true, title: { display: true, text: 'Quantity' } } } }
    });

    // --- Price Chart ---
    mrpPriceChartInstance = new Chart(priceCtx, {
        type: 'line',
        data: { labels: years, datasets: priceDatasets },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'top' }, title: { display: true, text: `Yearly Average Price History` }, datalabels: { align: 'top', anchor: 'end', backgroundColor: (context) => context.dataset.borderColor, borderRadius: 4, color: 'white', font: { size: 8, weight: 'bold' }, formatter: (value) => value ? value.toFixed(2) : '', padding: 4 } }, scales: { y: { beginAtZero: false, title: { display: true, text: 'Average Unit Price' } } } }
    });
    
    // --- Details Table ---
    const tableBody = document.getElementById('mrp-details-table-body');
    tableBody.innerHTML = ''; 
    productCodes.forEach((code, index) => {
        const product = allProducts[code];
        const color = qtyDatasets.find(d => d.label === code).borderColor;
        let formattedDate = 'N/A';
        if (product.lastOrderDate) {
            const d = product.lastOrderDate;
            const month = d.toLocaleString('en-US', { month: 'short' });
            const year = d.getFullYear();
            formattedDate = `${month}-${year}`;
        }
        const newRow = document.createElement('tr');
        newRow.innerHTML = `
            <td>${index + 1}</td>
            <td style="background-color: ${color.replace('1)', '0.2)')}; color: ${color}; font-weight: bold;">${code}</td>
            <td>${product.name || 'N/A'}</td>
            <td>${product.vendor || 'N/A'}</td>
            <td class="text-center">${product.orderCount}</td>
            <td class="text-center">${product.total}</td>
            <td class="text-center bg-yellow-100 font-semibold">${product.latestUnitPrice.toFixed(2)}</td>
            <td class="text-center bg-yellow-100 font-semibold">${formattedDate}</td>
            <td class="text-center bg-yellow-100 font-semibold">${product.latestQuantity}</td>
        `;
        tableBody.appendChild(newRow);
    });
}

// --- New MRP Search Filter Functions ---
function populateMrpSearchFilters() {
    const uniqueCustomers = [...new Set(mrpRawDataForReport.map(row => row.Customer).filter(Boolean))].sort();
    const uniqueProductGroups = [...new Set(mrpRawDataForReport.map(row => row['Product Group']).filter(Boolean))].sort();
    
    populateCheckboxList('mrp-customer-list', uniqueCustomers, 'handleMrpCustomerSelection(this)');
    populateCheckboxList('mrp-pg-list', uniqueProductGroups, 'handleMrpPgSelection(this)');
}

function handleMrpCustomerSelection(checkbox) {
    handleMrpSearchSelection(checkbox, mrpSelectedCustomers);
}

function handleMrpPgSelection(checkbox) {
    handleMrpSearchSelection(checkbox, mrpSelectedProductGroups);
}

function handleMrpSearchSelection(checkbox, selectedArray) {
    const value = checkbox.value;
    if (checkbox.checked) {
        if (!selectedArray.includes(value)) selectedArray.push(value);
    } else {
        const index = selectedArray.indexOf(value);
        if (index > -1) selectedArray.splice(index, 1);
    }
    renderMrpRightPanelFromSearch();
}

function clearMrpSearchFilters() {
    mrpSelectedCustomers = [];
    mrpSelectedProductGroups = [];
    document.querySelectorAll('#mrp-customer-list input, #mrp-pg-list input').forEach(cb => cb.checked = false);
    renderMrpRightPanelFromSearch();
}

function renderMrpRightPanelFromSearch() {
    // Deselect any week row
    document.querySelectorAll('#mrp-left-table-body tr').forEach(r => r.classList.remove('selected'));
    document.getElementById('details-week-title').textContent = "Search Results";

    if (mrpSelectedCustomers.length === 0 && mrpSelectedProductGroups.length === 0) {
        document.getElementById('mrp-right-table-container').style.display = 'none';
        document.getElementById('mrp-details-section').classList.add('hidden');
        return;
    }

    let filteredData = mrpRawDataForReport;

    if (mrpSelectedCustomers.length > 0) {
        filteredData = filteredData.filter(row => mrpSelectedCustomers.includes(row.Customer));
    }
    if (mrpSelectedProductGroups.length > 0) {
        filteredData = filteredData.filter(row => mrpSelectedProductGroups.includes(row['Product Group']));
    }

    renderMrpDetailsTable(filteredData);
}