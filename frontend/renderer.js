/**
 * Financial Data Engine - Final Dynamic Renderer (English Version)
 * Emir Tasan - Warwick FinTech Project v1.0
 */

const API_BASE = 'http://127.0.0.1:5000/api';
let activeReport = 'report1'; 

window.onload = () => {
    document.getElementById('entitySelect').addEventListener('change', () => {
        if (activeReport) loadReport(activeReport);
    });

    const fileInput = document.getElementById('excelUpload');
    fileInput.addEventListener('change', async (event) => {
        const file = event.target.files[0];
        if (!file) return;

        updateStatus("Processing Data...", "bg-warning");
        document.getElementById('tableContainer').innerHTML = '<div class="text-center p-5"><div class="spinner-border text-primary"></div><p class="mt-3 text-muted">Initializing financial engine and analyzing dataset...</p></div>';

        try {
            const res = await fetch(`${API_BASE}/load_file`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ filepath: file.path })
            });

            // Improved error handling to catch the exact cause
            if (!res.ok) {
                const contentType = res.headers.get("content-type");
                if (contentType && contentType.indexOf("application/json") !== -1) {
                    const errorData = await res.json();
                    throw new Error(errorData.error || "Server error occurred.");
                } else {
                    throw new Error("Backend connection failed. Did you rebuild app.exe with PyInstaller?");
                }
            }

            const data = await res.json();

            // Unlock UI
            document.getElementById('entitySelect').disabled = false;
            document.querySelectorAll('.btn-report').forEach(btn => btn.disabled = false);

            await loadEntities();
            loadReport('report1');
            updateStatus("Engine Online", "bg-success");

        } catch (err) {
            showError(`File Upload Error: ${err.message}`);
            updateStatus("Engine Offline", "bg-danger");
        }
    });
};

async function loadEntities() {
    try {
        const res = await fetch(`${API_BASE}/entities`);
        const data = await res.json();
        const select = document.getElementById('entitySelect');
        select.innerHTML = data.map(e => `<option value="${e}">${e}</option>`).join('');
    } catch (e) { 
        showError("An error occurred while loading entities."); 
    }
}

async function loadReport(reportId) {
    activeReport = reportId;
    const entity = document.getElementById('entitySelect').value;
    const container = document.getElementById('tableContainer');
    
    updateSidebarUI(reportId);
    container.innerHTML = '<div class="text-center p-5"><div class="spinner-border text-primary"></div></div>';

    try {
        let url = `${API_BASE}/${reportId}?entity=${encodeURIComponent(entity)}`;
        if (reportId === 'report2') url += `&period=Full Year Monthly`;

        const res = await fetch(url);
        const data = await res.json();
        
        if (data.error) throw new Error(data.error);

        document.getElementById('reportTitle').innerText = formatReportTitle(reportId, entity);
        renderTable(data);
    } catch (err) {
        showError(`Report Error: ${err.message}`);
    }
}

function updateSidebarUI(reportId) {
    document.querySelectorAll('.btn-report').forEach(btn => btn.classList.remove('active-report'));
    const activeBtn = document.querySelector(`button[onclick="loadReport('${reportId}')"]`);
    if (activeBtn) activeBtn.classList.add('active-report');
}

function renderTable(data) {
    const container = document.getElementById('tableContainer');
    if (!data || data.length === 0) return;

    const cols = Object.keys(data[0]).filter(c => !['Is_Header', 'Is_Section', 'Level'].includes(c));
    let html = '<table class="table table-hover table-sm border"><thead><tr>';
    cols.forEach(c => html += `<th>${c.replace(/_/g, ' ')}</th>`);
    html += '</tr></thead><tbody>';

    const colorReports = ['report1', 'report2', 'report4'];

    data.forEach(row => {
        const isBoldRow = row.Is_Header || row.Level === 'Region' || row.Level === 'Total' || row.Is_Section;
        html += `<tr class="${isBoldRow ? 'table-secondary fw-bold' : ''}">`;
        
        const rowLabel = (row['Label'] || row['Metric'] || "").toLowerCase();
        
        const isPercentRow = /%|rate|turn|margin/i.test(rowLabel);
        const isExpenseRow = /cogs|sg&a|allocations|depreciation|expenses/i.test(rowLabel);
        const isProfitOrRevRow = /revenue|profit|ebitda|ebit/i.test(rowLabel) && !isPercentRow;
        const isCurrencyMetricRow = /(revenue|expenses|arpu|cogs|ebitda|profit|sales)/i.test(rowLabel) && !isPercentRow && !/days|hours/i.test(rowLabel);

        cols.forEach(c => {
            const colName = c.toLowerCase();
            let val = row[c];
            let cellClass = "";
            let formattedVal = val;

            if (typeof val === 'number') {
                const isRankCol = /rank/i.test(colName);
                const isVarianceCol = /yoy|mom|growth|variance|\$/i.test(colName) && !isRankCol;
                const isPercentCol = (/%|growth|yoy|margin|rate|turn/i.test(colName) || isPercentRow) && !/\$/i.test(colName) && !isRankCol;

                if (val === 0) {
                    formattedVal = "—";
                    cellClass = "neutral-value";
                } else {
                    if (isRankCol) {
                        formattedVal = val.toFixed(0); 
                    } else if (isPercentCol && colName !== 'label' && colName !== 'metric' && colName !== 'branch') {
                        formattedVal = (val * 100).toFixed(1) + "%"; 
                    } else {
                        const numString = val.toLocaleString('en-US', { 
                            minimumFractionDigits: (Math.abs(val) < 1000 && !Number.isInteger(val) ? 2 : 0),
                            maximumFractionDigits: 2 
                        });

                        let addDollar = false;
                        if (activeReport === 'report5') {
                            if (isCurrencyMetricRow && colName !== 'label' && colName !== 'metric') addDollar = true;
                        } else {
                            if (/\$|rev|cogs|sga|ebit|ebd|nov|dec|fy|alc|m\d|val|arpu|sales/i.test(colName) && !/unit|order|hc|headcount/i.test(colName) && !isPercentCol && !isRankCol) {
                                addDollar = true;
                            }
                        }
                        formattedVal = addDollar ? "$" + numString : numString;
                    }

                    if (isRankCol) {
                        cellClass = "neutral-value";
                    } else if (colorReports.includes(activeReport)) {
                        if (isVarianceCol) {
                            if (val < 0) cellClass = "neg-value";
                            else if (val > 0) cellClass = "pos-value";
                        } else {
                            if (isExpenseRow) cellClass = "neg-value"; 
                            else if (isProfitOrRevRow) cellClass = "pos-value"; 
                            else cellClass = "neutral-value"; 
                        }
                    } else {
                        cellClass = "neutral-value"; 
                    }
                }
                html += `<td class="${cellClass}">${formattedVal}</td>`;
            } else {
                html += `<td>${formattedVal || '—'}</td>`; 
            }
        });
        html += '</tr>';
    });
    container.innerHTML = html + '</tbody></table>';
}

function formatReportTitle(id, entity) {
    const names = { 
        report1: 'Income Statement (Comparison)', 
        report2: 'Monthly IS Trend Analysis', 
        report3: 'Branch Performance Rankings', 
        report4: 'Regional Performance Comparison', 
        report5: 'Operational Metrics Dashboard' 
    };
    return `${names[id]} - ${entity}`;
}

function showError(msg) {
    document.getElementById('tableContainer').innerHTML = `<div class="alert alert-danger m-3 shadow"><strong>Error:</strong> ${msg}</div>`;
}

function updateStatus(text, bgClass) {
    const el = document.getElementById('statusIndicator');
    if (el) el.innerHTML = `<span class="badge ${bgClass}">● ${text}</span>`;
}