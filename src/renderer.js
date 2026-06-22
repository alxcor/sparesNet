import jspreadsheet from 'jspreadsheet-ce';
import 'jspreadsheet-ce/dist/jspreadsheet.css';
import './style.css';

// Only for offline tests
//import path from 'path';
//import { fileURLToPath } from 'url';

// Variables
let table; 
let selectedRow = 0;
let selectedCol = 0;
let wizardStep = 0; // 0 = Idle, 1 = Reading All, 2 = Reading Successors

const plmLifecycle = {
    "sales release": "PM280",
    "delivery release": "PM300",
    "phase out announce": "PM400",
    "prod. cancellation": "PM410",
    "prod. discont.": "PM490",
    "end prod.lifecycl.": "PM500"
};

// Helper: Get Cell name
function getCellName(row, col) {
    let letter = '';
    let tempCol = col;
    while (tempCol >= 0) {
        letter = String.fromCharCode((tempCol % 26) + 65) + letter;
        tempCol = Math.floor(tempCol / 26) - 1;
    }
    return letter + (row + 1);
}

// Helper: Clear Row data (keep 1st column)
function clearRow(rowNo) {
    const colNo = table.headers.length;
    let data = '';
    // Save "successor for" data, if available
    if (colNo >= 29) {
        const rawData = table.getValueFromCoords(29, rowNo);
        data = rawData ? rawData.toString().trim() : '';
    }
    // Clear columns except index 0
    for (let i = 1; i < colNo; i++) {
        table.setValueFromCoords(i, rowNo, ' ');
    }
    // Restore the "successor for" data
    if (colNo >= 29) {
        table.setValueFromCoords(29, rowNo, data);
    }
    clearRowStyle(rowNo);
}

// Helper: Clear Row Style
function clearRowStyle(rowNo) {
    if (!table) return;
    const colNo = table.headers.length;
    const styleResetBatch = {};
    // Loop through every column in the target row
    for (let col = 0; col < colNo; col++) {
        const cellName = getCellName(rowNo, col);
        // Setting cell style to empty string (delete all style)
        styleResetBatch[cellName] = '';
    }
    // Apply the entire reset batch at once (perf.)
    table.setStyle(styleResetBatch);
}

// Helper: Format Cells in Report According to PMD
function getStatusStyle(col, statusValue) {
    if (!col) return '';
    if (!statusValue) return '';
    const text = statusValue.toString().toLowerCase().trim();
    // Type 1: Delivery Status:
    if ((text.length > 0) && (col == 1)) {
        if (text.includes('release') || text.includes('active')) {
            return 'pm300bkgd';
        }
        if (text.includes('phase out') || text.includes('announce')) {
            return 'pm410bkgd';
        }
        if (text.includes('cancel')) {
            return 'pm410bkgd';
        }
        if (text.includes('discont') || text.includes('end')) {
            return 'pm500bkgd';
        }
    }
    // Type 2: Note with or wo successor
    if ((text.length > 0) && (col == 2)) {
        if (text.includes('succ.') || text.includes('successor') || text.includes('replace')) {
            return 'pmsucbkgd';
        } else {
            return 'pmerrbkgd';
        }
    }
    // Type 2: No Notes
    if ((text.length <= 0) && (col == 2)) {
        return ' ';
    }
    // Type 3: Successor found
    if (col == 3) {
        if (text.length > 0) {
            return 'pmsucbkgd';
        } else {
            return 'pm500bkgd';
        }
    }
    // Default fallback (no styling)
    return 'pmerrbkgd';
}

// Helper: Milestone Texts to Milestone Codes
function replaceDateLabels(input) {
    if (!input) return input;
    let output = input.toString().toLowerCase();
    for (const [original, short] of Object.entries(plmLifecycle)) {
        const escaped = original.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const regex = new RegExp(escaped, "gi");
        output = output.replace(regex, short);
    }
    output = output.replace(/(\r\n|\r|\n)/g, "<br>");
    return output;
}

// Helper: normalize PM Codes
function normalizePMCode(statusValue) {
    if (!statusValue) return 'UNKNOWN';
    const text = statusValue.toString().toLowerCase().trim();
    if (text.includes('sales release')) return 'PM280';
    if (text.includes('delivery release') || text.includes('release') || text.includes('active')) return 'PM300';
    if (text.includes('phase out') || text.includes('announce')) return 'PM400';
    if (text.includes('cancel')) return 'PM410';
    if (text.includes('discont')) return 'PM490';
    if (text.includes('end')) return 'PM500';
    return 'UNKNOWN';
}

// Helper: Extract PLM Milestone Dates
function parseMilestoneDates(rawText) {
    const dataMap = {};
    if (!rawText) return dataMap;
    // Split the input text into individual lines
    const lines = rawText.toString().replace(/_x000d_/gi, '\n').split(/\r?\n/);
    lines.forEach (line => {
        const trimmed = line.trim();
        if (!trimmed) return; // Skip empty lines
        let stageLabel = trimmed;
        let dateString = null;
        if (trimmed.includes(':')) {
            const parts = trimmed.split(':');
            stageLabel = parts[0].trim();
            dateString = parts[1].trim();
        }
        // Standardize the key name based on our dictionary lookup
        const lookupKey = stageLabel.toLowerCase();
        if (plmLifecycle[lookupKey]) {
            const pmCode = plmLifecycle[lookupKey];
            if (dateString) {
                let parsedDate = null;
                const dStr = dateString.trim();
                // Handle DD.MM.YYYY
                if (/^\d{2}\.\d{2}\.\d{4}$/.test(dStr)) {
                    const [d, m, y] = dStr.split('.').map(Number);
                    parsedDate = new Date(y, m - 1, d);
                } 
                // Handle YYYY-MM-DD
                else if (/^\d{4}-\d{2}-\d{2}$/.test(dStr)) {
                    const [y, m, d] = dStr.split('-').map(Number);
                    parsedDate = new Date(y, m - 1, d);
                }
                // Fallback to native parsing
                else {
                    parsedDate = new Date(dStr);
                }
                dataMap[pmCode] = (parsedDate && !isNaN(parsedDate.getTime())) ? parsedDate : null;
            } else {
                dataMap[pmCode] = null;
            }
        }
    });
    console.log(dataMap);
    return dataMap;
}

// Helper: finalize the wizard workflow
function triggerWizardReport() {
    wizardStep = 0; // Reset wizard state back to idle
    let reportHTML = generateReport();
    if (reportHTML) {
        window.api.previewReport(reportHTML); // Show the final report preview
    }
}

// Helper: Move focus to selected cell
function selectAndFocusCell(colIndex, rowIndex) {
    if (!table) return;
    // Apply the selection highlight box around the cell
    table.updateSelectionFromCoords(colIndex, rowIndex, colIndex, rowIndex);
    // Fetch the cell's HTML element and scroll it smoothly into view
    const cellElement = table.getCellFromCoords(colIndex, rowIndex);
    if (cellElement) {
        cellElement.scrollIntoView({
            behavior: 'smooth', // 'auto' or 'smooth'
            block: 'nearest',   // Vertically aligns to screen edge if hidden
            inline: 'nearest'   // Horizontally aligns to screen edge if hidden
        });
    }
}

// Init -----------------------------------------
function select(instance, x1, y1, x2, y2) {
    // x1, y1 is the starting cell of the selection
    // x2, y2 is the ending cell (relevant if a range is dragged)
    selectedCol = x1; 
    selectedRow = y1;
    const cellElement = instance.getCellFromCoords(x1, y1);
    if (cellElement) {
        cellElement.scrollIntoView({
            behavior: 'auto',
            block: 'nearest',
            inline: 'nearest'
        });
    }
    const text = document.getElementById("row_number");
    text.textContent = `Row: ${selectedRow + 1}`;
}

// Initialize the spreadsheet
function initTable() {
    const container = document.getElementById('table_container');
    // Clear the container to prevent duplicate tables during Vite hot-reloads
    container.innerHTML = '';
    const options = {
        worksheets: [{
            data: [[]],
            minDimensions: [30, 500],
            tableOverflow: true,
            tableWidth: '100%',
            defaultColAlign: 'left',
            columns: [
                { type: 'text', title: 'Your Data...', width: 200, wordWrap: true },                    // 0
                { type: 'text', title: 'MLFB'        , width: 200, wordWrap: true },                    // 1
                { type: 'text', title: 'Product Description', width: 300, wordWrap: true },             // 2
                { type: 'text', title: 'Product Family', width: 150, wordWrap: true },                  // 3
                { type: 'text', title: 'Product Lifecycle (PLM)', width: 150, wordWrap: true },         // 4
                { type: 'text', title: 'PLM Effective Date', width: 300, wordWrap: true },              // 5
                { type: 'text', title: 'Notes', width: 150, wordWrap: true },                           // 6
                { type: 'text', title: 'Price Group', width: 150, wordWrap: true },                     // 7
                { type: 'text', title: 'Surcharge for Raw Materials', width: 150, wordWrap: true },     // 8
                { type: 'text', title: 'Metal Factor', width: 150, wordWrap: true },                    // 9
                { type: 'text', title: 'Export Control Regulations', width: 150, wordWrap: true },      //10
                { type: 'text', title: 'Dispatch Time', width: 150, wordWrap: true },                   //11
                { type: 'text', title: 'Net Weight (kg)', width: 150, wordWrap: true },                 //12
                { type: 'text', title: 'Product Dimensions (W x L x H)', width: 150, wordWrap: true },  //13
                { type: 'text', title: 'Packaging Dimension', width: 150, wordWrap: true },             //14
                { type: 'text', title: 'Minimum Order Quantity', width: 150, wordWrap: true },          //15
                { type: 'text', title: 'Quantity Unit', width: 150, wordWrap: true },                   //16
                { type: 'text', title: 'Packaging Quantity', width: 150, wordWrap: true },              //17
                { type: 'text', title: 'EAN', width: 150, wordWrap: true },                             //18
                { type: 'text', title: 'UPC', width: 150, wordWrap: true },                             //19
                { type: 'text', title: 'Commodity Code', width: 150, wordWrap: true },                  //20
                { type: 'text', title: 'KZ_FDB/ CatalogID', width: 150, wordWrap: true },               //21
                { type: 'text', title: 'Product Group', width: 150, wordWrap: true },                   //22
                { type: 'text', title: 'Country of origin', width: 150, wordWrap: true },               //23
                { type: 'text', title: 'Compliance ... to RoHS directive', width: 150, wordWrap: true },//24
                { type: 'text', title: 'Product class', width: 150, wordWrap: true },                   //25
                { type: 'text', title: 'Obligation Cat. for taking back', width: 150, wordWrap: true }, //26
                { type: 'text', title: 'Classifications', width: 150, wordWrap: true },                 //27
                { type: 'text', title: 'Successor', width: 150, wordWrap: true },                       //28
                { type: 'text', title: 'Successor for', width: 150, wordWrap: true },                   //29
            ],
        }],
        onselection: select,
        onchange: function(instance, cell, col, row, value) {
            // Reset existing styles first
            cell.style.backgroundColor = '';
            cell.style.color = '';
            cell.style.fontWeight = '';
            const text = (value || ' ').toLowerCase().trim();
            if (col == 1) { // Col.1 MLFB
                if (text.includes('error') || text.includes('err.')) {
                    cell.style.backgroundColor = '#F28794'; //Red
                    cell.style.color = '#000000';
                }
            }
            if (col == 4) { // Col.4 PLM Status
                const pmCode = normalizePMCode(value);
                if (pmCode === 'PM280' || pmCode === 'PM300') {
                    cell.style.backgroundColor = '#90E39A';
                } else if (pmCode === 'PM400') {
                    cell.style.backgroundColor = '#FFD6A3';
                } else if (pmCode === 'PM410') {
                    cell.style.backgroundColor = '#FFCE92';
                } else if (pmCode === 'PM490' || pmCode === 'PM500') {
                    cell.style.backgroundColor = '#F28794';
                }
            }
            if (col == 6) { // Col.6 Notes
                if (text.length > 0) {
                    if (text.includes('succ.') || text.includes('successor') || text.includes('replace')) {
                        cell.style.backgroundColor = '#5B9BD5'; //Blue
                    } else {
                        cell.style.backgroundColor = '#9EC3E6'; //Light Blue
                    }
                }
            }
        }
    };
    function resizeTable() {
        const container = document.getElementById('table_container');
        // Finds either a <footer class="footer"> or any footer tag
        const footer = document.querySelector('.footer') || document.querySelector('footer'); 
        if (!container || !table) return;
        // Find table container start from the top of the viewport
        const topOffset = container.getBoundingClientRect().top;
        // Measure the exact height of the footer
        const footerHeight = footer ? footer.offsetHeight : 40;
        // Subtract top layout and footer (and 25px extra) from viewport height
        const dynamicHeight = window.innerHeight - topOffset - footerHeight - 25;
        if (dynamicHeight > 100) {
            // Apply the calculated height to the container element
            container.style.height = `${dynamicHeight}px`;
            // Push the brand new width and height bounds into Jspreadsheet's core engine
            table.setViewport(container.offsetWidth, dynamicHeight);
        }
    }
    // Initialize and store the first worksheet in our 'table' variable
    const spreadsheet = jspreadsheet(container, options);
    table = spreadsheet[0];
    // Recalculate dimensions on window resize
    window.addEventListener('resize', resizeTable);
    // Run immediately to fit perfectly right after the app loads
    setTimeout(resizeTable, 100);
}

// Events, Commands -----------------------------

// Event: Window UI is ready
window.addEventListener('DOMContentLoaded',() => {
    initTable();
    const text = document.getElementById("row_number");
    text.textContent = `Row: ${selectedRow + 1}`;
})

// Buttons: Dark / Light Theme
document.getElementById('bt_dark_mode').addEventListener('click',async () => {
    const isDarkMode = await window.darkMode.toggle()
})

// Buttons: Clear Table
document.getElementById("clear_all").onclick = () => {
    clearAll();
};

// Buttons: Clear Data except column 1
document.getElementById("clear_row").onclick = () => {
    clearAllRows();
};

// Buttons: Import data from Excel file
document.getElementById('import_xls').addEventListener('click', async () => {
    try {
        const filePath = await window.excelAPI.selectOpenXlsFile();
        if (!filePath) return;
        const getHeaderR = true;
        const getAllData = true;
        let rawData;
        const options = { skipHeader: !getHeaderR };
        if (getAllData) {
            rawData = await window.excelAPI.importXlsData(filePath, options);
        } else {
            rawData = await window.excelAPI.importXlsFirstColumn(filePath, options);
        }
        let rows = [];
        // Get internal headers
        const expectedHeaders = table.options.columns.map(col => (col.title || '').toLowerCase().trim());
        if (getAllData) {
            const allRows = rawData.data || rawData.rows || rawData || [];
            // Clean out empty rows
            let cleanRows = allRows.filter(row => row && row.length > 0);
            // Check first row
            if (cleanRows.length > 0) {
                const firstRowNormalized = cleanRows[0].map(cell => cell ? cell.toString().toLowerCase().trim() : '');
                // See if at least one critical header matches (e.g., 'mlfb' or 'product description')
                const isHeaderRow = firstRowNormalized.includes('mlfb') || 
                                    firstRowNormalized.includes('product description') ||
                                    firstRowNormalized.every((val, i) => val === expectedHeaders[i]);
                if (isHeaderRow) {
                    cleanRows.shift(); // Remove the matching header row from your dataset
                }
            }
            rows = cleanRows;
        } else {
            // First column data processing
            const firstCol = rawData.firstCol || rawData || [];
            let nonNull = firstCol.filter(v => v !== null && v !== undefined && v !== '');
           
            // If the first item matches the first column's expected header, skip it
            if (nonNull.length > 0 && nonNull[0].toString().toLowerCase().trim() === expectedHeaders[0]) {
                nonNull.shift();
            }
            rows = nonNull.map(v => [v]);
        }
        // If table exists and supports setData, replace its data
        if (table && typeof table.setData === 'function') {
            // Optionally clear first
            table.setData([]); 
            table.setData(rows);
            let cell;
            let value;
            rows.forEach((row, r) => {
                const onChangeHandler = table.parent?.config?.onchange;
                if (typeof onChangeHandler === 'function') {
                    [1, 4, 6].forEach(colIdx => {
                        cell = table.getCellFromCoords(colIdx, r);
                        value = table.getValueFromCoords(colIdx, r);
                        if (cell) {
                            onChangeHandler(table, cell, colIdx, r, value);
                        }
                    });
                }
            });
        }
    } catch (err) {
        //console.error('Error reading Excel file:', err);
    }
});

// Buttons: Export data to an Excel file
document.getElementById('export_xls').addEventListener('click', async () => {
    try {
        if (!table) return;
        const filePath = await window.excelAPI.selectSaveXlsFile();
        if (!filePath) return; // User cancelled the save operation
        // Grab the column header titles
        const headers = table.options.columns.map(col => col.title || '');
        // Extract all data rows
        const rows = table.getData();
        // Merge headers and row arrays
        const exportData = [headers, ...rows];
        // Send everything to preload script
        await window.excelAPI.exportXlsData(filePath, exportData);
        //alert(`Data successfully exported to:\n${filePath}`);
    } catch (err) {
        //console.error('Error exporting Excel file:', err);
        //alert('An error occurred while attempting to export the Excel file.');
    }
});

// Buttons: Wizzard, Sort by MLFB, Read All, Check Successors, Generate Report
document.getElementById("report_wiz").onclick = () => {
    wizardStep = 1;  // Move to Step 1
    sortByColumn(0); // Sort rows
    readAll();       // Download web data for all rows
};

// Buttons: Sort by code
document.getElementById("sort_cod").onclick = () => {
    sortByColumn(0);
};

// Buttons: Sort by column
document.getElementById("sort_col").onclick = () => {
    sortByColumn(selectedCol);
};

// Buttons: Read data for current row
document.getElementById('read_row').onclick = () => {
    readRow();
};

// Buttons: Read data for all rows
document.getElementById("read_all").onclick = () => {
    readAll();
};

// Buttons: Add row for successor and read data
document.getElementById('succ_row').onclick = () => {
    checkSuccessor();
};

// Buttons: Add rows for successors and read data
document.getElementById('succ_all').onclick = () => {
    checkSuccessorAll();
};

// Buttons: Set report
document.getElementById('report_set').onclick = () => {
    let reportHTML = generateReport();
    if (!reportHTML) return;
    window.api.previewReport(reportHTML);
};

// Buttons: Set report and save to HTML
document.getElementById('report_htm').onclick = () => {
    let reportHTML = generateReport();
    if (!reportHTML) return;
    window.api.selectSaveHtmFile(reportHTML);
};

// Buttons: Set report and save to PDF
document.getElementById('report_pdf').onclick = () => {
    let reportHTML = generateReport();
    if (!reportHTML) return;
    window.api.selectSavePdfFile(reportHTML);
};

// Buttons: Open Spares on Web for Row
document.getElementById('int_sow').onclick = () => {
    const rawData = table.getValueFromCoords(0, selectedRow);
    const data = rawData ? rawData.toString().trim() : '';
    //const data = table.getValueFromCoords(0, selectedRow);
    if(data.length > 1){
        const siteBase = 'https://www.sow.siemens.com/?mask=single&mlfb='
        const siteEnds = '&getImageUrls=true';
        const siteName = siteBase + data + siteEnds;
        window.api.openExternal(siteName);
    }
};

// Buttons: Open Industry Mall on Web for Row
document.getElementById('int_mall').onclick = () => {
    const rawData = table.getValueFromCoords(0, selectedRow);
    const data = rawData ? rawData.toString().trim() : '';
    if(data.length > 1){
        const siteBase = 'https://sieportal.siemens.com/en-ww/products-services/detail/'
        const siteEnds = '?tree=CatalogTree';
        const siteName = siteBase + data + siteEnds;
        window.api.openExternal(siteName);
    }
};

// Buttons: Open SIOS on Web for Row
document.getElementById('int_sios').onclick = () => {
    const rawData = table.getValueFromCoords(0, selectedRow);
    const data = rawData ? rawData.toString().trim() : '';
    if(data.length > 1){
        const siteBase = 'https://sieportal.siemens.com/en-ww/products-services/detail/'
        const siteEnds = '?tree=CatalogTree';
        const siteName = siteBase + data + siteEnds;
        window.api.openExternal(siteName);
    }
};

// Buttons: Open Web Page Help
document.getElementById('web_web').onclick = () => {
    window.api.openExternal('https://alxcor.github.io/sparesnet/');
};

// Buttons: Open Github for Project
document.getElementById('web_github').onclick = () => {
    window.api.openExternal('https://github.com/alxcor/sparesnet/');
};

// Receive Progress from (sub-) browser window
window.api.onReadProgress((data) => {
    const text = document.getElementById("progress_text");
    const total = data.total > 0 ? data.total : 1;
    if (data.status != "Ready ") {
        text.textContent = `${data.index}/${data.total}: ${data.status} ${data.message}`;
    } else {
        text.textContent = `${data.status} ${data.message}`;
        if (wizardStep === 1) {
            // Step 1 (Read All) Complete. Now check successors.
            const launchedSuccessorRead = checkSuccessorAll();
            if (launchedSuccessorRead) {
                wizardStep = 2; // New successor rows found
            } else {
                triggerWizardReport(); // No new successors found, go to report
            }
        } else if (wizardStep === 2) {
            // Step 2 (check Successor) Complete, now generate the report.
            triggerWizardReport();
        }
    }
    const progress = document.getElementById("progress_bar");
    if (progress) {
        progress.value = (data.index / total) * 100;
    }
});

// Receive Log_Data from (sub-) browser window
window.api.onReadLog_Data((data) => {
    console.log(data);
});

// Receive Data from (sub-) browser window
window.api.onReadReceived((data) => {
    // Append to a list in main page
    if (table) {
        if(data.index >= 0){
            clearRow(data.index);
            table.setRowData(data.index, []);
            clearRowStyle(data.index);
            table.setValueFromCoords(1, data.index, data.mlfb_ || '');
            table.setValueFromCoords(2, data.index, data.descr || '');
            table.setValueFromCoords(3, data.index, data.famil || '');
            table.setValueFromCoords(4, data.index, data.plm_s || '');
            table.setValueFromCoords(5, data.index, data.plm_d || '');
            table.setValueFromCoords(6, data.index, data.notes || '');
            table.setValueFromCoords(7, data.index, data.pri_g || '');
            table.setValueFromCoords(8, data.index, data.raw_m || '');
            table.setValueFromCoords(9, data.index, data.met_f || '');
            table.setValueFromCoords(10, data.index, data.exp_c || '');
            table.setValueFromCoords(11, data.index, data.dispt || '');
            table.setValueFromCoords(12, data.index, data.net_w || '');
            table.setValueFromCoords(13, data.index, data.dim_n || '');
            table.setValueFromCoords(14, data.index, data.dim_p || '');
            table.setValueFromCoords(15, data.index, data.unitp || '');
            table.setValueFromCoords(16, data.index, data.unitq || '');
            table.setValueFromCoords(17, data.index, data.packq || '');
            table.setValueFromCoords(18, data.index, data.ean__ || '');
            table.setValueFromCoords(19, data.index, data.upc__ || '');
            table.setValueFromCoords(20, data.index, data.comcd || '');
            table.setValueFromCoords(21, data.index, data.catal || '');
            table.setValueFromCoords(22, data.index, data.prodg || '');
            table.setValueFromCoords(23, data.index, data.orign || '');
            table.setValueFromCoords(24, data.index, data.rohs_ || '');
            table.setValueFromCoords(25, data.index, data.pclas || '');
            table.setValueFromCoords(26, data.index, data.takeb || '');
            table.setValueFromCoords(27, data.index, data.clsif || '');
            table.setValueFromCoords(28, data.index, data.succe || '');
        }
    }
});

// Functions ------------------------------------

// Clear Table Data
function clearAll() {
    if (table) {
        table.setData([[]]); // Reset to empty
    }
};

// Clear Table Data, keep first column
function clearAllRows() {
    const rowNo = table.options.data.length;
    for (let i = 0; i < rowNo; i++) {
        clearRow(i);
    }
};

// Sort by Column
function sortByColumn(column) {
    if (column === undefined || column === null) return;
    if (table) {
        table.orderBy(column);
    }
};

// Read data for current Row
function readRow(){
    const rawData = table.getValueFromCoords(0, selectedRow);
    const siteName = rawData ? rawData.toString().trim() : '';
    clearRow(selectedRow);
    clearRowStyle(selectedRow);
    if(siteName.length > 1){
        const siteBase = 'https://sieportal.siemens.com/en-ww/products-services/detail/'
        const siteEnds = '?tree=CatalogTree';
        const data = [];
        data.push([selectedRow, siteBase + siteName + siteEnds]);
        window.api.startReadPages(data);
    }
};

// Test: Read data from local file for current Row
function readRowLocal(){
    clearRow(selectedRow)
    clearRowStyle(selectedRow);
    const filePath2 = `C:\\Work\\6ES7138-4FA02-0AB0 - Siemens SiePortal.html`;
    const data2 = [];
    data2.push([selectedRow, filePath2]);
    window.api.startReadPages(data2);
};

// Read data for all Rows
function readAll(){
    let i = 0;
    const data = [];
    const rowNo = table.options.data.length;
    for (i = 0; i < rowNo; i++) {
        const siteName = table.getValueFromCoords(0, i);
        if(siteName.length > 1){
            const siteBase = 'https://sieportal.siemens.com/en-ww/products-services/detail/'
            const siteEnds = '?tree=CatalogTree';
            data.push([i, siteBase + siteName + siteEnds]);
        }
    }
    window.api.startReadPages(data);
};

// Add row for successor and read data
function checkSuccessor(){
    // Get successor data, if any
    const rawData = table.getValueFromCoords(28, selectedRow);
    const successorData = rawData ? rawData.toString().trim() : '';
    // Check if a successor was found
    if (successorData.length <= 1) return;
    // Save original part number (from Column 0)
    const rawOriginal = table.getValueFromCoords(0, selectedRow);
    const originalData = rawOriginal ? rawOriginal.toString().trim() : '';
    const totalRows = table.options.data.length;
    // Scan all rows to see if this successor already exists anywhere in column 0
    for (let i = 0; i < totalRows; i++) {
        const cellRawData = table.getValueFromCoords(0, i);
        const cellValue = cellRawData ? cellRawData.toString().trim() : '';
        if (cellValue === successorData) return;    // Successor found in the table
    }
    const newRowIndex = selectedRow + 1;
    table.insertRow(1, newRowIndex, true);
    // Clear the new row's styling/data
    clearRow(newRowIndex);
    clearRowStyle(newRowIndex);
    // Write "my data": Copy old row column 28 to new row column 1
    table.setValueFromCoords(0, newRowIndex, successorData);
    // Write "succ.for": Copy old row column 1 to new row column 29
    table.setValueFromCoords(29, newRowIndex, originalData);
    // Change selection focus to the newly created row
    selectedRow = newRowIndex;
    selectedCol = 0;
    selectAndFocusCell(selectedCol, selectedRow);
    // Call readRow() for the new row
    readRow();
};

// Add rows for successors and read data
function checkSuccessorAll(){
    if (!table) return;
    const batchData = [];
    let insertedCount = 0;
    // Cache the original length so we don't loop endlessly into newly inserted rows
    const originalTotalRows = table.options.data.length; 
    // Loop forward from the first row up to the last original row
    for (let i = 0; i < originalTotalRows; i++) {
        // Calculate the current index of the original row since previous insertions shifted it down
        const currentIndex = i + insertedCount;
        // Get the successor part number for row 'currentIndex'
        const rawSuccessor = table.getValueFromCoords(28, currentIndex);
        const successorData = rawSuccessor ? rawSuccessor.toString().trim() : '';
        // If there is no valid successor, skip this row
        if (successorData.length <= 1) continue;
        // Get the original part number from Column 0
        const rawOriginal = table.getValueFromCoords(0, currentIndex);
        const originalData = rawOriginal ? rawOriginal.toString().trim() : '';
        // Scan the entire table to check if this successor already exists in Column 0
        let alreadyExists = false;
        const currentTotalRows = table.options.data.length; 
        for (let j = 0; j < currentTotalRows; j++) {
            const cellRawData = table.getValueFromCoords(0, j);
            const cellValue = cellRawData ? cellRawData.toString().trim() : '';
            if (cellValue === successorData) {
                alreadyExists = true;
                break; // Found it, stop searching
            }
        }
        // If it already exists somewhere in the table, skip inserting a duplicate
        if (alreadyExists) continue;
        // Insert the new row directly below the current item
        const newRowIndex = currentIndex + 1;
        table.insertRow(1, newRowIndex, true);
        // Clear data and styling for the newly created row safely
        clearRow(newRowIndex);
        clearRowStyle(newRowIndex);
        // Populate the new row placeholder
        table.setValueFromCoords(0, newRowIndex, successorData); // Successor goes to Col 0
        table.setValueFromCoords(29, newRowIndex, originalData);  // Original goes to Col 29
        // Increment offset counter so future iterations look at the correct shifted rows
        insertedCount++;
        // Collect the row index and URL into our batch array
        const siteBase = 'https://sieportal.siemens.com/en-ww/products-services/detail/';
        const siteEnds = '?tree=CatalogTree';
        batchData.push([newRowIndex, siteBase + successorData + siteEnds]);
    }
    // Trigger a single, unified API read for all newly added rows (identical approach to readAll)
    if (batchData.length > 0) {
        window.api.startReadPages(batchData);
        return true; // An async read was launched
    }
    return false;    // No new successors need fetching
};

// Generate HTML Report
function generateReport(){
    const data = table.getData();
    const filteredData = data.filter(row => {
        return row[0] && row[0].toString().trim() !== "";
    });
    // Count variables initialization
    let countPM280 = 0;
    let countPM300 = 0;
    let countPM400 = 0;
    let countPM410 = 0;
    let countPM490 = 0;
    let countPM500 = 0;
    let countPMact = 0;
    let countPMcan = 0;
    let countPMdis = 0;
    let countPMerr = 0;
    let countPMall = 0;
    // Highlights
    let redWithSuccessor = 0;
    let redWithoutSuccessor = 0;
    let orangeParts = 0;
    // Age (in months) per milestone stage
    const pmAges = {
        PM280: [],
        PM300: [],
        PM400: [],
        PM410: [],
        PM490: [],
        PM500: []
    };
    // Scan column 4 (Product Lifecycle PLM status) across active rows
    filteredData.forEach(row => {
        countPMall++;
        const currentPmCode = normalizePMCode(row[4]);
        const successorText = row[28] ? row[28].toString().trim() : '';
        // Calculate Standard PM Stage Categories
        if (currentPmCode === 'PM280') {
            countPM280++; countPMact++;
        } else if (currentPmCode === 'PM300') {
            countPM300++; countPMact++;
        } else if (currentPmCode === 'PM400') {
            countPM400++; countPMcan++;
            orangeParts++;
        } else if (currentPmCode === 'PM410') {
            countPM410++; countPMcan++;
            orangeParts++;
        } else if (currentPmCode === 'PM490') {
            countPM490++; countPMdis++;
            if (successorText) redWithSuccessor++; else redWithoutSuccessor++;
        } else if (currentPmCode === 'PM500') {
            countPM500++; countPMdis++;
            if (successorText) redWithSuccessor++; else redWithoutSuccessor++;
        } else {
            countPMerr++;
        }
        // Track specific age statistics
        if (currentPmCode !== 'UNKNOWN' && row[5]) {
            const milestoneDates = parseMilestoneDates(row[5]);
            const stageDate = milestoneDates[currentPmCode];
            if (stageDate && !isNaN(stageDate.getTime())) {
                const now = new Date();
                // Precise difference in months
                const diffMonths = (now.getFullYear() - stageDate.getFullYear()) * 12 + (now.getMonth() - stageDate.getMonth());
                if (diffMonths >= 0) {
                    pmAges[currentPmCode].push(diffMonths);
                }
            }
        }
    });
    // Helper function to calculate Min, Avg, Max and format to Years/Months
    function getAgeStats(agesArray) {
        if (!agesArray || agesArray.length === 0) return { min: '-', avg: '-', max: '-' };
        const min = Math.min(...agesArray);
        const max = Math.max(...agesArray);
        const sum = agesArray.reduce((a, b) => a + b, 0);
        const avg = Math.round(sum / agesArray.length);
        const format = (mos) => {
            if (mos < 0) return '0m';
            const yrs = Math.floor(mos / 12);
            const remMos = mos % 12;
            if (yrs > 0) return `${yrs}y ${remMos}m`;
            return `${mos}m`;
        };
        return {
            min: format(min),
            avg: format(avg),
            max: format(max)
        };
    }
    // Generate stats for each lifecycle stage
    const statsPM280 = getAgeStats(pmAges.PM280);
    const statsPM300 = getAgeStats(pmAges.PM300);
    const statsPM400 = getAgeStats(pmAges.PM400);
    const statsPM410 = getAgeStats(pmAges.PM410);
    const statsPM490 = getAgeStats(pmAges.PM490);
    const statsPM500 = getAgeStats(pmAges.PM500);
    // Calculate maximum timeline metrics for PM410 parts (0-10 Years Scale = 120 Months total)
    const hasPM410Parts = pmAges.PM410.length > 0;
    const maxPM410Months = hasPM410Parts ? Math.max(...pmAges.PM410) : 0;
    const filledMonths = Math.min(maxPM410Months, 120);
    const emptyMonths = 120 - filledMonths;
    const maxPM410Yrs = Math.floor(maxPM410Months / 12);
    const maxPM410Mos = maxPM410Months % 12;
    const durationText = maxPM410Yrs > 0 ? `${maxPM410Yrs} Years ${maxPM410Mos} Months` : `${maxPM410Mos} Months`;
    //
    const activeCount = countPMall - redWithoutSuccessor - redWithSuccessor - orangeParts;
    const pctNoSuccessor = countPMall > 0 ? Math.round((redWithoutSuccessor / countPMall) * 100) : 0;
    const pctWithSuccessor = countPMall > 0 ? Math.round((redWithSuccessor / countPMall) * 100) : 0;
    const pctOrange      = countPMall > 0 ? Math.round((orangeParts / countPMall) * 100) : 0;
    const pctActive      = countPMall > 0 ? Math.round((activeCount / countPMall) * 100) : 0;
    //
    let reportHTML = `
    <html>
    <head>
        <title>Spare Parts Report</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            * { box-sizing: border-box; }
            body {
                font-family: Arial, sans-serif;
                font-size: 100%;
                color: #333333;
                background-color: #FFFFFF;
                padding: 5px;
            }
            h1 { font-size: 2.0rem; color: #2C3E50; margin-bottom: 5px; }
            p { font-size: 0.8rem; color: #7F8C8D; margin-top: 0; margin-bottom: 20px; }
            .toolbar {
                display: flex;
                flex-wrap: wrap;
                gap: 15px;
                align-items: center;
                margin-bottom: 5px;
            }
            button {
                background-color: #34495E;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                cursor: pointer;
                font-weight: bold;
                font-size: 0.85rem;
            }
            button:hover { background-color: #2C3E50; }
            #partSearch {
                padding: 8px 12px;
                flex-grow: 1;
                min-width: 180px;
                max-width: 500px;
                border: 1px solid #BDC3C7;
                border-radius: 4px;
                font-size: 0.85rem;
            }
            .global-dashboard-container {
                display: flex;
                flex-direction: column;
                gap: 10px;
                margin-top: 15px;
                margin-bottom: 20px;
                background: #FAFAFA;
                border: 1px solid #CFD8DC;
                border-radius: 6px;
                padding: 12px;
            }
            .dashboard-bar-row {
                display: flex;
                flex-direction: column;
                gap: 4px;
            }
            .dashboard-bar-label {
                font-size: 0.75rem;
                font-weight: bold;
                color: #455A64;
            }
            .dashboard-bar-wrapper {
                border-radius: 4px;
                overflow: hidden;
                display: flex;
                height: 14px;
                font-size: 0.68rem;
                text-align: center;
                line-height: 14px;
                font-weight: bold;
                background-color: #ECEFF1;
            }
            .bar-segment { 
                transition: all 0.2s ease; 
                white-space: nowrap;
                overflow: hidden;
                text-overflow: ellipsis;
                padding: 0 4px;
            }
            .head-container {
                display: flex;
                flex-direction: row;
                justify-content: space-between;
                align-items: stretch;
                gap: 20px;
                margin-bottom: 10px;
            }
            .head-container > div {
                background: #FFFFFF;
                border-radius: 6px;
                padding: 5px;
                box-shadow: 0 1px 3px rgba( 0, 0, 0, 0.1);
                border: 1px solid #CFD8DC;
                display: flex;
                flex-direction: column;
            }
            .head-container > div:first-child { flex: 1; min-width: 300px; }
            .head-container > div:last-child { flex: 1; min-width: 400px; }
            .head-container > div:last-child table { width: 100%; height: 100%; border-collapse: collapse; }
            .head-container > div:last-child td { padding: 0px; padding-left: 10px; }
            table.fixed { 
                table-layout: fixed; 
                width: 100%; 
                border-collapse: collapse;
                background: #FFFFFF;
                box-shadow: 0 1px 3px rgba( 0, 0, 0, 0.1);
                border-radius: 4px;
            }
            th {
                background-color: #ECEFF1; 
                color: #37474F; 
                font-weight: 600;
                border-bottom: 2px solid #CFD8DC;
                padding: 12px 10px;
            }
            td { 
                border-bottom: 1px solid #E0E0E0;
                padding: 10px;
                font-size: 0.8rem;
            }
            .badge-container {
                display: flex;
                flex-direction: column;
                gap: 5px;
                align-items: center;
                justify-content: center;
            }
            .status-badge, .successor-badge {
                display: block;
                width: 90%;
                max-width: 150px;
                padding: 5px 8px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 0.75rem;
                text-align: center;
                box-shadow: inset 0 -1px 0 rgba(0, 0, 0, 0.15);
            }
            .status-badge { cursor: help; }
            .note-container {
                margin-top: 6px;
                padding: 6px 10px;
                border-radius: 4px;
                font-size: 0.75rem;
                border-left: 5px solid;
            }
            .colw10 { width: 40px; }
            .colw40 { width: 180px; }
            .colw30 { width: 170px; }
            .colw20 { width: 90px; }
            .pm000bkgd { color: #505070; background-color: #FFFFFF; background: linear-gradient(to right, #E7E7F3, #A0A0C0);}
            .pm280bkgd { color: #186221; background-color: #90E39A; background: linear-gradient(to right, #DEF7E1, #90E39A);}
            .pm300bkgd { color: #186221; background-color: #90E39A; background: linear-gradient(to right, #DEF7E1, #90E39A);}
            .pm400bkgd { color: #A35A00; background-color: #FFD6A3; background: linear-gradient(to right, #FFEDD6, #FFE799);}
            .pm410bkgd { color: #7A4300; background-color: #FFCE92; background: linear-gradient(to right, #FFEDD6, #FFCE92);}
            .pm490bkgd { color: #6E0C18; background-color: #F28794; background: linear-gradient(to right, #F9C8CD, #F28794);}
            .pm500bkgd { color: #6E0C18; background-color: #F28794; background: linear-gradient(to right, #F9C8CD, #F28794);}
            .pmerrbkgd { color: #474747; background-color: #EBEBEB; background: linear-gradient(to right, #F5F5F5, #E0E0E0);}
            .pmsucbkgd { color: #193E61; background-color: #9EC3E6; background: linear-gradient(to right, #CEE1F2, #9EC3E6);}
            .pmnotbkgd { color: #193E61; background-color: #5B9BD5; background: linear-gradient(to right, #FFFFFF, #5B9BD5);}
            .no-print { display: none; }
            .no-wrap {
                white-space: nowrap;
                word-break: normal;
                overflow-wrap: normal;
                overflow: hidden;
                text-overflow: ellipsis;
            }
            .report-footer {
                text-align: center;
                font-size: 11px;
                color: #a0a0a0;
                margin-top: 50px;
                padding-top: 20px;
                border-top: 1px solid #f0f0f0;
                font-family: Arial, sans-serif;
            }
            .report-link {
                color: #707070;
                text-decoration: none;
                cursor: pointer;
                font-weight: 500;
                transition: color 0.2s ease;
            }
            .report-link:hover {
                color: #24292e;
                text-decoration: underline;
            }
            @media print {
                body, table, td, th {
                    -webkit-print-color-adjust: exact !important;
                    print-color-adjust: exact !important;
                }
                button, #partSearch, .toolbar { display: none !important; }
                body { font-size: 12px !important; }
                h1 { 
                    font-size: 1.4rem !important; 
                    margin-bottom: 2px !important;
                }
                p { 
                    margin-bottom: 10px !important; 
                }
                th {
                    padding: 6px 8px !important;
                    font-size: 0.75rem !important;
                }
                td { 
                    padding: 5px 6px !important;
                    font-size: 0.7rem !important;
                }
                .status-badge, .successor-badge {
                    font-size: 0.65rem !important;
                    padding: 2px 4px !important;
                }
                .note-container {
                    font-size: 0.65rem !important;
                    padding: 4px 6px !important;
                    margin-top: 3px !important;
                }
                tr {
                    break-inside: avoid !important;
                    page-break-inside: avoid !important;
                }
                thead {
                    display: table-header-group !important;
                }
                thead tr {
                    break-after: avoid !important;
                    page-break-after: avoid !important;
                }
            }
            @media (max-width: 768px) {
               .head-container { flex-direction: column; }
                #partSearch { width: 100%; }
            }
        </style>
    </head>
    <body>
        <div class="head-container">
            <div>
                <h1>Spare Parts Report</h1>
                <p>Generated by SparesNet on: ${new Date().toLocaleString()}</p>
                <div class="toolbar">
                    <button onclick="window.print()">Print</button>
                    <button id="saveHtmlBtn">Save</button>
                    <input type="text" id="partSearch" placeholder="🔍 Search part number or description...">
                    <div style="display: flex; gap: 8px; flex-wrap: wrap; align-items: center; font-size: 0.85rem; font-weight: bold; background: #ECEFF1; padding: 6px 12px; border-radius: 4px; border: 1px solid #CFD8DC; margin-top: 10px;">
                        <span style="color: #37474F; margin-right: 2px;">Filter PM:</span>
                        <label style="cursor:pointer; margin-right: 5px; color: #005A9E;"><input type="checkbox" id="toggleAllPM" checked> <strong>(All)</strong></label>
                        <label style="cursor:pointer;"><input type="checkbox" class="pm-filter" value="PM280" checked> PM280</label>
                        <label style="cursor:pointer;"><input type="checkbox" class="pm-filter" value="PM300" checked> PM300</label>
                        <label style="cursor:pointer;"><input type="checkbox" class="pm-filter" value="PM400" checked> PM400</label>
                        <label style="cursor:pointer;"><input type="checkbox" class="pm-filter" value="PM410" checked> PM410</label>
                        <label style="cursor:pointer;"><input type="checkbox" class="pm-filter" value="PM490" checked> PM490</label>
                        <label style="cursor:pointer;"><input type="checkbox" class="pm-filter" value="PM500" checked> PM500</label>
                        <label style="cursor:pointer;"><input type="checkbox" class="pm-filter" value="UNKNOWN" checked> Unknown</label>
                    </div>
                </div>
                <div class="dashboard-bar-row" id="pmd_state2">
                  <div class="dashboard-bar-label">PMD Consolidated Phases</div>
                  <div class="dashboard-bar-wrapper">
                    ${countPMact > 0 ? `<div class="bar-segment pm300bkgd" style="flex: ${countPMact}; color: #186221;" title="Active: ${countPMact}">${countPMact} Active</div>` : ''}
                    ${countPMcan > 0 ? `<div class="bar-segment pm410bkgd" style="flex: ${countPMcan}; color: #7A4300;" title="Phase-Out: ${countPMcan}">${countPMcan} Phase-Out</div>` : ''}
                    ${countPMdis > 0 ? `<div class="bar-segment pm500bkgd" style="flex: ${countPMdis}; color: #6E0C18;" title="Discontinued: ${countPMdis}">${countPMdis} Discontinued</div>` : ''}
                    ${countPMerr > 0 ? `<div class="bar-segment pmerrbkgd" style="flex: ${countPMerr}; color: #474747;" title="Unknown: ${countPMerr}">${countPMerr} Unknown</div>` : ''}
                  </div>
                </div>
            </div>
            <div>
                <table style="font-weight: bold; padding: 4px;">
                    <thead>
                        <tr style="background-color: #E7E7F3; color: #303050; font-size: 0.75rem;">
                            <th style="border: 0px; text-align: left;">Milestone Stage</th>
                            <th style="border: 0px; text-align: center;" colspan="2">Count</th>
                            <th style="border: 0px; text-align: center;">Min Age</th>
                            <th style="border: 0px; text-align: center;">Avg Age</th>
                            <th style="border: 0px; text-align: center;">Max Age</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr class="pm280bkgd" style="font-size: 0.75rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>PM280:</strong> Sales Release</td>
                            <td style="border: 0px; text-align: center;">${countPM280}</td>
                            <td style="border: 0px; text-align: center;" class="pm280bkgd" rowspan="2">${countPMact}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM280.min}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM280.avg}</td>
                            <td style="border: 0px; text-align: center; font-weight: bold;">${statsPM280.max}</td>
                        </tr>
                        <tr class="pm300bkgd" style="font-size: 0.75rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>PM300:</strong> Delivery Rel.</td>
                            <td style="border: 0px; text-align: center;">${countPM300}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM300.min}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM300.avg}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM300.max}</td>
                        </tr>
                        <tr class="pm400bkgd" style="font-size: 0.75rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>PM400:</strong> Phase Out Ann.</td>
                            <td style="border: 0px; text-align: center;">${countPM400}</td>
                            <td style="border: 0px; text-align: center;" class="pm410bkgd" rowspan="2">${countPMcan}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM400.min}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM400.avg}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM400.max}</td>
                        </tr>
                        <tr class="pm410bkgd" style="font-size: 0.75rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>PM410:</strong> Prod. Cancell.</td>
                            <td style="border: 0px; text-align: center;">${countPM410}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM410.min}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM410.avg}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM410.max}</td>
                        </tr>
                        <tr class="pm490bkgd" style="font-size: 0.75rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>PM490:</strong> Prod. Discont.</td>
                            <td style="border: 0px; text-align: center;">${countPM490}</td>
                            <td style="border: 0px; text-align: center;" class="pm500bkgd" rowspan="2">${countPMdis}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM490.min}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM490.avg}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM490.max}</td>
                        </tr>
                        <tr class="pm500bkgd" style="font-size: 0.75rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>PM500:</strong> End Lifecycl.</td>
                            <td style="border: 0px; text-align: center;">${countPM500}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM500.min}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM500.avg}</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">${statsPM500.max}</td>
                        </tr>
                        <tr class="pmerrbkgd" style="font-size: 0.75rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>NO PM:</strong> Unknown.</td>
                            <td style="border: 0px; text-align: center;">${countPMerr}</td>
                            <td style="border: 0px; text-align: center;" class="pmerrbkgd">-</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">-</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">-</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">-</td>
                        </tr>
                        <tr class="pmsucbkgd" style="font-size: 0.8rem;">
                            <td style="border: 0px; text-align: left;" class="no-wrap"><strong>Total:</strong> </td>
                            <td style="border: 0px; text-align: center;">${countPMall}</td>
                            <td style="border: 0px; text-align: center;" class="pmsucbkgd">-</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">-</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">-</td>
                            <td style="border: 0px; text-align: center; font-weight: normal;">-</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
        <div class="global-dashboard-container">
            <div class="dashboard-bar-row">
                <div class="dashboard-bar-label">PM Milestone Stages Distribution</div>
                <div class="dashboard-bar-wrapper">
                    ${countPM280 > 0 ? `
                        <div class="bar-segment pm280bkgd" 
                            style="flex: ${countPM280}; color: #186221;"
                            title="PM280: Sales Release ${countPM280} items">
                            ✅ ${countPM280} PM280
                        </div>` : ''}
                    ${countPM300 > 0 ? `
                        <div class="bar-segment pm300bkgd" 
                            style="flex: ${countPM300}; color: #186221;"
                            title="PM300: Delivery Release ${countPM300} items">
                            ✅ ${countPM300} PM300
                        </div>` : ''}
                    ${countPM400 > 0 ? `
                        <div class="bar-segment pm400bkgd" 
                            style="flex: ${countPM400}; color: #A35A00;"
                            title="PM400: Phase Out Announcement ${countPM400} items">
                            ⏳ ${countPM400} PM400
                        </div>` : ''}
                    ${countPM410 > 0 ? `
                        <div class="bar-segment pm410bkgd" 
                            style="flex: ${countPM410}; color: #7A4300;"
                            title="PM410: Product Cancellation ${countPM410} items">
                            ⏳ ${countPM410} PM410
                        </div>` : ''}
                    ${countPM490 > 0 ? `
                        <div class="bar-segment pm490bkgd" 
                            style="flex: ${countPM490}; color: #6E0C18;"
                            title="PM490: Product Discontinuation ${countPM490} items">
                            ⚠️ ${countPM490} PM490
                        </div>` : ''}
                    ${countPM500 > 0 ? `
                        <div class="bar-segment pm500bkgd" 
                            style="flex: ${countPM500}; color: #6E0C18;"
                            title="PM500: End of Lifecycle ${countPM500} items">
                            ⚠️ ${countPM500} PM500
                        </div>` : ''}
                    ${countPMerr > 0 ? `
                        <div class="bar-segment pmerrbkgd" 
                            style="flex: ${countPMerr}; color: #474747;"
                            title="Unknown: ${countPMerr} items">
                            ${countPMerr} Unknown
                        </div>` : ''}
                </div>
            </div>
            <div class="dashboard-bar-row">
                <div class="dashboard-bar-label">Risk Lifecycle Context (Critical Gaps & Phase-Outs)</div>
                <div class="dashboard-bar-wrapper">
                    ${activeCount > 0 ? `
                        <div class="bar-segment pm300bkgd" 
                            style="flex: ${activeCount};" 
                            title="Stable (Active Lifecycle Status): ${activeCount} items (${pctActive}%)">
                            ✅ ${activeCount} Active Status
                        </div>` : ''}
                    ${orangeParts > 0 ? `
                        <div class="bar-segment pm410bkgd" 
                            style="flex: ${orangeParts};" 
                            title="Monitored Risk (Phase-Out Announced / Cancelled): ${orangeParts} items (${pctOrange}%)">
                            ⏳ ${orangeParts} Phase-Out
                        </div>` : ''}
                    ${redWithSuccessor > 0 ? `
                        <div class="bar-segment pmsucbkgd" 
                            style="flex: ${redWithSuccessor};" 
                            title="Mitigated Risk (Discontinued with Successor Available): ${redWithSuccessor} items (${pctWithSuccessor}%)">
                            🔄 ${redWithSuccessor} With Successor
                        </div>` : ''}
                    ${redWithoutSuccessor > 0 ? `
                        <div class="bar-segment pm500bkgd" 
                            style="flex: ${redWithoutSuccessor};" 
                            title="Critical Risk (Discontinued with No Successor): ${redWithoutSuccessor} items (${pctNoSuccessor}%)">
                            ⚠️ ${redWithoutSuccessor} No Successor
                        </div>` : ''}
                </div>
            </div>
            <div class="dashboard-bar-row">
                <div class="dashboard-bar-label">Longest Time Elapsed Since PM410 Production Cancellation Milestone (0 - 10 Years Margins)</div>
                <div class="dashboard-bar-wrapper" style="background-color: #ECEFF1;">
                    ${hasPM410Parts ? `
                        <div class="bar-segment pm410bkgd" style="flex: ${filledMonths}; color: #7A4300; text-align: center; font-weight: bold;" title="Longest PM410 change duration: ${durationText}">
                            Longest PM410: ${durationText} ${maxPM410Months > 120 ? '(Capped at 10y)' : ''}
                        </div>
                        ${emptyMonths > 0 ? `<div style="flex: ${emptyMonths};"></div>` : ''}
                    ` : `
                        <div class="bar-segment pmerrbkgd" style="flex: 120; color: #474747; text-align: center;">No PM410 Part Milestone Data Available</div>
                    `}
                </div>
        </div>
        <table class="fixed">
            <thead>
                <tr>
                    <th class="colw10">No</th>
                    <th class="colw40">Part Number</th>
                    <th>Description & Notes</th>
                    <th class="colw30">State</th>
                    <th class="colw20">Disp.Time</th>
                </tr>
            </thead>
            <tbody>
            ${filteredData.map((row, index) => {
                // Dynamically tag element with centralized translation rule
                const rowPmCode = normalizePMCode(row[4]);
                // Build plaintext timeline string for tooltip replacement
                const rawTimelineText = row[5] ? row[5].toString().replace(/_x000d_/gi, '\n') : '';
                // Parse the milestone dates to extract the last valid chronological date entry
                let lastMilestoneInfo = '';
                if (row[5]) {
                    const milestoneDates = parseMilestoneDates(row[5]);
                    let latestDate = null;
                    let latestCode = null;
                    // Iterate through standard codes matching your configuration map
                    for (const [pmCode, dateObj] of Object.entries(milestoneDates)) {
                        if (dateObj && !isNaN(dateObj.getTime())) {
                            if (!latestDate || dateObj >= latestDate) {
                                latestDate = dateObj;
                                latestCode = pmCode;
                            }
                        }
                    }
                    if (latestDate && latestCode) {
                        // Formatting option matches standard localized short date conversion
                        const formattedDate = latestDate.toLocaleDateString();
                        lastMilestoneInfo = `${latestCode}: ${formattedDate}`;
                    }
                }
                // Safely verify if a system warning note exists
                const hasNote = row[6] && row[6].toString().trim() !== "";
                const systemNoteHtml = hasNote
                    ? `<div class="note-container ${getStatusStyle(2, row[6])}">
                        <strong>Status Note:</strong> ${row[6]}
                       </div>` 
                    : '';
                const successorForText = row[29] ? row[29].toString().trim() : '';
                const dynamicSuccessorNoteHtml = successorForText
                    ? `<div class="note-container pmsucbkgd" style="margin-top: 6px;">
                        <strong>Successor:</strong> added automatically as a successor for <strong>${successorForText}</strong>.
                    </div>`
                    : '';
                // Successor part number / alert note logic
                const successorText = row[28] ? row[28].toString().trim() : '';
                const statusText = row[4] ? row[4].toString().toLowerCase().trim() : '';
                let successorBadgeHtml = '';
                if (successorText) {
                    successorBadgeHtml = `
                        <span class="successor-badge ${getStatusStyle(3, row[28])}">
                            Successor: ${successorText}
                        </span>`;
                } else if (statusText.includes('discont') || statusText.includes('end')) {
                    successorBadgeHtml = `
                        <span class="successor-badge pm500bkgd">
                            No successor found
                        </span>`;
                }
                return `
                <tr class="part-row" data-pm="${rowPmCode}">
                    <td style="text-align: center; color: #7F8C8D; font-weight: normal;">${index + 1}</td>
                    <td style="font-weight: bold; color: #2C3E50;">${row[1]}</td>
                    <td>
                        <div style="font-weight: 500; line-height: 1.4; color: #333;">${row[2]}</div>
                        ${systemNoteHtml}
                    </td>
                    <td style="text-align: center; vertical-align: middle;">
                        <div class="badge-container">
                            <span class="status-badge ${getStatusStyle(1, row[4])}" title="${rawTimelineText}">
                                ${row[4] || 'Unknown'}${lastMilestoneInfo ? `<br><small style="font-weight:normal; opacity:0.85;">ℹ️ ${lastMilestoneInfo}</small>` : ''}
                            </span>
                            ${successorBadgeHtml}
                            ${dynamicSuccessorNoteHtml}
                        </div>
                    </td>
                    <td style="text-align: center; color: #2C3E50;">${row[11] || '-'}</td>
                </tr>
                `;
            }).join('')}
            </tbody>
        </table>
        <script>
            const searchInput = document.getElementById('partSearch');
            const checkboxes = document.querySelectorAll('.pm-filter');
            const toggleAllCheckbox = document.getElementById('toggleAllPM');
            const rows = document.querySelectorAll('.part-row');
            function filterTable() {
                const query = searchInput.value.toUpperCase();
                // Build an array of checked PM stage values
                const activePMs = Array.from(checkboxes)
                    .filter(cb => cb.checked)
                    .map(cb => cb.value);
                rows.forEach(row => {
                    const partNumber = row.cells[1].textContent.toUpperCase();
                    const description = row.cells[2].textContent.toUpperCase();
                    const rowPm = row.getAttribute('data-pm');
                    // Conditions
                    const matchesSearch = partNumber.includes(query) || description.includes(query);
                    const matchesPM = activePMs.includes(rowPm);
                    if (matchesSearch && matchesPM) {
                        row.style.display = '';
                    } else {
                        row.style.display = 'none';
                    }
                });
            }
            // Debouncer for typing filters
            function debounce(func, delay) {
                let timeout;
                return (...args) => {
                    clearTimeout(timeout);
                    timeout = setTimeout(() => func.apply(this, args), delay);
                };
            }
            searchInput.addEventListener('input', debounce(filterTable, 150));
            checkboxes.forEach(cb => {
                cb.addEventListener('change', () => {
                    const allChecked = Array.from(checkboxes).every(c => c.checked);
                    const noneChecked = Array.from(checkboxes).every(c => !c.checked);
                    if (allChecked) {
                        toggleAllCheckbox.checked = true;
                        toggleAllCheckbox.indeterminate = false;
                    } else if (noneChecked) {
                        toggleAllCheckbox.checked = false;
                        toggleAllCheckbox.indeterminate = false;
                    } else {
                        toggleAllCheckbox.indeterminate = true;
                    }
                    filterTable();
                });
            });
            toggleAllCheckbox.addEventListener('change', function() {
                checkboxes.forEach(cb => cb.checked = this.checked);
                filterTable();
            });
            const printButton = document.getElementById('saveHtmlBtn');
            if (printButton) {
                printButton.addEventListener('click', () => {
                    if (window.api && typeof window.api.selectSaveHtmFile === 'function') {
                        window.api.selectSaveHtmFile(document.documentElement.outerHTML);
                    }// else {
                    //    alert("API context bridge is not available.");
                    //}
                });
            }
        </script>
    </body>
    </html>`;
    return reportHTML;
};

