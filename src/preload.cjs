//console.log("Preload: Preload loaded");
const { contextBridge, ipcRenderer, shell } = require('electron');
const XLSX = require('xlsx');

contextBridge.exposeInMainWorld('darkMode', {
    toggle: () => ipcRenderer.invoke('dark-mode:toggle'),
});

contextBridge.exposeInMainWorld('api', {
    startReadPages: (data) => ipcRenderer.send('start_read_pages', data),
    onReadProgress: (callback) => ipcRenderer.on('read_progress', (_event, data) => callback(data)),
    onReadReceived: (callback) => ipcRenderer.on('read_received', (_event, data) => callback(data)),
    onReadLog_Data: (callback) => ipcRenderer.on('read_log_data', (_event, data) => callback(data)),
    openExternal: (url) => shell.openExternal(url),
    previewReport: (html) => ipcRenderer.send('preview_report', html),
    selectSaveHtmFile: (html) => ipcRenderer.invoke('select-save-htm-file', html),
    selectSavePdfFile: (html) => ipcRenderer.invoke('select-save-pdf-file', html)
});

contextBridge.exposeInMainWorld('excelAPI', {
    selectOpenXlsFile: () => ipcRenderer.invoke('select-open-xls-file'),
    selectSaveXlsFile: () => ipcRenderer.invoke('select-save-xls-file'),
    importXlsFirstColumn: (filePath, options = {}) => {
        const { skipHeader = false } = options;
        try {
            const workbook = XLSX.readFile(filePath);
            const sheetName = workbook.SheetNames.includes('Data') ? 'Data' : workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            // Extract first column values (preserve row order)
            let firstCol = rows.map(row => (Array.isArray(row) ? row[0] : undefined));
            // Normalize undefined -> null so JSON is predictable
            firstCol = firstCol.map(v => (v === undefined ? null : v));
            if (skipHeader) firstCol = firstCol.slice(1);
            return { sheetName, firstCol, rowCount: rows.length };
        } catch (err) {
            // Re-throw so renderer can catch and show an error
            throw err;
        }
    },
    importXlsData: (filePath, options = {}) => {
        const { skipHeader = false } = options;
        try {
            const workbook = XLSX.readFile(filePath);
            const sheetName = workbook.SheetNames.includes('Data') ? 'Data' : workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            if (skipHeader) rows = rows.slice(1);
            const cleanRows = rows.map(row => {
                if (!Array.isArray(row)) return [];
                return row.map(cell => (cell === undefined ? null : cell));
            });
            return { sheetName, data: cleanRows, rowCount: rows.length };
        } catch (err) {
            // Re-throw so renderer can catch and show an error
            throw err;
        }
    },
    exportXlsData: (filePath, data) => {
        try {
            // Create a new blank workbook container
            const workbook = XLSX.utils.book_new();
            // Convert 2D Array of Arrays (aoa) into a worksheet worksheet
            const worksheet = XLSX.utils.aoa_to_sheet(data);
            // Append the worksheet to the workbook
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
            // Write the file securely to the specified path (creates file if non-existent)
            XLSX.writeFile(workbook, filePath);
            return true;
        } catch (err) {
            throw err;
        }
    },
});
