import { app, BrowserWindow, BrowserView, ipcMain, nativeTheme, dialog } from 'electron';
import pie from 'puppeteer-in-electron';
import puppeteer from 'puppeteer-core';
import path from 'path';
import fs from 'fs/promises';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Configurable defaults
const READ_CONFIG = {
    maxRetries: 3,
    initBackoffMs: 800,
    backoffFactor:   2,
    perReqTmOutMs: 15000, // page wait timeout
    interReqDelMs:  2000, // pause between sites
    viewBounds: { x: 300, y: 200, width: 800, height: 600 }
};

let main;
let view;
let browser;

let activeView = null;
let isRunning = false;

// Must be called before app.whenReady()
await pie.initialize(app);

async function createWindow() {
    browser = await pie.connect(app, puppeteer);
    // Main window
    main = new BrowserWindow({
        width: 1000,
        height: 700,
        webPreferences: {
        preload: path.join(__dirname, 'preload.cjs'),
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: false
        }
    });
    // Load GUI
    main.loadFile(path.join(__dirname, 'dist/index.html'));
}

app.whenReady().then(createWindow);

// Handle 'Dark/Light mode' on toggle
ipcMain.handle('dark-mode:toggle', () => {
    if (nativeTheme.shouldUseDarkColors) {
        nativeTheme.themeSource = 'light'
    } else {
        nativeTheme.themeSource = 'dark'
    }
    return nativeTheme.shouldUseDarkColors
})

// Handle 'Dark/Light mode' on start
ipcMain.handle('dark-mode:system', () => {
    nativeTheme.themeSource = 'system'
})

// Open xls file [select a file and return panth and name]
ipcMain.handle('select-open-xls-file', async () => {
    const appFolder = app.getPath('userData'); 
    const defaultPath1 = path.join(appFolder, 'Spreadsheet_Export.xlsx');
    const result = await dialog.showOpenDialog({
        title: 'Import Data from Excel File',
        defaultPath: defaultPath1,
        filters: [
            { name: 'Excel Files', extensions: ['xls', 'xlsx', 'xlsm'] }
        ],
        properties: ['openFile']
    });
    if (result.canceled) return null;
    return result.filePaths[0];
});

// Handle the Save Excel File dialog box
ipcMain.handle('select-save-xls-file', async () => {
    const appFolder = app.getPath('userData'); 
    const defaultPath1 = path.join(appFolder, 'Spreadsheet_Export.xlsx');
    const { filePath } = await dialog.showSaveDialog({
        title: 'Export Data to Excel',
        defaultPath: defaultPath1,
        filters: [
            { name: 'Excel Files', extensions: ['xls', 'xlsx', 'xlsm'] }
        ]
    });
    //if (result.canceled) return null;
    return filePath;
});

// Handle the Preview Html Report
ipcMain.on('preview_report', async (event, html) => {
    const preview = new BrowserWindow({
        width: 900,
        height: 700,
        title: 'Report Preview',
        webPreferences: {
            nodeIntegration: false,
            preload: path.join(__dirname, 'preload.cjs'),
            contextIsolation: true,
            sandbox: false
        }
    });
    preview.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
});

// Handle the Save Html Report File dialog box
ipcMain.handle('select-save-htm-file', async (event, html) => {
    const appFolder = app.getPath('userData'); 
    const defaultPath1 = path.join(appFolder, 'SparesNetReport.html');
    const { filePath } = await dialog.showSaveDialog({
        title: 'Export Report to Html',
        defaultPath: defaultPath1,
        filters: [
            { name: 'HTML Files', extensions: ['html', 'htm'] }
        ]
    });
    if (filePath) {
        try {
            await fs.writeFile(filePath, html, 'utf-8');
            return { success: true, path: filePath };
        }
        catch {
        }
    }
    return { success: false };
});

// Handle the Export to PDF Report File
ipcMain.handle('select-save-pdf-file', async (event, html) => {
    //AICI
    const appFolder = app.getPath('userData'); 
    const defaultPath1 = path.join(appFolder, 'SparesNetReport.pdf');
    const { filePath } = await dialog.showSaveDialog({
        title: 'Export Report to Pdf',
        defaultPath: defaultPath1,
        filters: [
            { name: 'PDF Files', extensions: ['pdf'] }
        ]
    });
    if (!filePath) {
        return { success: false, cancelled: true };
    }
    const preview = new BrowserWindow({
        width: 900,
        height: 700,
        show: false,
        //title: 'Report Preview',
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true
        }
    });
    try {
        await preview.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
        const pdfOptions = {
            marginsType: 0, // 0 = default margins, 1 = no margins, 2 = minimum margins
            pageSize: 'A4',
            printBackground: true, // Retains your background colors
            printSelectionOnly: false,
            landscape: false
        };
        const pdfBuffer = await preview.webContents.printToPDF(pdfOptions);
        await fs.writeFile(filePath, pdfBuffer);
        return { success: true, path: filePath };
    } catch (error) {
        //console.error('Native PDF Generation failed:', error);
        return { success: false, error: error.message };
    } finally {
        preview.close();
    }
});


// Helper: create a BrowserView instance
function createBrowserView() {
    if (activeView) return activeView;
    activeView = new BrowserView({
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            webSecurity: true,
            // optional: for headless/offscreen rendering
            // offscreen: true
        }
    });
    return activeView;
}

// Helper: safely destroy the BrowserView
function destroyBrowserView() {
    try {
        if (!activeView) return;
        // If attached, remove first
        if (main && main.getBrowserViews().includes(activeView)) {
            main.removeBrowserView(activeView);
        }
        // Destroy webContents and dereference
        try { activeView.webContents.destroy(); } catch (err) { /* ignore */ }
        activeView = null;
    } catch (err) {
        activeView = null;
    } finally {
        if (main) main.focus();
    }
}

// Helper: attach view and map Puppeteer page
async function attachAndGetPage() {
    const view = createBrowserView();
    main.setBrowserView(view);
    view.setBounds(READ_CONFIG.viewBounds);
    view.setAutoResize({ width: true, height: true });
    // Ensure Puppeteer maps to the attached view
    const page = await pie.getPage(browser, view);
    return { view, page };
}

// Helper: detach but keep view instance alive
function detachViewSafe() {
    try {
        if (activeView && main.getBrowserViews().includes(activeView)) {
            main.removeBrowserView(activeView);
        }
    } catch (err) {
    } finally {
        if (main) main.focus();
    }
}

// Helper: Formats PLC milestone text
function formatPlmDataForExcel(rawText) {
    // @param {string} rawText - The unformatted text scraped from the page
    // @returns {string} - Formatted multi-line text separated by CRLF (\r\n)
    if (!rawText) return "";
    const normalized = rawText.replace(/\s+/g, ' ').trim();
    // Exact milestones ordered chronologically by lifecycle stage
    const phases = [
        "Sales Release",
        "Delivery Release",
        "Phase Out Announce",
        "Prod. Cancellation",
        "Prod. Discont.",
        "End Prod.Lifecycl."
    ];
    const cellLines = [];
    // 3. Extract dates attached to each phase if they exist
    for (const phase of phases) {
        // Escape special characters in phase names
        const escapedPhase = phase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        // Look for an optional date in brackets (e.g. "03/14/2000")
        const regex = new RegExp(escapedPhase + "(?:\\s*\\(([^)]+)\\))?");
        const match = normalized.match(regex);
        if (match) {
            const date = match[1];
            if (date) {
                // Phase exists and has date
                cellLines.push(`${phase}: ${date}`);
            } else {
                 // Phase exists but hasn't received an active date yet
                cellLines.push(phase);
            }
        }
    }
    // Join using CRLF (\r\n) for Excel Cell
    return cellLines.join("\r\n");
}

// Helper: Detect and auto-dismiss cookie banners or exceptions safely
async function dismissCookieBanner(page) {
    try {
        // Wait briefly since cookie wrappers often inject 200-400ms after DOMContentLoaded
        await new Promise(r => setTimeout(r, 500));
        // Direct targeted IDs/Classes for common corporate compliance engines (e.g., OneTrust)
        const targetSelectors = [
            '#onetrust-reject-all-handler', // OneTrust Preferable Reject
            '#onetrust-accept-btn-handler', // OneTrust Fallback Accept
            '#uc-btn-deny',                 // Usercentrics Deny
            '#uc-deny-all-button',          // Usercentrics Deny All [*]
            '#uc-btn-accept',               // Usercentrics Accept
            '.onetrust-close-btn-handler'
        ];
        for (const selector of targetSelectors) {
            const btn = await page.$(selector);
            if (btn) {
                const isVisible = await page.evaluate(el => {
                    const style = window.getComputedStyle(el);
                    return style && style.display !== 'none' && style.visibility !== 'hidden' && el.offsetWidth > 0;
                }, btn);
                if (isVisible) {
                    await btn.click();
                    //await new Promise(r => setTimeout(r, 600)); // Wait for fading animations
                    await new Promise(r => setTimeout(r, 6000)); // Wait for fading animations
                    return true;
                }
            }
        }
        // Fallback: Programmatic DOM query for layout text labels
        const clickedViaText = await page.evaluate(() => {
            const interactables = Array.from(document.querySelectorAll('button, a, [role="button"], sie-button'));
            // Priority ordering: explicit Rejections first, then general Acceptance
            const triggerPhrases = ['reject all', 'deny all', 'decline', 'refuse', 'accept all', 'allow all', 'accept cookies'];
            for (const item of interactables) {
                const text = (item.textContent || '').trim().toLowerCase();
                if (triggerPhrases.some(phrase => text === phrase || text.includes(phrase))) {
                    const style = window.getComputedStyle(item);
                    if (style && style.display !== 'none' && style.visibility !== 'hidden' && item.offsetWidth > 0) {
                        item.click();
                        return true;
                    }
                }
            }
            return false;
        });
        if (clickedViaText) {
            await new Promise(r => setTimeout(r, 600));
        }
    } catch (err) {
        // If a banner isn't present or fails to click
    }
}

// Helper: navigate with retries and optional proxy
async function navigateWithRetries(page, url, options = {}) {
    const { maxRetries, initBackoffMs, backoffFactor, perReqTmOutMs } = READ_CONFIG;
    let attempt = 0;
    let lastErr = null;
    // Try to connect
    while (attempt < maxRetries) {
        try {
            attempt++;
            // If proxy is provided: set it via page.browserContext().overridePermissions or launch args.
            //
            // Navigate using view.webContents to preserve BrowserView behavior
            const navPromise = page.waitForNavigation({ timeout: perReqTmOutMs, waitUntil: 'domcontentloaded' })
                .catch(() => null); // swallow if navigation event not fired
            //
            await page.goto(url, { timeout: perReqTmOutMs, waitUntil: 'domcontentloaded' }).catch(() => null);
            // Wait for DOMContent or network settle
            await navPromise;
            // Optionally wait for selector outside this function
            return { success: true };
        } catch (err) {
            lastErr = err;
            const backoff = initBackoffMs * Math.pow(backoffFactor, attempt - 1);
            //console.warn(`Warning: main: Navigate attempt ${attempt} failed for ${url}: ${err.message}. Backing off ${backoff}ms`);
            await new Promise(r => setTimeout(r, backoff));
        }
    }
    return { success: false, error: lastErr };
}

// Helper: Scroll step by step
async function autoScroll(page, maxScrolls){
    await page.evaluate(async (maxScrolls) => {
        await new Promise((resolve) => {
            var totalHeight = 0;
            var distance = 100;
            var scrolls = 0;    // scrolls counter
            var timer = setInterval(() => {
                var scrollHeight = document.body.scrollHeight;
                window.scrollBy(0, distance);
                totalHeight += distance;
                scrolls++;      // increment counter
               // stop scrolling if reached the end or the maximum number of scrolls
                if(totalHeight >= scrollHeight - window.innerHeight || scrolls >= maxScrolls){
                    clearInterval(timer);
                    resolve();
                }
            }, 1000);
        });
    }, maxScrolls);  // pass maxScrolls to the function
}

// Handle 'Start' event: create view, use  it, destroy view
ipcMain.on('start_read_pages', async (event, data) => {
    if (isRunning) {
        main.webContents.send('read_progress', { status: 'Busy', message: 'Reading already running' });
        return;
    }
    isRunning = true;
    const defaults = {
        index: -1,
        mlfb_: ' ', //MLFB
        artno: ' ', //MLFB
        descr: ' ', //Product Description
        famil: ' ', //Product Family
        plm_s: ' ', //Product Lifecycle (PLM)
        plm_d: ' ', //PLM Effective Date
        notes: ' ', //Notes
        pri_g: ' ', //Price Group
        raw_m: ' ', //Surcharge for Raw Materials
        met_f: ' ', //Metal Factor
        exp_c: ' ', //'Export Control Regulations
        dispt: ' ', //Dispatch Time
        net_w: ' ', //Net Weight (kg)
        dim_n: ' ', //Product Dimensions (W x L x H)
        dim_p: ' ', //Packaging Dimension
        unitp: ' ', //Minimum Order Quantity
        unitq: ' ', //Quantity Unit
        packq: ' ', //Packaging Quantity
        ean__: ' ', //EAN
        upc__: ' ', //UPC
        comcd: ' ', //Commodity Code
        catal: ' ', //KZ_FDB/ CatalogID
        prodg: ' ', //Product Group
        orign: ' ', //Country of origin
        rohs_: ' ', //Compliance with RoHS
        pclas: ' ', //Product class
        takeb: ' ', //Obligation for taking back
        clsif: ' ', //Classifications
        succe: ' ', //Successor
    };
    main.webContents.send('read_progress', { status: 'Start reading ...', message: ' ', index: 0, total: data.length });
    // Create view instance once for the whole run
    createBrowserView();
    let i = 0;
    try {
        for (i = 0; i < data.length; i++) {
            const info = { ... defaults };  //shallow clone
            const [row, urlEntry] = data[i];
            info.index = row;
            // Accept either string or { url, proxy } objects
            const url = typeof urlEntry === 'string' ? urlEntry : urlEntry.url;
            const proxy = typeof urlEntry === 'object' ? urlEntry.proxy : undefined;
            main.webContents.send('read_progress', { status: 'Open ', message: url, index: i, total: data.length});
            let page;
            try {
                // Attach view and get Puppeteer page mapped to it
                const attached = await attachAndGetPage();
                page = attached.page;
                // Navigate with retries
                const navResult = await navigateWithRetries(page, url, { proxy, waitUntil: 'domcontentloaded' });
                //const navResult = await navigateWithRetries(page, url, { proxy });
                if (!navResult.success) throw navResult.error || new Error('Navigation failed');
                // Dismiss Cookies early
                await dismissCookieBanner(page);
                // Wait for selector and extract
                const anchorSelector = '.product-metadata-description';
                await page.waitForSelector(anchorSelector, { timeout: 10000, visible: true });
                // Locate the target element using a selector
                const targetElement = await page.$('#specifications');
                // Scroll the target element into view, in order to display commercial data
                await targetElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
                // Dismiss Cookies again
                await dismissCookieBanner(page);
                await new Promise(r => setTimeout(r, 600));
                // Hover (back) to trigger the Product Lifecycle popup
                let plm_data = ' ';
                try {
                    // Find all metadata item blocks on the page
                    const items = await page.$$('.product-metadata-item');
                    for (const item of items) {
                        const text = await page.evaluate(el => el.textContent, item);
                        if (text && text.toLowerCase().includes('product lifecycle')) {
                            // Execute the entire event dispatching and extraction sequence directly inside the DOM context
                            plm_data = await page.evaluate(async (el) => {
                                // Locate Siemens tooltip icon within specific lifecycle row
                                const trigger = el.querySelector('sie-icon, [sieuipopover], .product-metadata-item__tooltip-icon, svg');
                                if (!trigger) return 'No trigger icon found';
                                // Center the element
                                trigger.scrollIntoView({ block: 'center', inline: 'center' });
                                // Helper function to fire pure JS events
                                const fireEvent = (eventName) => {
                                    const ev = new Event(eventName, { bubbles: true, cancelable: true });
                                    trigger.dispatchEvent(ev);
                                };
                                // Programmatically force Angular listeners to trigger without OS mouse focus
                                fireEvent('pointerenter');
                                fireEvent('mouseenter');
                                fireEvent('mouseover');
                                fireEvent('focus');
                                // Fallback: fire a click event in case responsive breakpoints expect tap behavior
                                const clickEv = new MouseEvent('click', { bubbles: true, cancelable: true });
                                trigger.dispatchEvent(clickEv);
                                // Safe sleep to let the layout engine render and insert text content into the overlay
                                await new Promise(r => setTimeout(r, 600));
                                // Scan known Siemens/Angular global overlay containers
                                const popoverSelectors = [
                                    'sie-popover', '.sie-popover', '.sie-ui-popover', 
                                    '.mat-mdc-tooltip', '.mat-tooltip', '.cdk-overlay-pane', 
                                    '.popover-content', '[role="tooltip"]', '.tooltip'
                                ];
                                for (const sel of popoverSelectors) {
                                    const overlays = document.querySelectorAll(sel);
                                    if (overlays.length > 0) {
                                        // Read the newest overlay added (bottom-up sequence)
                                        const lastOverlay = overlays[overlays.length - 1];
                                        const txt = lastOverlay.textContent ? lastOverlay.textContent.trim() : '';
                                        if (txt.length > 0) return txt;
                                    }
                                }
                                // Super fallback: Check the last elements attached directly to document body
                                const bodyChildren = document.body.children;
                                for (let i = bodyChildren.length - 1; i >= bodyChildren.length - 6 && i >= 0; i--) {
                                    const child = bodyChildren[i];
                                    const tag = child.tagName.toLowerCase();
                                    const cls = child.className.toLowerCase();
                                    if (tag.includes('popover') || cls.includes('popover') || 
                                        cls.includes('overlay') || cls.includes('tooltip')) {
                                        const txt = child.textContent ? child.textContent.trim() : '';
                                        if (txt.length > 0) return txt;
                                    }
                                }
                                return ' ';
                            }, item);
                            // Fallback: native Puppeteer hover
                            if (!plm_data || plm_data.trim() === '') {
                                const triggerIcon = await item.$('sie-icon, [sieuipopover], .product-metadata-item__tooltip-icon, svg');
                                if (triggerIcon) {
                                    await triggerIcon.hover().catch(() => {});
                                    await new Promise(r => setTimeout(r, 500));
                                    const actualTooltipSelector = 'sie-popover, .sie-popover, .sie-ui-popover, .mat-mdc-tooltip, .mat-tooltip, .cdk-overlay-pane';
                                    plm_data = await page.evaluate((sel) => {
                                        const elements = document.querySelectorAll(sel);
                                        for (let i = elements.length - 1; i >= 0; i--) {
                                            const txt = elements[i].textContent ? elements[i].textContent.trim() : '';
                                            if (txt.length > 0) return txt;
                                        }
                                        return ' ';
                                    }, actualTooltipSelector).catch(() => ' ');
                                }
                            }
                            break; // Stop iterating once the lifecycle metadata row has been processed
                        }
                    }
                } catch (hoverErr) {
                }
                // Format extracted plm data
                plm_data = formatPlmDataForExcel(plm_data || ' ');
                // Single DOM pass to extract multiple IDs and attributes
                const extrData = await page.evaluate((defaultsInPage, plm_data_) => {
                    // Helper function: read text
                    const readText = (sel) => {
                        const el = document.querySelector(sel);
                        return el ? el.textContent.trim() : null;
                    };
                    // Helper function: read all text
                    const readAllText = (sel) => {
                        const nodes = Array.from(document.querySelectorAll(sel));
                        const texts = nodes.map(n => (n.textContent || '').trim()).filter(Boolean);
                        return texts.length ? texts.join('\n\n') : null;
                    };
                    // Helper function: read attribute
                    const readAttr = (sel, attr) => {
                        const el = document.querySelector(sel);
                        return el ? el.getAttribute(attr) : null;
                    };
                    // Helper function: read number
                    const readNumber = (sel) => {
                        const t = readText(sel);
                        if (!t) return null;
                        const n = parseFloat(t.replace(/[^\d\.\-]/g, ''));
                        return Number.isFinite(n) ? n : null;
                    };
                    // Helper function: normalize keywords
                    const norm = s => (s || '').replace(/[:\u00A0]/g, '').trim().toLowerCase();
                    // Helper: Read Product Data
                    function setResultFromLabel1(result, document, labelText, key) {
                        try {
                            const targetNorm = norm(labelText).toLowerCase();
                            const items = document.querySelectorAll('.product-metadata-item');
                            for (const block of items) {
                                // Extract and normalize the label from the current block
                                let currentLabel = null;
                                const labelSelectors = [
                                    '.product-metadata-item__title-wrapper .product-metadata-item__label',
                                    '.product-metadata-item__label',
                                    'p.product-metadata-item__label'
                                ];
                                for (const selector of labelSelectors) {
                                    const el = block.querySelector(selector);
                                    if (el?.textContent?.trim()) {
                                        currentLabel = norm(el.textContent).toLowerCase();
                                        break;
                                    }
                                }
                                // Check if this is the label we are looking for
                                if (currentLabel && currentLabel.includes(targetNorm)) {
                                    let valueText = null;
                                    const titleWrap = block.querySelector('.product-metadata-item__title-wrapper');
                                    // Try getting the text from the immediate next sibling element
                                    if (titleWrap && titleWrap.nextElementSibling?.textContent?.trim()) {
                                        valueText = titleWrap.nextElementSibling.textContent.trim();
                                    } 
                                    // Fallback selectors if the sibling structure doesn't match
                                    else {
                                        const fallbackEl = block.querySelector(
                                            '.product-metadata-item__label-wrapper p, .product-metadata-item__value, .product-metadata-description p'
                                        );
                                        if (fallbackEl?.textContent?.trim()) {
                                            valueText = fallbackEl.textContent.trim();
                                        }
                                    }
                                    // Assign the found value (or null if empty) and exit immediately
                                    result[key] = valueText || '-';
                                    return;
                                }
                            }
                            // If the loop finishes and no matching label was found
                            result[key] = 'Err. Not Found';
                        } catch (err) {
                            // Keeps your custom error tracking intact
                            result[key] = ' ';
                        }
                    }
                    function setResultFromLabel2(result, document, labelText, key) {
                        try {
                            const targetNorm = norm(labelText).toLowerCase();
                            // Target the correct main content area for commercial data
                            //const sectionContainer = document.querySelector('.commercial-data-section');
                            const sectionContainer = document.querySelector('.commercial-data');
                            if (!sectionContainer) {
                                result[key] = '-'; // Main section container missing completely
                                return;
                            }
                            // Query ALL elements inside this section that could contain label text
                            const potentialLabels = Array.from(sectionContainer.querySelectorAll('.commercial-data-section__subtitle'));
                            // Find the element using fully normalized partial matching
                            const targetLabelEl = potentialLabels.find(el => {
                                if (!el.textContent) return false;
                                // Normalize the web page text using the EXACT same rules as the search term
                                const textNorm = norm(el.textContent).toLowerCase();
                                // Check if the page text contains your search term
                                return textNorm.includes(targetNorm); 
                            });
                            if (!targetLabelEl) {
                                result[key] = '--'; // Label string not found anywhere in the section
                                return;
                            }
                            // Verify the layout block/row wrapper exists
                            const block = targetLabelEl.parentElement;
                            if (!block) {
                                result[key] = '---';
                                return;
                            }
                            // Get the text of the immediate next element (the value cell) in the same row
                            const valueEl = targetLabelEl.nextElementSibling;
                            if (valueEl && valueEl.textContent?.trim()) {
                                result[key] = valueEl.textContent.trim();
                            } else {
                                result[key] = ' '; // Next element is missing or has no text
                            }
                        } catch (err) {
                            result[key] = ' '; // General catch-all error for safety
                        }
                    }
                    function getSuccesorFromNotes(NoteText) {
                        try {
                            const regex = /(?:successor|available):\s*([A-Z0-9]{3,7}-[A-Z0-9-]+)/i;
                            const match = NoteText.match(regex);
                            if (match && match[1]){
                                return match[1].trim();
                            } else {
                                return ' ';
                            }
                        } catch (err) {
                            return ' '; // General catch-all error for safety
                        }
                    }
                    const result = { ... defaultsInPage };
                    result.mlfb_ = readText('.intro-section__content-headline') || ' ';
                    result.descr = readText('.product-metadata-description') || ' ';
                    setResultFromLabel2(result, document, 'Product Family', 'famil');
                    setResultFromLabel1(result, document, 'Product lifecycle', 'plm_s');
                    result.plm_d = plm_data_ || ' ';
                    result.notes = readAllText('.intro-section__banner-message') || ' ';
                    setResultFromLabel2(result, document, 'Price Group', 'pri_g');
                    setResultFromLabel2(result, document, 'Surcharge for raw material', 'raw_m');
                    setResultFromLabel2(result, document, 'Metal Factor', 'met_f');
                    setResultFromLabel2(result, document, 'Export Control Regulations', 'exp_c');
                    setResultFromLabel2(result, document, 'Standard delivery time', 'dispt');
                    setResultFromLabel1(result, document, 'Net weight', 'net_w');
                    result.dim_n = ' ';
                    setResultFromLabel1(result, document, 'Packaging dimensions', 'dim_p');
                    setResultFromLabel2(result, document, 'Quantity Unit', 'unitq');
                    setResultFromLabel2(result, document, 'Packaging Quantity', 'packq');
                    setResultFromLabel2(result, document, 'Minimum Order Quantity', 'unitp');
                    setResultFromLabel2(result, document, 'EAN', 'ean__');
                    setResultFromLabel2(result, document, 'UPC', 'upc__');
                    setResultFromLabel2(result, document, 'Commodity Code', 'comcd');
                    setResultFromLabel2(result, document, 'Catalog ID', 'catal');
                    setResultFromLabel2(result, document, 'Product Group', 'prodg');
                    setResultFromLabel2(result, document, 'Country of origin', 'orign');
                    setResultFromLabel2(result, document, 'RoHS directive', 'rohs_');
                    setResultFromLabel1(result, document, 'Product class', 'pclas');
                    setResultFromLabel2(result, document, 'Take-Back Obligations', 'takeb');
                    setResultFromLabel2(result, document, 'Classification', 'clsif');
                    result.succe = getSuccesorFromNotes(result.notes);
                    return result;
                }, info, plm_data);
                Object.assign(info, extrData);
                main.webContents.send('read_received', info);
                main.webContents.send('read_progress', { status: 'Done ', message: url, index: i, total: data.length});
            } catch (err) {
                //console.error('Read error for', url, err);
                info.mlfb_ = 'Error: ' + (err && err.message ? err.message : String(err));
                main.webContents.send('read_received', info);
                main.webContents.send('read_progress', { status: 'Error item ', message: url &&err&&err.message, index: i, total: data.length});
            } finally {
                // Detach after each item so it can't intercept UI
                detachViewSafe();
                // small stabilization pause
                await new Promise(r => setTimeout(r, 500));
            }
            // Inter-request delay
            await new Promise(r => setTimeout(r, READ_CONFIG.interReqDelMs));
        }
        main.webContents.send('read_progress', { status: 'Ready ', message: ' ', index: i, total: data.length});
    } catch (fatal) {
        //console.error('Fatal scraping error:', fatal);
        main.webContents.send('read_progress', { status: 'Fatal Err. ', message: fatal && fatal.message, index: i, total: data.length});
    } finally {
        // Destroy the view instance and free resources
        destroyBrowserView();
        isRunning = false;
    }
});

