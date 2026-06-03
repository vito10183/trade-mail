// ====== 設定區 ======
const TRADE_SHEET_NAME = '交易明細';
const INVENTORY_SHEET_NAME = '庫存明細';
const PROCESSED_LABEL_DETAIL = '已處理-交易明細表';

// ====== 每日自動處理函數 ======
function processAllEmails() {
    processTradeDetailEmails();
    cleanupTradeSheet();  
    cleanupInventorySheet();        
    sortAllSheets();       
}

function processTradeDetailEmails() {
    const threads = GmailApp.search(`subject:交易明細表 "統一綜合證券" -label:${PROCESSED_LABEL_DETAIL}`);
    if (threads.length === 0) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tradeSheet = getOrCreateSheet(ss, TRADE_SHEET_NAME, getTradeHeaders());
    const inventorySheet = getOrCreateSheet(ss, INVENTORY_SHEET_NAME, getInventoryHeaders());
    const label = getOrCreateLabel(PROCESSED_LABEL_DETAIL);

    threads.forEach(thread => {
        thread.getMessages().forEach(message => {
            const plainBody = message.getPlainBody();
            const receivedDate = Utilities.formatDate(message.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
            
            const tradeDateMatch = plainBody.match(/交易日\s*[:：]\s*(\d{4}\/\d{2}\/\d{2})/);
            const tradeDate = tradeDateMatch ? tradeDateMatch[1] : '';

            const trades = parseTradeDetails(plainBody);
            trades.forEach(trade => tradeSheet.appendRow([tradeDate, ...trade, receivedDate, '交易明細表']));

            const inventory = parseInventoryDetails(plainBody);
            inventory.forEach(inv => inventorySheet.appendRow([tradeDate, ...inv, receivedDate]));

            message.markRead();
        });
        thread.addLabel(label);
    });
}

function formatStockCode(rawCode) {
    const code = String(rawCode).trim();
    // 使用物件對照表取代一堆 if，未來要新增例外代碼更直觀
    const specialCodes = { '50': '0050', '050': '0050', '6208': '006208', '9816': '009816', '56': '0056', '056': '0056', '878': '00878' };
    
    if (specialCodes[code]) return specialCodes[code];
    return (code.length > 0 && code.length < 4) ? code.padStart(4, '0') : code;
}

function parseTradeDetails(plainBody) {
    const results = [];
    const regex = /(電子單|語音單|臨櫃|APP|智慧單|網路單|行動語音)\s+(普|資|券)\s+(買|賣)\s+(\d{4,6})\s*[\(（](.+?)[\)）]/g;
    let match;
    
    while ((match = regex.exec(plainBody)) !== null) {
        const [ , source, tradeType, buySell, rawCode, stockName] = match;
        const remainingStr = plainBody.substring(match.index + match[0].length);
        const nextMatch = remainingStr.match(/(?:電子單|語音單|臨櫃|APP|智慧單|網路單|行動語音)\s+(?:普|資|券)|合計/);
        const extractArea = nextMatch ? remainingStr.substring(0, nextMatch.index) : remainingStr;
        
        const numbers = extractArea.match(/-?[\d,]+\.?\d*/g);
        if (numbers && numbers.length >= 10) { 
            const nums = numbers.map(n => parseFloat(n.replace(/,/g, '')));
            results.push([source, tradeType, buySell, formatStockCode(rawCode), stockName, nums[0], nums[1], nums[2], nums[3], nums[12] || nums[nums.length - 1]]);
        }
    }
    return results;
}

function parseInventoryDetails(plainBody) {
    const results = [];
    const sectionMatch = plainBody.match(/庫存明細([\s\S]+?)現股庫存市值總計/);
    if (!sectionMatch) return results;

    const inventoryStr = sectionMatch[1].replace(/<http[^>]+>|個股新聞|個股速覽/g, '');
    const regex = /(\d{4,6})\s*[\(（](.+?)[\)）]/g;
    let match;
    
    while ((match = regex.exec(inventoryStr)) !== null) {
        const [ , rawCode, stockName] = match;
        const remainingStr = inventoryStr.substring(match.index + match[0].length);
        const nextStockMatch = remainingStr.match(/\d{4,6}\s*[\(（]/);
        const extractArea = nextStockMatch ? remainingStr.substring(0, nextStockMatch.index) : remainingStr;

        const numbers = extractArea.match(/-?[\d,]+\.?\d*/g);
        if (numbers && numbers.length >= 8) {
            const nums = numbers.map(n => parseFloat(n.replace(/,/g, '')));
            // 直接擷取前8個數字，讓陣列合併更優雅
            results.push([formatStockCode(rawCode), stockName, ...nums.slice(0, 8)]);
        }
    }
    return results;
}

// ====== 工具函數 ======
function getOrCreateSheet(ss, sheetName, headers) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    }
    return sheet;
}

function getTradeHeaders() {
    return ['交易日', '來源', '交易類型', '買賣', '股票代碼', '股票名稱', '單價', '數量', '手續費', '交易稅', '淨收付', '收信時間', '郵件類型'];
}

function getInventoryHeaders() {
    return ['日期', '股票代碼', '股票名稱', '現股股數', '融資股數', '融資金額', '融券股數', '融券保證金', '融券擔保品', '收盤價', '現股庫存市值', '收信時間'];
}

function getOrCreateLabel(labelName) {
    return GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
}

function cleanupTradeSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TRADE_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const seen = new Set();
    const uniqueData = [];

    data.forEach(row => {
        while (row.length < 13) row.push(''); 
        if (row[12] === '成交回報') return; // 過濾掉舊的成交回報

        row[4] = formatStockCode(row[4]);
        const key = `${row[0]}|${row[4]}|${row[6]}|${row[7]}`;

        if (!seen.has(key)) {
            seen.add(key);
            uniqueData.push(row);
        }
    });
    
    if (uniqueData.length > 0) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
        if (sheet.getMaxRows() < uniqueData.length + 1) {
            sheet.insertRowsAfter(sheet.getMaxRows(), uniqueData.length + 1 - sheet.getMaxRows());
        }
        sheet.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
        sheet.getRange(2, 5, uniqueData.length, 1).setNumberFormat('@');
    }
}

function cleanupInventorySheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;

    const data = sheet.getDataRange().getValues();
    const seen = new Set();
    const uniqueData = [];

    // 從 i=1 開始，避開標題列
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        while (row.length < 12) row.push('');
        
        const key = `${row[0]}|${row[1]}`;
        if (!seen.has(key)) {
            seen.add(key);
            uniqueData.push(row);
        }
    }

    if (uniqueData.length > 0) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
        if (sheet.getMaxRows() < uniqueData.length + 1) {
            sheet.insertRowsAfter(sheet.getMaxRows(), uniqueData.length + 1 - sheet.getMaxRows());
        }
        sheet.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
    }
}

function createTimeTrigger() {
    ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
    ScriptApp.newTrigger('processAllEmails').timeBased().everyHours(12).create();
}

function sortAllSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    [TRADE_SHEET_NAME, INVENTORY_SHEET_NAME].forEach(name => {
        const sheet = ss.getSheetByName(name);
        if (sheet && sheet.getLastRow() > 1) {
            sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({column: 1, ascending: true});
        }
    });
    Logger.log("全部工作表排序完成！");
}
