// ====== 設定區 ======
var TRADE_SHEET_NAME = '交易明細';
var INVENTORY_SHEET_NAME = '庫存明細';
var PROCESSED_LABEL_DETAIL = '已處理-交易明細表';
var PROCESSED_LABEL_REPORT = '已處理-成交回報';


// ====== 每日自動處理函數 ======
function processAllEmails() {
    processTradeDetailEmails();
    processTradeReportEmails();
    cleanupTradeSheet();  
    cleanupInventorySheet();        
    splitInventoryByStock();  
    sortAllSheets();       
}

function processTradeDetailEmails() {
    // 【破案關鍵】捨棄特定的 from: 信箱，改用標題與關鍵字搜尋
    var threads = GmailApp.search('subject:交易明細表 "統一綜合證券" -label:' + PROCESSED_LABEL_DETAIL);
    if (threads.length === 0) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tradeSheet = getOrCreateSheet(ss, TRADE_SHEET_NAME, getTradeHeaders());
    var inventorySheet = getOrCreateSheet(ss, INVENTORY_SHEET_NAME, getInventoryHeaders());
    var label = getOrCreateLabel(PROCESSED_LABEL_DETAIL);

    for (var i = 0; i < threads.length; i++) {
        var messages = threads[i].getMessages();
        for (var j = 0; j < messages.length; j++) {
            var message = messages[j];
            var plainBody = message.getPlainBody();
            var receivedDate = Utilities.formatDate(message.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

            var tradeDateMatch = plainBody.match(/交易日\s*[:：]\s*(\d{4}\/\d{2}\/\d{2})/);
            var tradeDate = tradeDateMatch ? tradeDateMatch[1] : '';

            var trades = parseTradeDetailsV2(plainBody);
            for (var k = 0; k < trades.length; k++) {
                tradeSheet.appendRow([tradeDate].concat(trades[k]).concat([receivedDate, '交易明細表']));
            }

            var inventory = parseInventoryDetailsV2(plainBody);
            for (var k = 0; k < inventory.length; k++) {
                inventorySheet.appendRow([tradeDate].concat(inventory[k]).concat([receivedDate]));
            }

            message.markRead();
        }
        threads[i].addLabel(label);
    }
}

function processTradeReportEmails() {
    // 【破案關鍵】捨棄特定的 from: 信箱，改用標題搜尋
    var threads = GmailApp.search('subject:統一證券成交回報資料 -label:' + PROCESSED_LABEL_REPORT);
    if (threads.length === 0) return;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tradeSheet = getOrCreateSheet(ss, TRADE_SHEET_NAME, getTradeHeaders());
    var label = getOrCreateLabel(PROCESSED_LABEL_REPORT);

    for (var i = 0; i < threads.length; i++) {
        var messages = threads[i].getMessages();
        for (var j = 0; j < messages.length; j++) {
            var message = messages[j];
            var plainBody = message.getPlainBody();
            var receivedDate = Utilities.formatDate(message.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

            var trades = parseTradeReport(plainBody);
            for (var k = 0; k < trades.length; k++) {
                tradeSheet.appendRow(trades[k].concat([receivedDate, '成交回報']));
            }
            message.markRead();
        }
        threads[i].addLabel(label);
    }
}

function formatStockCode(rawCode) {
    if (rawCode === '50' || parseInt(rawCode, 10) === 50) return '0050';
    if (rawCode === '6208' || parseInt(rawCode, 10) === 6208) return '006208';
    return rawCode.length <= 4 ? rawCode.padStart(4, '0') : rawCode.padStart(6, '0');
}

function parseTradeReport(plainBody) {
    var results = [];
    var lines = plainBody.split('\n');
    var dateMatch = plainBody.match(/於\s*(\d{4}-\d{2}-\d{2})\s*交易/);
    var tradeDate = dateMatch ? dateMatch[1].replace(/-/g, '/') : '';

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        var match = line.match(/^(\d+)\s+(\d+)\s+(\w+)\s+(\d+)\s+(.+?)\s+(金贏島|金麗島)\s+(買進|賣出)\s+(普通|融資|融券)\s+(盤中零股|整股|盤後零股|盤後定價)\s+([\d.]+)\s+(\d+)\s+(\d{2}:\d{2}:\d{2})\s+(\d{4}-\d{2}-\d{2})/);
        if (match) {
            results.push([tradeDate, '電子單', match[8], match[7], formatStockCode(match[4]), match[5], parseFloat(match[10]), parseInt(match[11]), 0, 0, 0]);
        }
    }
    return results;
}

/**
 * 【終極版】解析交易明細：不逐行看，直接在整封信找關鍵字
 */
function parseTradeDetailsV2(plainBody) {
    var results = [];
    var regex = /(電子單|語音單|臨櫃|APP|智慧單|網路單)\s+(普|資|券)\s+(買|賣)\s+(\d{4,6})\s*[\(（](.+?)[\)）]/g;
    var match;
    
    while ((match = regex.exec(plainBody)) !== null) {
        var source = match[1];
        var tradeType = match[2];
        var buySell = match[3];
        var stockCode = formatStockCode(match[4]);
        var stockName = match[5];
        
        var remainingStr = plainBody.substring(match.index + match[0].length);
        
        var nextMatch = remainingStr.match(/(?:電子單|語音單|臨櫃|APP|智慧單|網路單)\s+(?:普|資|券)|合計/);
        var extractArea = nextMatch ? remainingStr.substring(0, nextMatch.index) : remainingStr;
        
        var numbers = extractArea.match(/-?[\d,]+\.?\d*/g);
        
        if (numbers && numbers.length >= 10) { 
            var numericValues = numbers.map(function(n) { return parseFloat(n.replace(/,/g, '')); });
            results.push([
                source, tradeType, buySell, stockCode, stockName,
                numericValues[0], numericValues[1], numericValues[2], numericValues[3], numericValues[12] || numericValues[numericValues.length - 1]
            ]);
        }
    }
    return results;
}

/**
 * 【終極版】解析庫存明細：不逐行看，直接在整封信找關鍵字
 */
function parseInventoryDetailsV2(plainBody) {
    var results = [];
    var inventorySectionMatch = plainBody.match(/庫存明細([\s\S]+?)現股庫存市值總計/);
    if (!inventorySectionMatch) return results;

    var inventoryStr = inventorySectionMatch[1]
                        .replace(/<http[^>]+>/g, '')
                        .replace(/個股新聞/g, '')
                        .replace(/個股速覽/g, '');

    var regex = /(\d{4,6})\s*[\(（](.+?)[\)）]/g;
    var match;
    
    while ((match = regex.exec(inventoryStr)) !== null) {
        var stockCode = formatStockCode(match[1]);
        var stockName = match[2];
        
        var remainingStr = inventoryStr.substring(match.index + match[0].length);
        
        var nextStockMatch = remainingStr.match(/\d{4,6}\s*[\(（]/);
        var extractArea = nextStockMatch ? remainingStr.substring(0, nextStockMatch.index) : remainingStr;

        var numbers = extractArea.match(/-?[\d,]+\.?\d*/g);
        if (numbers && numbers.length >= 8) {
            var numericValues = numbers.map(function(n) { return parseFloat(n.replace(/,/g, '')); });
            results.push([
                stockCode, stockName,
                numericValues[0], numericValues[1], numericValues[2], numericValues[3],
                numericValues[4], numericValues[5], numericValues[6], numericValues[7]
            ]);
        }
    }
    return results;
}

// ====== 工具函數 ======
function getOrCreateSheet(ss, sheetName, headers) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
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
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
        label = GmailApp.createLabel(labelName);
    }
    return label;
}

function cleanupTradeSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TRADE_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    var seen = {};
    var uniqueData = [];

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        while (row.length < 13) row.push(''); // 確保有13欄
        
        row[4] = formatStockCode(String(row[4]).replace(/^0+/, ''));
        var key = row[0] + '|' + row[4] + '|' + row[6] + '|' + row[7];

        if (seen[key]) {
            if (row[12] === '交易明細表') {
                for (var j = 0; j < uniqueData.length; j++) {
                    var existKey = uniqueData[j][0] + '|' + uniqueData[j][4] + '|' + uniqueData[j][6] + '|' + uniqueData[j][7];
                    if (existKey === key && uniqueData[j][12] === '成交回報') {
                        uniqueData[j] = row;
                        break;
                    }
                }
            }
        } else {
            seen[key] = true;
            uniqueData.push(row);
        }
    }
    
    if (uniqueData.length > 0) {
        // 【防彈修正】改用 clearContent 只清除文字，保留格子
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
        
        // 【防彈修正】如果試算表剩下的格子不夠放新資料，自動新增足夠的空白列
        if (sheet.getMaxRows() < uniqueData.length + 1) {
            sheet.insertRowsAfter(sheet.getMaxRows(), uniqueData.length + 1 - sheet.getMaxRows());
        }
        
        sheet.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
        sheet.getRange(2, 5, uniqueData.length, 1).setNumberFormat('@');
    }
}

function cleanupInventorySheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return;

    var data = sheet.getDataRange().getValues();
    var uniqueData = [];
    var seen = {};

    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        while (row.length < 12) row.push(''); // 確保有12欄
        
        var key = row[0] + '|' + row[1];
        if (!seen[key]) {
            seen[key] = true;
            uniqueData.push(row);
        }
    }

    if (uniqueData.length > 0) {
        // 【防彈修正】清除文字，保留格子
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
        
        // 【防彈修正】自動擴建不足的列數
        if (sheet.getMaxRows() < uniqueData.length + 1) {
            sheet.insertRowsAfter(sheet.getMaxRows(), uniqueData.length + 1 - sheet.getMaxRows());
        }
        
        sheet.getRange(2, 1, uniqueData.length, uniqueData[0].length).setValues(uniqueData);
    }
}

function splitInventoryByStock() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName(INVENTORY_SHEET_NAME);
    if (!sourceSheet || sourceSheet.getLastRow() < 2) return;

    var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    var data = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();
    var targetStocks = ['0050', '006208'];
    var stockData = { '0050': [], '006208': [] };

    for (var i = 0; i < data.length; i++) {
        var stockCode = formatStockCode(String(data[i][1]).replace(/^0+/, ''));
        if (stockData[stockCode]) stockData[stockCode].push(data[i]);
    }

    for (var s = 0; s < targetStocks.length; s++) {
        var code = targetStocks[s];
        var sheetName = '庫存明細-' + code;
        var sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        } else if (sheet.getLastRow() > 1) {
            // 【防彈修正】清除文字，保留格子
            sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getMaxColumns()).clearContent();
        }

        var rows = stockData[code];
        if (rows.length > 0) {
            // 【防彈修正】自動擴建不足的列數
            if (sheet.getMaxRows() < rows.length + 1) {
                sheet.insertRowsAfter(sheet.getMaxRows(), rows.length + 1 - sheet.getMaxRows());
            }
            sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
            sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
        }
    }
}


function createTimeTrigger() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
    ScriptApp.newTrigger('processAllEmails').timeBased().everyHours(12).create();
}
/**
 * 自動將所有工作表依照「日期」由舊到新排序
 */
function sortAllSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    // 這裡放入你所有的工作表名稱
    var sheetNames = [TRADE_SHEET_NAME, INVENTORY_SHEET_NAME, '庫存明細-0050', '庫存明細-006208'];
    
    for (var i = 0; i < sheetNames.length; i++) {
        var sheet = ss.getSheetByName(sheetNames[i]);
        if (sheet && sheet.getLastRow() > 1) {
            // 扣除第一行的標題，將其餘資料範圍針對 A 欄 (第 1 欄的日期) 進行 A-Z 遞增排序
            var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
            range.sort({column: 1, ascending: true});
        }
    }
    Logger.log("全部工作表排序完成！");
}
