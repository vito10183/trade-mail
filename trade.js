// ====== 設定區 ======
var TRADE_SHEET_NAME = '交易明細';
var INVENTORY_SHEET_NAME = '庫存明細';
var PROCESSED_LABEL_DETAIL = '已處理-交易明細表';
var PROCESSED_LABEL_REPORT = '已處理-成交回報';

// ====== 主要處理函數 ======

/**
 * 處理所有郵件（交易明細表 + 成交回報資料）
 * 處理完成後自動分割庫存明細並清理交易明細
 */
function processAllEmails() {
    processTradeDetailEmails();
    processTradeReportEmails();
    cleanupTradeSheet();          // 去除重複並修正股票代碼
    splitInventoryByStock();      // 分割庫存明細
}

/**
 * 處理交易明細表郵件
 */
function processTradeDetailEmails() {
    var threads = GmailApp.search('from:service@mailagent.pscnet.com.tw subject:交易明細表 -label:' + PROCESSED_LABEL_DETAIL);

    if (threads.length === 0) {
        Logger.log('沒有新的交易明細表郵件');
        return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tradeSheet = getOrCreateSheet(ss, TRADE_SHEET_NAME, getTradeHeaders());
    var inventorySheet = getOrCreateSheet(ss, INVENTORY_SHEET_NAME, getInventoryHeaders());
    var label = getOrCreateLabel(PROCESSED_LABEL_DETAIL);

    var totalTrades = 0;
    var totalInventory = 0;

    for (var i = 0; i < threads.length; i++) {
        var messages = threads[i].getMessages();

        for (var j = 0; j < messages.length; j++) {
            var message = messages[j];
            var plainBody = message.getPlainBody();
            var receivedDate = Utilities.formatDate(message.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

            var tradeDateMatch = plainBody.match(/交易日：(\d{4}\/\d{2}\/\d{2})/);
            var tradeDate = tradeDateMatch ? tradeDateMatch[1] : '';

            var lines = plainBody.split('\n').map(function (l) { return l.trim(); });

            // 解析交易明細
            var trades = parseTradeDetailsV2(lines);
            for (var k = 0; k < trades.length; k++) {
                var row = [tradeDate].concat(trades[k]).concat([receivedDate, '交易明細表']);
                tradeSheet.appendRow(row);
                totalTrades++;
            }

            // 解析庫存明細
            var inventory = parseInventoryDetailsV2(lines);
            for (var k = 0; k < inventory.length; k++) {
                var row = [tradeDate].concat(inventory[k]).concat([receivedDate]);
                inventorySheet.appendRow(row);
                totalInventory++;
            }

            message.markRead();
        }

        threads[i].addLabel(label);
    }

    Logger.log('交易明細表: ' + totalTrades + ' 筆交易, ' + totalInventory + ' 筆庫存');
}

/**
 * 處理成交回報資料郵件
 */
function processTradeReportEmails() {
    var threads = GmailApp.search('from:service@mailagent.pscnet.com.tw subject:統一證券成交回報資料 -label:' + PROCESSED_LABEL_REPORT);

    if (threads.length === 0) {
        Logger.log('沒有新的成交回報資料郵件');
        return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tradeSheet = getOrCreateSheet(ss, TRADE_SHEET_NAME, getTradeHeaders());
    var label = getOrCreateLabel(PROCESSED_LABEL_REPORT);

    var totalTrades = 0;

    for (var i = 0; i < threads.length; i++) {
        var messages = threads[i].getMessages();

        for (var j = 0; j < messages.length; j++) {
            var message = messages[j];
            var plainBody = message.getPlainBody();
            var receivedDate = Utilities.formatDate(message.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

            // 解析成交回報
            var trades = parseTradeReport(plainBody);

            for (var k = 0; k < trades.length; k++) {
                var row = trades[k].concat([receivedDate, '成交回報']);
                tradeSheet.appendRow(row);
                totalTrades++;
            }

            message.markRead();
        }

        threads[i].addLabel(label);
    }

    Logger.log('成交回報資料: ' + totalTrades + ' 筆交易');
}

/**
 * 解析成交回報資料郵件
 */
/**
 * 解析成交回報資料郵件（修正版 - 保留股票代碼格式）
 */
function parseTradeReport(plainBody) {
    var results = [];
    var lines = plainBody.split('\n');

    var dateMatch = plainBody.match(/於\s*(\d{4}-\d{2}-\d{2})\s*交易/);
    var tradeDate = dateMatch ? dateMatch[1].replace(/-/g, '/') : '';

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();

        var match = line.match(/^(\d+)\s+(\d+)\s+(\w+)\s+(\d+)\s+(.+?)\s+(金贏島|金麗島)\s+(買進|賣出)\s+(普通|融資|融券)\s+(盤中零股|整股|盤後零股|盤後定價)\s+([\d.]+)\s+(\d+)\s+(\d{2}:\d{2}:\d{2})\s+(\d{4}-\d{2}-\d{2})/);

        if (match) {
            // 保留股票代碼為字串，補零到4位或6位
            var rawCode = match[4];
            var stockCode;
            if (rawCode === '50' || parseInt(rawCode) === 50) {
                stockCode = '0050';
            } else if (rawCode === '6208' || parseInt(rawCode) === 6208) {
                stockCode = '006208';
            } else if (rawCode.length <= 4) {
                stockCode = rawCode.padStart(4, '0');
            } else {
                stockCode = rawCode.padStart(6, '0');
            }

            results.push([
                tradeDate,
                '電子單',
                match[8],
                match[7],
                stockCode,           // 保留字串格式
                match[5],
                parseFloat(match[10]),
                parseInt(match[11]),
                0,
                0,
                0
            ]);
        }
    }

    return results;
}

/**
 * 解析交易明細 V2
 */
function parseTradeDetailsV2(lines) {
    var results = [];

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];
        var match = line.match(/^(電子單|語音單|臨櫃|APP|網路單)\s+(普|資|券)\s+(買|賣)\s+(\d{4,6})$/);

        if (match && i + 1 < lines.length) {
            var source = match[1];
            var tradeType = match[2];
            var buySell = match[3];
            var rawCode = match[4];
            var stockCode;
            if (rawCode === '50' || parseInt(rawCode) === 50) {
                stockCode = '0050';
            } else if (rawCode === '6208' || parseInt(rawCode) === 6208) {
                stockCode = '006208';
            } else if (rawCode.length <= 4) {
                stockCode = rawCode.padStart(4, '0');
            } else {
                stockCode = rawCode.padStart(6, '0');
            }

            var nextLine = lines[i + 1];
            var nameMatch = nextLine.match(/^\((.+?)\)\s+(.+)/);

            if (nameMatch) {
                var stockName = nameMatch[1];
                var numbersStr = nameMatch[2];
                var numbers = numbersStr.match(/-?[\d,]+\.?\d*/g);

                if (numbers && numbers.length >= 4) {
                    var numericValues = numbers.map(function (n) {
                        return parseFloat(n.replace(/,/g, ''));
                    });

                    results.push([
                        source,
                        tradeType,
                        buySell,
                        stockCode,
                        stockName,
                        numericValues[0],
                        numericValues[1],
                        numericValues[2],
                        numericValues[3],
                        numericValues[numericValues.length - 1]
                    ]);
                }
            }
        }
    }

    return results;
}

/**
 * 解析庫存明細 V2
 */
function parseInventoryDetailsV2(lines) {
    var results = [];
    var inInventorySection = false;

    for (var i = 0; i < lines.length; i++) {
        var line = lines[i];

        if (line.indexOf('庫存明細') >= 0) {
            inInventorySection = true;
            continue;
        }

        if (line.indexOf('現股庫存市值總計') >= 0 || line.indexOf('市場公告') >= 0) {
            inInventorySection = false;
            continue;
        }

        if (inInventorySection && /^\d{4,6}$/.test(line)) {
            var stockCode = line;
            var stockName = '';
            var allNumbers = [];

            for (var j = i + 1; j < Math.min(i + 8, lines.length); j++) {
                var nextLine = lines[j];

                if (nextLine.indexOf('<http') >= 0) {
                    var afterUrl = nextLine.replace(/<http[^>]+>/g, '');
                    var nums = afterUrl.match(/[\d,]+\.?\d*/g);
                    if (nums) {
                        for (var n = 0; n < nums.length; n++) {
                            var num = nums[n].replace(/,/g, '');
                            if (num.length > 0 && !isNaN(parseFloat(num))) {
                                allNumbers.push(parseFloat(num));
                            }
                        }
                    }
                    continue;
                }

                var nameMatch = nextLine.match(/^\((.+?)\)/);
                if (nameMatch) {
                    stockName = nameMatch[1];
                    continue;
                }

                if (/^[\d\s.,\-]+$/.test(nextLine)) {
                    var nums = nextLine.match(/-?[\d,]+\.?\d*/g);
                    if (nums) {
                        for (var n = 0; n < nums.length; n++) {
                            var num = nums[n].replace(/,/g, '');
                            if (num.length > 0 && !isNaN(parseFloat(num))) {
                                allNumbers.push(parseFloat(num));
                            }
                        }
                    }
                }

                if (allNumbers.length >= 8) break;
            }

            if (stockName && allNumbers.length >= 8) {
                results.push([
                    stockCode, stockName,
                    allNumbers[0], allNumbers[1], allNumbers[2], allNumbers[3],
                    allNumbers[4], allNumbers[5], allNumbers[6], allNumbers[7]
                ]);
            }
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

/**
 * 重設並處理所有郵件
 */
function resetAndProcessAll() {
    // 清除標籤
    var labels = [PROCESSED_LABEL_DETAIL, PROCESSED_LABEL_REPORT];
    for (var l = 0; l < labels.length; l++) {
        var label = GmailApp.getUserLabelByName(labels[l]);
        if (label) {
            var threads = label.getThreads();
            for (var i = 0; i < threads.length; i++) {
                threads[i].removeLabel(label);
            }
            Logger.log('已移除 ' + threads.length + ' 封 ' + labels[l] + ' 標籤');
        }
    }

    // 清空工作表
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    clearSheet(ss, TRADE_SHEET_NAME);
    clearSheet(ss, INVENTORY_SHEET_NAME);

    // 重新處理
    processAllEmails();
}

function clearSheet(ss, sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet && sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
        Logger.log('已清空 ' + sheetName);
    }
}

/**
 * 清理交易明細：去除重複資料並修正股票代碼格式
 */
function cleanupTradeSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TRADE_SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
        Logger.log('交易明細無資料需要清理');
        return;
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // 建立唯一鍵來識別重複
    var seen = {};
    var uniqueData = [];
    var duplicateCount = 0;

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var tradeDate = row[0];
        var stockCode = String(row[4]);
        var price = row[6];
        var qty = row[7];
        var mailType = row[12];

        // 修正股票代碼格式
        var rawCode = stockCode.replace(/^0+/, '');
        if (rawCode === '50') {
            row[4] = '0050';
        } else if (rawCode === '6208') {
            row[4] = '006208';
        } else if (rawCode.length <= 4) {
            row[4] = rawCode.padStart(4, '0');
        } else {
            row[4] = rawCode.padStart(6, '0');
        }

        // 建立唯一鍵（日期+股票+價格+數量）
        var key = tradeDate + '|' + row[4] + '|' + price + '|' + qty;

        if (seen[key]) {
            // 已存在，檢查是否應該保留這筆
            // 優先保留「交易明細表」（資料較完整）
            if (mailType === '交易明細表') {
                // 用交易明細表覆蓋成交回報
                for (var j = 0; j < uniqueData.length; j++) {
                    var existKey = uniqueData[j][0] + '|' + uniqueData[j][4] + '|' + uniqueData[j][6] + '|' + uniqueData[j][7];
                    if (existKey === key && uniqueData[j][12] === '成交回報') {
                        uniqueData[j] = row;
                        break;
                    }
                }
            }
            duplicateCount++;
        } else {
            seen[key] = true;
            uniqueData.push(row);
        }
    }

    // 清空並重新寫入
    if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
    }

    if (uniqueData.length > 0) {
        sheet.getRange(2, 1, uniqueData.length, lastCol).setValues(uniqueData);

        // 設定股票代碼欄為純文字格式
        var codeRange = sheet.getRange(2, 5, uniqueData.length, 1);
        codeRange.setNumberFormat('@');
    }

    Logger.log('交易明細清理完成：移除 ' + duplicateCount + ' 筆重複，保留 ' + uniqueData.length + ' 筆');
}

/**
 * 建立定時觸發器 - 每12小時
 */
function createTimeTrigger() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }

    ScriptApp.newTrigger('processAllEmails')
        .timeBased()
        .everyHours(12)
        .create();

    Logger.log('已建立觸發器：每 12 小時執行 processAllEmails');
}

/**
 * 將庫存明細分割成 0050 和 006208 兩個工作表
 */
function splitInventoryByStock() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName('庫存明細');

    if (!sourceSheet || sourceSheet.getLastRow() < 2) {
        Logger.log('庫存明細無資料');
        return;
    }

    var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    var data = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();

    // 定義要分割的股票
    var targetStocks = ['0050', '006208'];
    var stockData = { '0050': [], '006208': [] };

    // 分類資料
    for (var i = 0; i < data.length; i++) {
        var rawCode = String(data[i][1]).replace(/^0+/, '');  // 移除前導零
        var stockCode;

        // 判斷是 4 位數股票還是 6 位數股票
        if (rawCode === '50' || rawCode.length <= 2) {
            stockCode = '0050';
        } else if (rawCode === '6208') {
            stockCode = '006208';
        } else if (rawCode.length <= 4) {
            stockCode = rawCode.padStart(4, '0');
        } else {
            stockCode = rawCode.padStart(6, '0');
        }

        if (stockCode === '0050') {
            stockData['0050'].push(data[i]);
        } else if (stockCode === '006208') {
            stockData['006208'].push(data[i]);
        }
    }

    // 為每個股票建立/更新工作表
    for (var s = 0; s < targetStocks.length; s++) {
        var code = targetStocks[s];
        var sheetName = '庫存明細-' + code;
        var sheet = ss.getSheetByName(sheetName);

        // 如果工作表不存在，建立新的
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        } else {
            // 清空舊資料（保留標題）
            if (sheet.getLastRow() > 1) {
                sheet.deleteRows(2, sheet.getLastRow() - 1);
            }
        }

        // 寫入資料
        var rows = stockData[code];
        if (rows.length > 0) {
            sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

            // 設定股票代碼欄為純文字格式
            var codeRange = sheet.getRange(2, 2, rows.length, 1);
            codeRange.setNumberFormat('@');
        }

        Logger.log(sheetName + '：' + rows.length + ' 筆資料');
    }

    Logger.log('庫存明細分割完成！');
}
