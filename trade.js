// ====== 設定區 ======
var TRADE_SHEET_NAME = '交易明細';
var INVENTORY_SHEET_NAME = '庫存明細';
var PROCESSED_LABEL_DETAIL = '已處理-交易明細表';
var PROCESSED_LABEL_REPORT = '已處理-成交回報';

// ====== 主要處理函數 ======

/**
 * 處理所有郵件（交易明細表 + 成交回報資料）
 */
function processAllEmails() {
    processTradeDetailEmails();
    processTradeReportEmails();
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
            var stockCode = match[4];
            if (stockCode.length < 4) {
                stockCode = stockCode.padStart(4, '0');
            } else if (stockCode.length > 4 && stockCode.length < 6) {
                stockCode = stockCode.padStart(6, '0');
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
            var stockCode = match[4];

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
 * 測試處理結果
 */
function testProcess() {
    // 檢查交易明細表郵件
    var detailThreads = GmailApp.search('from:service@mailagent.pscnet.com.tw subject:交易明細表');
    Logger.log('交易明細表郵件: ' + detailThreads.length + ' 封');

    // 檢查成交回報郵件
    var reportThreads = GmailApp.search('from:service@mailagent.pscnet.com.tw subject:統一證券成交回報資料');
    Logger.log('成交回報資料郵件: ' + reportThreads.length + ' 封');

    // 檢查工作表
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tradeSheet = ss.getSheetByName('交易明細');
    var inventorySheet = ss.getSheetByName('庫存明細');

    Logger.log('交易明細工作表行數: ' + (tradeSheet ? tradeSheet.getLastRow() : '不存在'));
    Logger.log('庫存明細工作表行數: ' + (inventorySheet ? inventorySheet.getLastRow() : '不存在'));

    // 檢查標籤
    var label1 = GmailApp.getUserLabelByName('已處理-交易明細表');
    var label2 = GmailApp.getUserLabelByName('已處理-成交回報');

    Logger.log('已處理-交易明細表 標籤郵件數: ' + (label1 ? label1.getThreads().length : '標籤不存在'));
    Logger.log('已處理-成交回報 標籤郵件數: ' + (label2 ? label2.getThreads().length : '標籤不存在'));
}

/**
 * 檢查12月成交回報郵件
 */
function checkDecemberReports() {
    var threads = GmailApp.search('from:service@mailagent.pscnet.com.tw subject:統一證券成交回報資料 after:2025/12/01');

    Logger.log('12月成交回報郵件: ' + threads.length + ' 封');

    if (threads.length > 0) {
        var message = threads[0].getMessages()[0];
        var plainBody = message.getPlainBody();
        var date = message.getDate();

        Logger.log('最新郵件日期: ' + date);
        Logger.log('===== 郵件內容 =====');
        Logger.log(plainBody.substring(0, 2000));

        // 測試解析
        var trades = parseTradeReport(plainBody);
        Logger.log('===== 解析結果 =====');
        Logger.log('解析到 ' + trades.length + ' 筆交易');
        for (var i = 0; i < trades.length; i++) {
            Logger.log(JSON.stringify(trades[i]));
        }
    }

    // 檢查標籤狀態
    var label = GmailApp.getUserLabelByName('已處理-成交回報');
    if (label) {
        var labeledThreads = label.getThreads();
        Logger.log('已處理-成交回報 標籤: ' + labeledThreads.length + ' 封');
    } else {
        Logger.log('已處理-成交回報 標籤不存在');
    }
}

/**
 * 強制處理所有成交回報（忽略標籤）
 */
function forceProcessReports() {
    var threads = GmailApp.search('from:service@mailagent.pscnet.com.tw subject:統一證券成交回報資料');

    Logger.log('找到 ' + threads.length + ' 封成交回報郵件');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tradeSheet = ss.getSheetByName('交易明細');

    if (!tradeSheet) {
        Logger.log('錯誤：找不到交易明細工作表');
        return;
    }

    var label = getOrCreateLabel('已處理-成交回報');
    var totalTrades = 0;

    for (var i = 0; i < threads.length; i++) {
        var messages = threads[i].getMessages();

        for (var j = 0; j < messages.length; j++) {
            var message = messages[j];
            var plainBody = message.getPlainBody();
            var receivedDate = Utilities.formatDate(message.getDate(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

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

    Logger.log('共寫入 ' + totalTrades + ' 筆成交回報交易');
}

/**
 * 檢查工作表內容和缺漏
 */
function checkSheetAndMissing() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tradeSheet = ss.getSheetByName('交易明細');

    if (!tradeSheet) {
        Logger.log('找不到交易明細工作表');
        return;
    }

    var lastRow = tradeSheet.getLastRow();
    var lastCol = tradeSheet.getLastColumn();

    Logger.log('交易明細工作表: ' + lastRow + ' 行, ' + lastCol + ' 欄');

    // 顯示標題
    if (lastRow >= 1) {
        var headers = tradeSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        Logger.log('標題: ' + headers.join(' | '));
    }

    // 顯示最後幾筆資料
    if (lastRow >= 2) {
        Logger.log('===== 最後 5 筆資料 =====');
        var startRow = Math.max(2, lastRow - 4);
        var data = tradeSheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
        for (var i = 0; i < data.length; i++) {
            Logger.log('行 ' + (startRow + i) + ': ' + data[i].join(' | '));
        }
    }

    // 統計郵件類型分佈
    if (lastRow >= 2 && lastCol >= 13) {
        var allData = tradeSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        var typeCount = {};
        for (var i = 0; i < allData.length; i++) {
            var type = allData[i][lastCol - 1] || '無類型';
            typeCount[type] = (typeCount[type] || 0) + 1;
        }
        Logger.log('===== 郵件類型統計 =====');
        for (var t in typeCount) {
            Logger.log(t + ': ' + typeCount[t] + ' 筆');
        }
    }

    // 統計交易月份分佈
    Logger.log('===== 月份統計 =====');
    if (lastRow >= 2) {
        var allData = tradeSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        var monthCount = {};
        for (var i = 0; i < allData.length; i++) {
            var dateStr = String(allData[i][0]);
            var month = dateStr.substring(0, 7) || '無日期';
            monthCount[month] = (monthCount[month] || 0) + 1;
        }
        for (var m in monthCount) {
            Logger.log(m + ': ' + monthCount[m] + ' 筆');
        }
    }
}


/**
 * 1. 刪除舊資料（無郵件類型的）
 */
function deleteOldData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('交易明細');

    if (!sheet || sheet.getLastRow() < 2) {
        Logger.log('沒有資料需要刪除');
        return;
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // 從最後一行往前刪除（避免行號變動問題）
    var deletedCount = 0;
    for (var i = data.length - 1; i >= 0; i--) {
        var mailType = data[i][lastCol - 1];  // 最後一欄是郵件類型
        if (!mailType || mailType === '' || mailType === '無類型') {
            sheet.deleteRow(i + 2);  // +2 因為資料從第2行開始
            deletedCount++;
        }
    }

    Logger.log('已刪除 ' + deletedCount + ' 筆舊資料');
}

/**
 * 2. 修正股票代碼格式（補零）
 */
function fixStockCodes() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('交易明細');

    if (!sheet || sheet.getLastRow() < 2) return;

    var lastRow = sheet.getLastRow();
    var stockCodeCol = 5;  // 股票代碼在第5欄
    var range = sheet.getRange(2, stockCodeCol, lastRow - 1, 1);
    var values = range.getValues();

    var fixedCount = 0;
    for (var i = 0; i < values.length; i++) {
        var code = String(values[i][0]);
        // 如果是4位數或以下，補零到4位
        if (/^\d+$/.test(code) && code.length < 4) {
            values[i][0] = code.padStart(4, '0');
            fixedCount++;
        } else if (/^\d+$/.test(code) && code.length === 4) {
            values[i][0] = code;  // 保持為文字
        } else if (/^\d+$/.test(code) && code.length < 6) {
            values[i][0] = code.padStart(6, '0');
            fixedCount++;
        }
    }

    range.setValues(values);
    Logger.log('已修正 ' + fixedCount + ' 筆股票代碼');
}

/**
 * 3. 執行清理和修正
 */
function cleanupAndFix() {
    deleteOldData();
    fixStockCodes();
    Logger.log('清理完成！');
}

/**
 * 確認目前資料
 */
function verifyData() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('交易明細');

    var lastRow = sheet.getLastRow();
    Logger.log('目前共 ' + (lastRow - 1) + ' 筆資料');

    // 顯示前5筆
    if (lastRow >= 2) {
        Logger.log('===== 前 5 筆資料 =====');
        var data = sheet.getRange(2, 1, Math.min(5, lastRow - 1), 13).getValues();
        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            Logger.log((i + 1) + '. ' + row[0] + ' | ' + row[4] + ' | ' + row[5] + ' | ' + row[3] + ' | ' + row[6] + ' x ' + row[7] + ' | ' + row[12]);
        }
    }

    // 月份統計
    Logger.log('===== 月份統計 =====');
    var allDates = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var monthCount = {};
    for (var i = 0; i < allDates.length; i++) {
        var d = new Date(allDates[i][0]);
        var key = d.getFullYear() + '/' + String(d.getMonth() + 1).padStart(2, '0');
        monthCount[key] = (monthCount[key] || 0) + 1;
    }
    var sortedMonths = Object.keys(monthCount).sort();
    for (var i = 0; i < sortedMonths.length; i++) {
        Logger.log(sortedMonths[i] + ': ' + monthCount[sortedMonths[i]] + ' 筆');
    }
}

/**
 * 修正股票代碼欄位格式（設為文字）
 */
function fixStockCodeFormat() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('交易明細');

    if (!sheet || sheet.getLastRow() < 2) return;

    var lastRow = sheet.getLastRow();
    var stockCodeCol = 5;  // 股票代碼在第5欄

    // 設定整個欄位為純文字格式
    var range = sheet.getRange(2, stockCodeCol, lastRow - 1, 1);
    range.setNumberFormat('@');  // @ 表示純文字

    // 重新寫入值（補零）
    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
        var code = String(values[i][0]);
        if (/^\d+$/.test(code)) {
            if (code.length <= 4) {
                values[i][0] = code.padStart(4, '0');
            } else if (code.length < 6) {
                values[i][0] = code.padStart(6, '0');
            }
        }
    }

    range.setValues(values);
    Logger.log('已修正 ' + values.length + ' 筆股票代碼格式');
}

/**
 * 修正庫存明細的股票代碼格式
 */
function fixInventoryStockCodeFormat() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('庫存明細');

    if (!sheet || sheet.getLastRow() < 2) return;

    var lastRow = sheet.getLastRow();
    var stockCodeCol = 2;  // 庫存明細的股票代碼在第2欄

    var range = sheet.getRange(2, stockCodeCol, lastRow - 1, 1);
    range.setNumberFormat('@');

    var values = range.getValues();
    for (var i = 0; i < values.length; i++) {
        var code = String(values[i][0]);
        if (/^\d+$/.test(code)) {
            if (code.length <= 4) {
                values[i][0] = code.padStart(4, '0');
            } else if (code.length < 6) {
                values[i][0] = code.padStart(6, '0');
            }
        }
    }

    range.setValues(values);
    Logger.log('已修正庫存明細 ' + values.length + ' 筆股票代碼格式');
}

/**
 * 修正所有工作表的股票代碼
 */
function fixAllStockCodes() {
    fixStockCodeFormat();           // 交易明細
    fixInventoryStockCodeFormat();  // 庫存明細
    Logger.log('所有股票代碼已修正！');
}


/**
 * 將庫存明細分成不同股票的工作表
 */
function splitInventoryByStock() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = ss.getSheetByName('庫存明細');

    if (!sourceSheet || sourceSheet.getLastRow() < 2) return;

    var headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
    var data = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn()).getValues();

    // 按股票代碼分組
    var stockGroups = {};
    for (var i = 0; i < data.length; i++) {
        var stockCode = String(data[i][1]);  // 第2欄是股票代碼
        if (!stockGroups[stockCode]) {
            stockGroups[stockCode] = [];
        }
        stockGroups[stockCode].push(data[i]);
    }

    // 為每個股票建立工作表
    for (var code in stockGroups) {
        var sheetName = '庫存明細-' + code;
        var sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
        }

        var rows = stockGroups[code];
        if (rows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
        }

        Logger.log('已建立 ' + sheetName + '：' + rows.length + ' 筆');
    }

    Logger.log('分割完成！');
}

/**
 * 修正庫存明細分割工作表的格式
 */
function fixSplitInventoryFormat() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = ['庫存明細-0050', '庫存明細-006208', '庫存明細-50', '庫存明細-6208'];

    for (var s = 0; s < sheetNames.length; s++) {
        var sheet = ss.getSheetByName(sheetNames[s]);
        if (!sheet || sheet.getLastRow() < 2) continue;

        var lastRow = sheet.getLastRow();

        // 修正日期格式（第1欄）
        var dateRange = sheet.getRange(2, 1, lastRow - 1, 1);
        var dates = dateRange.getValues();
        for (var i = 0; i < dates.length; i++) {
            var d = new Date(dates[i][0]);
            if (!isNaN(d.getTime())) {
                var year = d.getFullYear();
                var month = String(d.getMonth() + 1).padStart(2, '0');
                var day = String(d.getDate()).padStart(2, '0');
                dates[i][0] = year + '/' + month + '/' + day;
            }
        }
        dateRange.setValues(dates);

        // 修正股票代碼格式（第2欄）
        var codeRange = sheet.getRange(2, 2, lastRow - 1, 1);
        codeRange.setNumberFormat('@');  // 設為純文字
        var codes = codeRange.getValues();
        for (var i = 0; i < codes.length; i++) {
            var code = String(codes[i][0]);
            if (/^\d+$/.test(code)) {
                if (code.length <= 4) {
                    codes[i][0] = code.padStart(4, '0');
                } else if (code.length < 6) {
                    codes[i][0] = code.padStart(6, '0');
                }
            }
        }
        codeRange.setValues(codes);

        Logger.log('已修正 ' + sheetNames[s] + ' 格式');
    }

    // 重新命名工作表（如果名稱不正確）
    var sheet50 = ss.getSheetByName('庫存明細-50');
    if (sheet50) {
        sheet50.setName('庫存明細-0050');
        Logger.log('已重新命名為 庫存明細-0050');
    }

    var sheet6208 = ss.getSheetByName('庫存明細-6208');
    if (sheet6208) {
        sheet6208.setName('庫存明細-006208');
        Logger.log('已重新命名為 庫存明細-006208');
    }

    Logger.log('格式修正完成！');
}

/**
 * 強制修正日期格式
 */
function fixDateFormat() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetNames = ['庫存明細-0050', '庫存明細-006208'];

    for (var s = 0; s < sheetNames.length; s++) {
        var sheet = ss.getSheetByName(sheetNames[s]);
        if (!sheet || sheet.getLastRow() < 2) continue;

        var lastRow = sheet.getLastRow();
        var dateRange = sheet.getRange(2, 1, lastRow - 1, 1);

        // 先設定為純文字格式
        dateRange.setNumberFormat('@');

        var dates = dateRange.getValues();
        for (var i = 0; i < dates.length; i++) {
            var val = dates[i][0];
            var d;

            // 嘗試解析不同格式的日期
            if (val instanceof Date) {
                d = val;
            } else {
                d = new Date(val);
            }

            if (!isNaN(d.getTime())) {
                var year = d.getFullYear();
                var month = String(d.getMonth() + 1).padStart(2, '0');
                var day = String(d.getDate()).padStart(2, '0');
                dates[i][0] = year + '/' + month + '/' + day;
            } else {
                Logger.log('無法解析日期: ' + val);
            }
        }

        dateRange.setValues(dates);
        Logger.log('已修正 ' + sheetNames[s] + ' 的日期格式，共 ' + dates.length + ' 筆');
    }

    Logger.log('日期格式修正完成！');
}