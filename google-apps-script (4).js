/**
 * 支出記錄系統 - Google Apps Script 後端
 * 
 * 此腳本用於處理來自前端的HTTP請求，並將支出數據寫入Google Sheets
 * 
 * 部署說明：
 * 1. 打開 https://script.google.com/
 * 2. 創建新項目，將此代碼複製貼上
 * 3. 點擊「部署」> 「新增部署作業」
 * 4. 類型選擇「網頁應用程式」
 * 5. 執行身分：選擇「我」
 * 6. 存取權限：選擇「任何人」
 * 7. 點擊「部署」並複製Web App URL
 * 8. 將URL貼到前端HTML的CONFIG.SCRIPT_URL中
 */

// 配置常數
const CONFIG = {
  // Google Sheets ID（從URL中獲取）
  SHEET_ID: '19iLhAAWJsogeTqfQciJyyiv9gBEB31MdRpGHWarfPxE',
  
  // 工作表名稱
  SHEET_NAME: 'Sheet1',
  
  // 表頭欄位
  HEADERS: ['日期', '項目', '類別', '金額', '支付方式', '備註', '記錄時間'],
  
  // 允許的來源域名（CORS設定）
  // 在 Apps Script Web App 部署時，如果存取權限設置為「任何人」，
  // 則 Apps Script 會自動處理 CORS 相關的響應頭，無需手動設置。
  // 此處的 ALLOWED_ORIGINS 僅作為備註，實際不會在 createResponse 中使用。
  ALLOWED_ORIGINS: ['*'] 
};

/**
 * 處理HTTP GET請求
 * 用於測試API是否正常運作
 */
function doGet(e) {
  try {
    return createResponse({
      success: true,
      message: '支出記錄API運行正常',
      timestamp: new Date().toISOString(),
      version: '1.0.0'
    });
  } catch (error) {
    console.error('GET請求處理錯誤:', error);
    return createResponse({
      success: false,
      error: error.toString()
    }, 500);
  }
}

/**
 * 處理HTTP POST請求
 * 主要用於接收前端提交的支出數據
 */
function doPost(e) {
  try {
    // 解析請求數據
    const requestData = parseRequestData(e);
    console.log('收到請求數據:', requestData);
    
    // 驗證請求數據
    if (!validateRequestData(requestData)) {
      return createResponse({
        success: false,
        error: '請求數據格式不正確'
      }, 400);
    }
    
    // 根據操作類型處理請求
    switch (requestData.action) {
      case 'addExpense':
        return handleAddExpense(requestData);
      
      case 'initSheet':
        return handleInitSheet(requestData);
      
      default:
        return createResponse({
          success: false,
          error: '不支援的操作類型'
        }, 400);
    }
    
  } catch (error) {
    console.error('POST請求處理錯誤:', error);
    return createResponse({
      success: false,
      error: '伺服器內部錯誤: ' + error.toString()
    }, 500);
  }
}

/**
 * 解析請求數據
 */
function parseRequestData(e) {
  try {
    // 嘗試解析JSON數據
    if (e.postData && e.postData.contents) {
      return JSON.parse(e.postData.contents);
    }
    
    // 如果沒有JSON數據，返回空對象
    return {};
    
  } catch (error) {
    console.error('解析請求數據失敗:', error);
    throw new Error('請求數據格式不正確');
  }
}

/**
 * 驗證請求數據
 */
function validateRequestData(data) {
  // 檢查是否有action字段
  if (!data.action) {
    return false;
  }
  
  // 根據不同操作驗證數據
  switch (data.action) {
    case 'addExpense':
      return validateExpenseData(data.data);
    
    case 'initSheet':
      return true; // 初始化不需要額外驗證
    
    default:
      return false;
  }
}

/**
 * 驗證支出數據
 */
function validateExpenseData(data) {
  if (!data) return false;
  
  // 檢查必填欄位
  const requiredFields = ['date', 'item', 'category', 'amount', 'paymentMethod'];
  
  for (let field of requiredFields) {
    if (!data[field] || data[field].toString().trim() === '') {
      console.error(`缺少必填欄位: ${field}`);
      return false;
    }
  }
  
  // 驗證金額格式
  const amount = parseFloat(data.amount);
  if (isNaN(amount) || amount < 0) {
    console.error('金額格式不正確');
    return false;
  }
  
  // 驗證日期格式
  const date = new Date(data.date);
  if (isNaN(date.getTime())) {
    console.error('日期格式不正確');
    return false;
  }
  
  return true;
}

/**
 * 處理新增支出請求
 */
function handleAddExpense(requestData) {
  try {
    const expenseData = requestData.data;
    
    // 獲取或創建工作表
    const sheet = getOrCreateSheet();
    
    // 確保表頭存在
    ensureHeaders(sheet);
    
    // 準備要插入的數據行
    const rowData = [
      expenseData.date,
      expenseData.item,
      expenseData.category,
      parseFloat(expenseData.amount),
      expenseData.paymentMethod,
      expenseData.note || '',
      new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    ];
    
    // 插入數據到工作表
    sheet.appendRow(rowData);
    
    // 格式化新插入的行
    formatNewRow(sheet, sheet.getLastRow());
    
    console.log('成功新增支出記錄:', rowData);
    
    return createResponse({
      success: true,
      message: '支出記錄已成功保存',
      data: {
        row: sheet.getLastRow(),
        timestamp: new Date().toISOString()
      }
    });
    
  } catch (error) {
    console.error('新增支出記錄失敗:', error);
    return createResponse({
      success: false,
      error: '新增支出記錄失敗: ' + error.toString()
    }, 500);
  }
}

/**
 * 處理初始化工作表請求
 */
function handleInitSheet(requestData) {
  try {
    const sheet = getOrCreateSheet();
    ensureHeaders(sheet);
    
    return createResponse({
      success: true,
      message: '工作表初始化完成',
      data: {
        sheetName: sheet.getName(),
        headers: CONFIG.HEADERS
      }
    });
    
  } catch (error) {
    console.error('初始化工作表失敗:', error);
    return createResponse({
      success: false,
      error: '初始化工作表失敗: ' + error.toString()
    }, 500);
  }
}

/**
 * 獲取或創建工作表
 */
function getOrCreateSheet() {
  try {
    // 嘗試打開指定的Google Sheets
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    
    // 嘗試獲取指定的工作表
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    
    // 如果工作表不存在，創建新的
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
      console.log(`創建新工作表: ${CONFIG.SHEET_NAME}`);
    }
    
    return sheet;
    
  } catch (error) {
    console.error('獲取工作表失敗:', error);
    throw new Error('無法訪問Google Sheets，請檢查工作表ID和權限設定');
  }
}

/**
 * 確保表頭存在
 */
function ensureHeaders(sheet) {
  try {
    // 獲取工作表的第一行數據
    const lastColumn = sheet.getLastColumn();
    const firstRow = (lastColumn > 0) ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
    
    // 如果第一行為空或與預期表頭不符，設置表頭
    if (firstRow.length === 0 || !arraysEqual(firstRow.slice(0, CONFIG.HEADERS.length), CONFIG.HEADERS)) {
      // 清空第一行（如果存在）
      if (lastColumn > 0) {
        sheet.getRange(1, 1, 1, lastColumn).clearContent();
      }
      
      // 設置表頭
      sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
      
      // 格式化表頭
      formatHeaders(sheet);
      
      console.log('表頭已設置:', CONFIG.HEADERS);
    }
    
  } catch (error) {
    console.error('設置表頭失敗:', error);
    throw error;
  }
}

/**
 * 格式化表頭
 */
function formatHeaders(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length);
  
  headerRange
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
}

/**
 * 格式化新插入的數據行
 */
function formatNewRow(sheet, rowNumber) {
  // 確保至少有 CONFIG.HEADERS.length 列
  const numColumns = Math.max(sheet.getLastColumn(), CONFIG.HEADERS.length);
  const dataRange = sheet.getRange(rowNumber, 1, 1, numColumns);
  
  // 設置邊框
  dataRange.setBorder(true, true, true, true, true, true);
  
  // 格式化金額欄位（第4欄）
  const amountCell = sheet.getRange(rowNumber, 4);
  amountCell.setNumberFormat('#,##0.00');
  
  // 格式化日期欄位（第1欄）
  const dateCell = sheet.getRange(rowNumber, 1);
  dateCell.setNumberFormat('yyyy-mm-dd');
  
  // 設置行高
  sheet.setRowHeight(rowNumber, 25);
}

/**
 * 比較兩個陣列是否相等
 */
function arraysEqual(arr1, arr2) {
  if (arr1.length !== arr2.length) return false;
  
  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false;
  }
  
  return true;
}

/**
 * 創建HTTP響應
 * Google Apps Script Web App 在部署時，如果存取權限設置為「任何人」，
 * 會自動處理 CORS 相關的響應頭，無需手動設置。
 */
function createResponse(data, statusCode = 200) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  
  // Apps Script Web App 會自動處理 CORS 響應頭，無需手動添加。
  // 如果需要手動添加，則需要使用 HtmlOutput，並在 HTML 模板中設置。
  
  return output;
}

/**
 * 處理OPTIONS請求（CORS預檢請求）
 * Google Apps Script Web App 會自動處理 OPTIONS 請求，無需手動實現。
 */
function doOptions(e) {
  return createResponse({
    success: true,
    message: 'CORS preflight handled automatically by Apps Script'
  });
}

/**
 * 測試函數 - 用於在Apps Script編輯器中測試
 */
function testAddExpense() {
  const testData = {
    action: 'addExpense',
    data: {
      date: '2025-08-28',
      item: '測試午餐',
      category: '飲食',
      amount: '150.50',
      paymentMethod: '信用卡',
      note: '這是一個測試記錄'
    }
  };
  
  const mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(mockEvent);
  console.log('測試結果:', result.getContent());
}

/**
 * 測試函數 - 初始化工作表
 */
function testInitSheet() {
  const testData = {
    action: 'initSheet'
  };
  
  const mockEvent = {
    postData: {
      contents: JSON.stringify(testData)
    }
  };
  
  const result = doPost(mockEvent);
  console.log('初始化結果:', result.getContent());
}

/**
 * 獲取工作表統計信息
 */
function getSheetStats() {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    // 計算總支出（假設金額在第4欄）
    let totalExpense = 0;
    if (lastRow > 1) { // 排除表頭
      const amountRange = sheet.getRange(2, 4, lastRow - 1, 1);
      const amounts = amountRange.getValues().flat();
      totalExpense = amounts.reduce((sum, amount) => sum + (parseFloat(amount) || 0), 0);
    }
    
    return {
      totalRows: lastRow,
      totalColumns: lastColumn,
      totalRecords: Math.max(0, lastRow - 1), // 排除表頭
      totalExpense: totalExpense,
      lastUpdate: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('獲取統計信息失敗:', error);
    return null;
  }
}

