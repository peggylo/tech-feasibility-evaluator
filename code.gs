/**
 * 審查自動化系統
 * 使用 OpenAI API (GPT-4.1-mini) 進行技術可行性評估
 */

// ==================== 設定區 ====================

const CONFIG = {
  MODEL: "gpt-4.1-mini",           // OpenAI 模型
  MAX_RETRIES: 3,                  // API 失敗重試次數
  DELAY_MS: 500,                   // 批次處理間隔（毫秒）
  API_TIMEOUT: 30000,              // API timeout（30秒）
  
  // Sheet 名稱
  MAIN_SHEET: "原始申請資料",       // 請改成您的主要資料表名稱
  CLAUDE_FRAMEWORK_SHEET: "技術可行性評估框架_Claude",
  GPT_FRAMEWORK_SHEET: "技術可行性評估框架_GPT",
  
  // 必要欄位名稱（標題列）
  REQUIRED_COLUMNS: {
    方案別: "方案別",
    設備數量: "設備數",
    教學用途: "教學用途",
    教師背景: "教師背景",
    Claude評估: "Claude評估",
    Claude說明: "Claude說明",
    chatGPT評估: "chatGPT評估",
    chatGPT說明: "chatGPT說明"
  }
};

// ==================== 建立選單 ====================

/**
 * 當試算表開啟時自動執行
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 自動審查')
    .addItem('評估選取的列', 'evaluateSelectedRows')
    .addSeparator()
    .addItem('批次評估空白（Claude）', 'batchEvaluateClaude')
    .addItem('批次評估空白（GPT）', 'batchEvaluateGPT')
    .addItem('批次評估空白（兩者）', 'batchEvaluateBoth')
    .addToUi();
}

// ==================== 主要功能 ====================

/**
 * 功能1：評估選取的列
 */
function evaluateSelectedRows() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    ui.alert('錯誤', '請先選取要評估的列', ui.ButtonSet.OK);
    return;
  }
  
  // 取得選取的列號（可能多列）
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const rowNumbers = [];
  
  for (let i = 0; i < numRows; i++) {
    rowNumbers.push(startRow + i);
  }
  
  // 排除標題列
  const dataRows = rowNumbers.filter(row => row > 1);
  
  if (dataRows.length === 0) {
    ui.alert('錯誤', '請選取資料列（非標題列）', ui.ButtonSet.OK);
    return;
  }
  
  // 第一個問題：是否評估 Claude
  const claudeResponse = ui.alert(
    '評估 Claude 框架？',
    `即將評估 ${dataRows.length} 列\n\n是否使用 Claude 評估框架進行評估？\n\n（已有評估內容的會自動跳過）`,
    ui.ButtonSet.YES_NO
  );
  
  if (claudeResponse === ui.Button.CLOSE) {
    return; // 用戶關閉對話框
  }
  
  const evalClaude = (claudeResponse === ui.Button.YES);
  
  // 第二個問題：是否評估 GPT
  const gptResponse = ui.alert(
    '評估 GPT 框架？',
    `是否使用 GPT 評估框架進行評估？`,
    ui.ButtonSet.YES_NO
  );
  
  if (gptResponse === ui.Button.CLOSE) {
    return; // 用戶關閉對話框
  }
  
  const evalGPT = (gptResponse === ui.Button.YES);
  
  // 確認至少選了一個
  if (!evalClaude && !evalGPT) {
    ui.alert('提示', '未選擇任何評估方式，已取消', ui.ButtonSet.OK);
    return;
  }
  
  // 決定模式
  let mode;
  if (evalClaude && evalGPT) {
    mode = 'both';
  } else if (evalClaude) {
    mode = 'claude';
  } else {
    mode = 'gpt';
  }
  
  // 執行評估
  const result = processRows(dataRows, mode);
  
  // 顯示結果
  ui.alert(
    '評估完成！',
    `評估列數：${dataRows.length}\n成功：${result.success}\n跳過：${result.skipped}（已有評估）\n失敗：${result.failed}`,
    ui.ButtonSet.OK
  );
}

/**
 * 功能2：批次評估空白（Claude）
 */
function batchEvaluateClaude() {
  batchEvaluate('claude');
}

/**
 * 功能3：批次評估空白（GPT）
 */
function batchEvaluateGPT() {
  batchEvaluate('gpt');
}

/**
 * 功能4：批次評估空白（兩者）
 */
function batchEvaluateBoth() {
  batchEvaluate('both');
}

/**
 * 批次評估邏輯
 */
function batchEvaluate(mode) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MAIN_SHEET);
  
  if (!sheet) {
    ui.alert('錯誤', `找不到「${CONFIG.MAIN_SHEET}」工作表`, ui.ButtonSet.OK);
    return;
  }
  
  // 找出所有需要評估的列
  const emptyRows = findEmptyRows(sheet, mode);
  
  if (emptyRows.length === 0) {
    ui.alert('提示', '沒有找到需要評估的申請', ui.ButtonSet.OK);
    return;
  }
  
  // 確認
  const modeName = mode === 'claude' ? 'Claude' : mode === 'gpt' ? 'GPT' : '兩者';
  const response = ui.alert(
    '確認批次評估',
    `找到 ${emptyRows.length} 筆待評估的申請\n\n使用：${modeName} 評估框架\n\n確定要開始嗎？`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // 執行評估
  const startTime = new Date();
  ui.alert('批次評估中...', '請稍候，處理完成後會顯示結果', ui.ButtonSet.OK);
  
  const result = processRows(emptyRows, mode);
  
  const duration = Math.round((new Date() - startTime) / 1000);
  
  // 顯示結果
  ui.alert(
    '批次評估完成！',
    `成功：${result.success} 筆\n跳過：${result.skipped} 筆（已有評估）\n失敗：${result.failed} 筆\n\n總耗時：${duration} 秒`,
    ui.ButtonSet.OK
  );
}

// ==================== 核心處理邏輯 ====================

/**
 * 處理多列評估
 */
function processRows(rowNumbers, mode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MAIN_SHEET);
  const columnMap = getColumnMap(sheet);
  
  let success = 0;
  let skipped = 0;
  let failed = 0;
  
  for (let i = 0; i < rowNumbers.length; i++) {
    const row = rowNumbers[i];
    
    // 顯示進度（每5筆更新一次）
    if (i % 5 === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `進度：${i + 1}/${rowNumbers.length}\n目前：第 ${row} 列`,
        '評估中...',
        3
      );
    }
    
    // 讀取該列資料
    const data = readRowData(sheet, row, columnMap);
    
    // 評估 Claude
    if (mode === 'claude' || mode === 'both') {
      const claudeResult = evaluateRow(data, 'claude', sheet, row, columnMap);
      if (claudeResult === 'success') success++;
      else if (claudeResult === 'skipped') skipped++;
      else failed++;
    }
    
    // 評估 GPT
    if (mode === 'gpt' || mode === 'both') {
      const gptResult = evaluateRow(data, 'gpt', sheet, row, columnMap);
      if (gptResult === 'success') success++;
      else if (gptResult === 'skipped') skipped++;
      else failed++;
    }
    
    // 延遲（避免 rate limit）
    if (i < rowNumbers.length - 1) {
      Utilities.sleep(CONFIG.DELAY_MS);
    }
  }
  
  return { success, skipped, failed };
}

/**
 * 評估單一列
 */
function evaluateRow(data, framework, sheet, row, columnMap) {
  // 檢查是否已有評估
  const evalColIndex = framework === 'claude' ? columnMap.Claude評估 : columnMap.chatGPT評估;
  const existingEval = sheet.getRange(row, evalColIndex).getValue();
  
  if (existingEval && existingEval.toString().trim() !== '') {
    return 'skipped';
  }
  
  // 讀取評估框架
  const prompt = buildPrompt(data, framework);
  
  // 呼叫 API
  const apiResult = callOpenAI(prompt);
  
  if (!apiResult.success) {
    // 寫入錯誤訊息
    const descColIndex = framework === 'claude' ? columnMap.Claude說明 : columnMap.chatGPT說明;
    
    sheet.getRange(row, evalColIndex).setValue('API錯誤');
    sheet.getRange(row, descColIndex).setValue(apiResult.error || '無法連接API，請稍後重試');
    
    return 'failed';
  }
  
  // 解析結果
  const parsed = parseAPIResponse(apiResult.data);
  
  if (!parsed.success) {
    const descColIndex = framework === 'claude' ? columnMap.Claude說明 : columnMap.chatGPT說明;
    
    sheet.getRange(row, evalColIndex).setValue('格式錯誤');
    sheet.getRange(row, descColIndex).setValue('AI回應格式不正確');
    
    return 'failed';
  }
  
  // 寫入結果
  const descColIndex = framework === 'claude' ? columnMap.Claude說明 : columnMap.chatGPT說明;
  
  sheet.getRange(row, evalColIndex).setValue(parsed.evaluation);
  sheet.getRange(row, descColIndex).setValue(parsed.description);
  
  return 'success';
}

/**
 * 組裝 Prompt
 */
function buildPrompt(data, framework) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = framework === 'claude' ? 
    CONFIG.CLAUDE_FRAMEWORK_SHEET : CONFIG.GPT_FRAMEWORK_SHEET;
  
  const frameworkSheet = ss.getSheetByName(sheetName);
  
  if (!frameworkSheet) {
    throw new Error(`找不到評估框架 Sheet：${sheetName}`);
  }
  
  // 讀取框架內容（資料是垂直排列：A欄=欄位名稱，B欄=內容）
  // A2=系統角色, B2=內容
  // A3=設備規格說明, B3=內容
  // A4=評估框架, B4=內容
  // A5=評估範例, B5=內容
  // A6=輸出格式要求, B6=內容
  
  const systemRole = frameworkSheet.getRange("B2").getValue();
  const deviceSpec = frameworkSheet.getRange("B3").getValue();
  const frameworkContent = frameworkSheet.getRange("B4").getValue();
  const examples = frameworkSheet.getRange("B5").getValue();
  const outputFormat = frameworkSheet.getRange("B6").getValue();
  
  // 組合完整 prompt
  const fullPrompt = `${systemRole}

${deviceSpec}

${frameworkContent}

${examples}

${outputFormat}

【待評估申請案】
方案別：${data.方案別 || "未提供"}
設備數量：${data.設備數量 || "未提供"}
教學用途：${data.教學用途 || "未提供"}
教師背景：${data.教師背景 || "未提供"}

請依據上述框架進行評估，直接輸出JSON格式，不要有其他文字。`;

  return fullPrompt;
}

/**
 * 呼叫 OpenAI API
 */
function callOpenAI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  
  if (!apiKey || apiKey.trim() === '') {
    return {
      success: false,
      error: '未設定 API Key。請在「專案設定」→「指令碼屬性」中設定 OPENAI_API_KEY'
    };
  }
  
  const url = 'https://api.openai.com/v1/chat/completions';
  
  const payload = {
    model: CONFIG.MODEL,
    messages: [
      {
        role: "user",
        content: prompt
      }
    ],
    temperature: 0.3,  // 降低隨機性，提高一致性
    max_tokens: 300    // 限制輸出長度
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  // 重試邏輯
  for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const statusCode = response.getResponseCode();
      
      if (statusCode === 200) {
        const json = JSON.parse(response.getContentText());
        const content = json.choices[0].message.content;
        
        return {
          success: true,
          data: content
        };
      } else {
        Logger.log(`API 錯誤 (嘗試 ${attempt}/${CONFIG.MAX_RETRIES}): ${statusCode}`);
        Logger.log('回應內容: ' + response.getContentText());
        
        if (attempt < CONFIG.MAX_RETRIES) {
          Utilities.sleep(1000 * attempt); // 遞增延遲
        }
      }
    } catch (error) {
      Logger.log(`API 異常 (嘗試 ${attempt}/${CONFIG.MAX_RETRIES}): ${error.toString()}`);
      
      if (attempt < CONFIG.MAX_RETRIES) {
        Utilities.sleep(1000 * attempt);
      }
    }
  }
  
  return {
    success: false,
    error: `API 呼叫失敗（已重試 ${CONFIG.MAX_RETRIES} 次）`
  };
}

/**
 * 解析 API 回應
 */
function parseAPIResponse(responseText) {
  try {
    // 清理可能的 markdown 格式
    let cleaned = responseText.trim();
    
    // 移除可能的 ```json 標記
    if (cleaned.startsWith('```json')) {
      cleaned = cleaned.substring(7);
    }
    if (cleaned.startsWith('```')) {
      cleaned = cleaned.substring(3);
    }
    if (cleaned.endsWith('```')) {
      cleaned = cleaned.substring(0, cleaned.length - 3);
    }
    
    cleaned = cleaned.trim();
    
    // 解析 JSON
    const json = JSON.parse(cleaned);
    
    if (!json.技術可行性 || !json.說明) {
      return {
        success: false,
        error: 'JSON 格式不完整'
      };
    }
    
    return {
      success: true,
      evaluation: json.技術可行性,
      description: json.說明
    };
    
  } catch (error) {
    Logger.log('JSON 解析失敗: ' + error.toString());
    Logger.log('原始回應: ' + responseText);
    
    return {
      success: false,
      error: 'JSON 解析失敗'
    };
  }
}

// ==================== 輔助函數 ====================

/**
 * 取得欄位對應表
 */
function getColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  
  for (const key in CONFIG.REQUIRED_COLUMNS) {
    const colName = CONFIG.REQUIRED_COLUMNS[key];
    const colIndex = headers.indexOf(colName);
    
    if (colIndex === -1) {
      throw new Error(`找不到必要欄位：${colName}`);
    }
    
    map[key] = colIndex + 1; // 轉換為 1-based index
  }
  
  return map;
}

/**
 * 讀取列資料
 */
function readRowData(sheet, row, columnMap) {
  return {
    方案別: sheet.getRange(row, columnMap.方案別).getValue(),
    設備數量: sheet.getRange(row, columnMap.設備數量).getValue(),
    教學用途: sheet.getRange(row, columnMap.教學用途).getValue(),
    教師背景: sheet.getRange(row, columnMap.教師背景).getValue()
  };
}

/**
 * 找出需要評估的空白列
 */
function findEmptyRows(sheet, mode) {
  const columnMap = getColumnMap(sheet);
  const lastRow = sheet.getLastRow();
  const emptyRows = [];
  
  for (let row = 2; row <= lastRow; row++) {
    let needEval = false;
    
    if (mode === 'claude' || mode === 'both') {
      const claudeEval = sheet.getRange(row, columnMap.Claude評估).getValue();
      if (!claudeEval || claudeEval.toString().trim() === '') {
        needEval = true;
      }
    }
    
    if (mode === 'gpt' || mode === 'both') {
      const gptEval = sheet.getRange(row, columnMap.chatGPT評估).getValue();
      if (!gptEval || gptEval.toString().trim() === '') {
        needEval = true;
      }
    }
    
    if (needEval) {
      // 確認該列有教學用途（不是空列）
      const 教學用途 = sheet.getRange(row, columnMap.教學用途).getValue();
      if (教學用途 && 教學用途.toString().trim() !== '') {
        emptyRows.push(row);
      }
    }
  }
  
  return emptyRows;
}

// ==================== 測試函數（可選） ====================

/**
 * 測試 API 連接
 */
function testAPIConnection() {
  const testPrompt = "請回應一個JSON格式：{\"技術可行性\": \"測試\", \"說明\": \"這是測試\"}";
  const result = callOpenAI(testPrompt);
  
  if (result.success) {
    Logger.log('✅ API 連接成功！');
    Logger.log('回應內容：' + result.data);
  } else {
    Logger.log('❌ API 連接失敗：' + result.error);
  }
}
