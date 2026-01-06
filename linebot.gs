/**
 * 執行 initialize() 並把結果寫到 Logger，方便在 GAS 編輯器中一鍵測試
 */
function runInitializeAndReport() {
  try {
    const res = initialize();
    Logger.log("initialize() 返回: " + res);
  } catch (e) {
    Logger.log("initialize() 執行失敗: " + e + (e.stack ? "\n" + e.stack : ""));
  }

  try {
    ensureSsAvailable_();
    const logSheet = ss.getSheetByName("LOG");
    if (!logSheet) {
      Logger.log("找不到 LOG 工作表");
      return;
    }
    const lastRow = logSheet.getLastRow();
    if (lastRow < 1) {
      Logger.log("LOG 表沒有資料");
      return;
    }
    const startRow = Math.max(1, lastRow - 20 + 1);
    const rowCount = Math.max(1, Math.min(20, lastRow - startRow + 1));
    const colCount = Math.min(2, logSheet.getLastColumn());
    const data = logSheet.getRange(startRow, 1, rowCount, colCount).getValues();
    Logger.log("LOG 表最近紀錄: " + JSON.stringify(data));
  } catch (e) {
    Logger.log("讀取 LOG 表失敗: " + e + (e.stack ? "\n" + e.stack : ""));
  }
}
/**
 * LINE Bot Assistant
 * Version: 4.9.0 (Persona Transitions)
 * Last Updated: 2026-01-06
 * Key changes:
 * - [Fix] 暴力排版：中文標點後強制插入空行 (不依賴 AI 換行)
 * - [Deploy] 使用 Clasp 自動部署至 Google Apps Script
 * - [Fix] 排版優化：強制句子間空一行 (Code-based Formatting)
 * - [New] 實作 Batch Logging 機制：日誌寫入延後至回覆後一次性處理 (Speed Up!)
 * - Provider 讀 Prompt!A1（XAI / OPENROUTER）
 * - Model 讀 Prompt!A2（依供應商）
 * - /clear 會清除 provider/model/prompt/history 快取並讓新設定立即生效
 * - initialize() 幫你補 A1/A2 預設與註解
 * - 其他功能維持（Hybrid Cache、History、Batch Queue、Loading、Quota、Retry-Key 等）
 */

// =========================================================================
// Constants and Configuration
// =========================================================================

const SHEET_NAMES = {
  RECORDS: "所有紀錄",
  LOG: "LOG",
  PROMPT: "Prompt",
  LAST_CONVERSATION: "上次對話",
  INDIVIDUAL_MODE: "個別模式",
};

const TIMEOUT = {
  API_FETCH: 20000, // 20 seconds
  LINE_API: 10000, // 10 seconds
};

const RETRY = {
  MAX_ATTEMPTS: 2,
  DELAY: 1000,
};

const MAX_OUTPUT_TOKENS = 500;
const HISTORY_PAIR_LIMIT = 10; // 只保留最近 10 對 (user+assistant)
const HISTORY_LENGTH_LIMIT = HISTORY_PAIR_LIMIT * 2;
const CACHE_TTL_SEC = 3600; // 1 hour for history cache
const PROMPT_CACHE_EXPIRATION = 1800; // 30 min

const CACHE_KEYS = {
  GLOBAL_BASE_PROMPT: "globalBasePrompt_C1",
  SPECIFIC_PROMPT_PREFIX: "specificPrompt_B1_",
  HISTORY_PREFIX: "hist:", // + ns:promptNum:contextId
  PROVIDER: "provider_A1",
  MODEL: "model_A2",
};

const HIST_NS_PROP_KEY = "HIST_NS_V1";
const LINE_TEXT_MAX = 4000;

// ===== Push 開關與限額守門 =====
const ALLOW_PUSH =
  (PropertiesService.getScriptProperties().getProperty("ALLOW_PUSH") ||
    "false") === "true";

var LOG_BUFFER = []; // 用來暫存日誌

// Active Spreadsheet handle (used by many helper functions). If the script
// is bound to a Sheet this will work. For standalone scripts, you can set
// a SPREADSHEET_ID in Script Properties and it will try to open by ID.
let ss = null;
try {
  ss = SpreadsheetApp.getActiveSpreadsheet();
} catch (e) {
  ss = null;
}
if (!ss) {
  const fallbackId =
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (fallbackId) {
    try {
      ss = SpreadsheetApp.openById(fallbackId);
    } catch (e) {
      ss = null;
    }
  }
}

function ensureSsAvailable_() {
  if (!ss)
    throw new Error(
      "Active spreadsheet not found. Open this script from the target Google Sheet or set SPREADSHEET_ID in Script Properties."
    );
}
// Note: initialization (sheet creation / maintenance) is handled by initialize()
// and should not run at module load time to avoid repeated execution when
// time-based triggers invoke functions. Call initialize() manually or via the
// provided runInitializeAndReport() helper.

// =========================================================================
// Core Message Handling Logic
// =========================================================================

function handleMessage(userMessage, userId, replyToken, contextId) {
  try {
    if (
      !userMessage ||
      typeof userMessage !== "string" ||
      userMessage.trim() === ""
    ) {
      writeLog(`空訊息，略過: contextId=${contextId}, userId=${userId}`);
      return;
    }

    const trimmedMessage = userMessage.trim();
    writeLog(
      `處理訊息: '${trimmedMessage.substring(
        0,
        50
      )}...' in context ${contextId}`
    );

    if (isCommand(trimmedMessage)) {
      const response = handleCommand(trimmedMessage, userId, contextId);
      replyMessage(replyToken, response);
      if (trimmedMessage.toLowerCase() !== "/reset") {
        queueRecord({
          userId: userId,
          text: trimmedMessage,
          groupId: contextId,
          role: "user",
          resetFlag: "",
        });
      }
      return;
    }

    // 1:1 加載動畫
    if (contextId === userId) {
      showLoadingAnimation(userId, 15);
    }

    const isLongOrComplex =
      trimmedMessage.length > 300 ||
      /分析|總結|產生圖|抓網址|翻譯/i.test(trimmedMessage);
    let usedReply = false;
    if (isLongOrComplex) {
      replyMessage(replyToken, "處理中，請稍候...");
      usedReply = true;
    }

    const basePrompt = getGlobalBasePrompt();
    const specificPrompt = getFullPrompt(contextId); // ← 傳入 contextId
    const combinedPrompt = `${basePrompt}\n\n${specificPrompt}`.trim();

    const currentHistory = getHistoryFromCacheOrSheet(contextId);
    writeLog(
      `獲取 context ${contextId} 的歷史: ${currentHistory.length} 條 (上限 ${HISTORY_LENGTH_LIMIT})`
    );

    const userMsgObj = { role: "user", content: trimmedMessage };
    const messages = [
      { role: "system", content: combinedPrompt },
      ...currentHistory,
      userMsgObj,
    ];

    writeLog(
      `呼叫 AI API，${messages.length} 條訊息 (含 system) for context ${contextId}`
    );
    const start = Date.now();
    const assistantResponseText = callChatGPTWithRetry(messages);
    const took = Date.now() - start;

    if (assistantResponseText && assistantResponseText.trim() !== "") {
      const finalRawText = assistantResponseText.trim();
      const finalText = formatForLineMobile(finalRawText);

      if (contextId === userId) {
        if ((usedReply || took > 45000) && canUsePush(contextId, userId)) {
          pushMessage(userId, finalText);
        } else if (!usedReply) {
          replyMessage(replyToken, finalText);
          usedReply = true;
        } else {
          writeLog("已回『處理中』但 push 不允許或超額，省額度不補發。");
        }
      } else {
        if (!usedReply) replyMessage(replyToken, finalText);
      }

      queueRecord({
        userId: userId,
        text: trimmedMessage,
        groupId: contextId,
        role: "user",
        resetFlag: "",
      });
      queueRecord({
        userId: userId,
        text: finalText,
        groupId: contextId,
        role: "assistant",
        resetFlag: "",
      });
      const assistantMsgObj = { role: "assistant", content: finalText };
      updateHistorySheetAndCache(
        contextId,
        currentHistory,
        userMsgObj,
        assistantMsgObj
      );
    } else {
      writeLog(`AI API 調用失敗或回應為空 for context ${contextId}`);
      const errorMsg = "抱歉，暫時無法處理你的請求，稍後再試。";
      if (!usedReply) {
        replyMessage(replyToken, errorMsg);
      } else if (contextId === userId && canUsePush(contextId, userId)) {
        pushMessage(userId, errorMsg);
      }
      queueRecord({
        userId: userId,
        text: trimmedMessage,
        groupId: contextId,
        role: "user",
        resetFlag: "",
      });
      queueRecord({
        userId: userId,
        text: "[AI FAILED]",
        groupId: contextId,
        role: "assistant",
        resetFlag: "",
      });
    }
  } catch (error) {
    writeLog(
      "處理訊息錯誤 (handleMessage): " +
        error +
        (error.stack ? "\nStack: " + error.stack : "")
    );
    try {
      const errorMsg = "哎呀，處理你的訊息時出了問題，稍後再試。";
      if (replyToken) {
        replyMessage(replyToken, errorMsg);
      } else if (contextId === userId && canUsePush(contextId, userId)) {
        pushMessage(userId, errorMsg);
      }
    } catch (replyError) {
      writeLog("發送錯誤回覆失敗: " + replyError);
    }
  }
}

// =========================================================================
// Hybrid History Handling (Cache + "上次對話" Sheet)
// =========================================================================

function getHistoryFromCacheOrSheet(contextId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = buildHistoryKey_(contextId);
  let cachedHistory = cache.get(cacheKey);
  if (cachedHistory) {
    try {
      const history = JSON.parse(cachedHistory);
      if (Array.isArray(history)) return history;
    } catch (e) {
      writeLog(`解析歷史快取錯誤 for ${contextId}: ${e}`);
    }
  }
  const historyFromSheet = getHistoryFromSheet(contextId);
  const jsonStr = JSON.stringify(historyFromSheet);
  safeJsonPutToCache_(cache, cacheKey, jsonStr, CACHE_TTL_SEC);
  writeLog(
    `從工作表讀取並快取歷史 for ${contextId}: ${historyFromSheet.length} 條`
  );
  return historyFromSheet;
}

function getHistoryFromSheet(contextId) {
  const functionName = "getHistoryFromSheet";
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.LAST_CONVERSATION
    );
    if (!sheet) {
      writeLog(
        `${functionName}: 工作表 ${SHEET_NAMES.LAST_CONVERSATION} 不存在，嘗試初始化`
      );
      initialize();
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        SHEET_NAMES.LAST_CONVERSATION
      );
      if (!sheet)
        throw new Error(
          `工作表 ${SHEET_NAMES.LAST_CONVERSATION} 不存在且無法自動創建`
        );
      return [];
    }

    const textFinder = sheet
      .getRange("A:A")
      .createTextFinder(contextId)
      .matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (foundCell) {
      const row = foundCell.getRow();
      const historyJson = sheet.getRange(row, 2).getValue();
      if (
        historyJson &&
        typeof historyJson === "string" &&
        historyJson.trim() !== ""
      ) {
        try {
          const history = JSON.parse(historyJson);
          return Array.isArray(history) ? history : [];
        } catch (parseError) {
          writeLog(
            `${functionName}: 解析 context ${contextId} (行 ${row}) JSON 失敗: ${parseError}.`
          );
          return [];
        }
      } else {
        return [];
      }
    } else {
      return [];
    }
  } catch (error) {
    writeLog(`${functionName}: 讀取歷史錯誤 for ${contextId}: ${error}`);
    return [];
  }
}

function updateHistorySheetAndCache(
  contextId,
  previousHistory,
  userMessage,
  assistantMessage
) {
  const functionName = "updateHistorySheetAndCache";
  return withLock_(() => {
    try {
      let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        SHEET_NAMES.LAST_CONVERSATION
      );
      if (!sheet) {
        writeLog(
          `${functionName}: 工作表 ${SHEET_NAMES.LAST_CONVERSATION} 不存在，嘗試初始化`
        );
        initialize();
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
          SHEET_NAMES.LAST_CONVERSATION
        );
        if (!sheet)
          throw new Error(
            `工作表 ${SHEET_NAMES.LAST_CONVERSATION} 不存在且無法自動創建`
          );
      }
      let base = Array.isArray(previousHistory) ? previousHistory.slice() : [];
      if (base.length % 2 !== 0) base.shift();
      let newHistory = [...base, userMessage, assistantMessage];
      while (newHistory.length > HISTORY_LENGTH_LIMIT) {
        newHistory.shift();
        newHistory.shift();
      }
      const newHistoryJson = JSON.stringify(newHistory);
      const textFinder = sheet
        .getRange("A:A")
        .createTextFinder(contextId)
        .matchEntireCell(true);
      const foundCell = textFinder.findNext();
      if (foundCell) {
        sheet.getRange(foundCell.getRow(), 2).setValue(newHistoryJson);
      } else {
        sheet.appendRow([contextId, newHistoryJson]);
        writeLog(`${functionName}: 為 context ${contextId} 新增了歷史行`);
      }
      const cacheKey = buildHistoryKey_(contextId);
      safeJsonPutToCache_(
        CacheService.getScriptCache(),
        cacheKey,
        newHistoryJson,
        CACHE_TTL_SEC
      );
    } catch (error) {
      writeLog(`${functionName}: 更新歷史錯誤 for ${contextId}: ${error}`);
    }
  });
}

function clearHistorySheetAndCache(contextId) {
  const functionName = "clearHistorySheetAndCache";
  return withLock_(() => {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        SHEET_NAMES.LAST_CONVERSATION
      );
      if (!sheet) {
        writeLog(
          `${functionName}: 工作表 ${SHEET_NAMES.LAST_CONVERSATION} 不存在，無法清除 for ${contextId}`
        );
        return;
      }
      const textFinder = sheet
        .getRange("A:A")
        .createTextFinder(contextId)
        .matchEntireCell(true);
      const foundCell = textFinder.findNext();
      if (foundCell) {
        const row = foundCell.getRow();
        sheet.getRange(row, 2).clearContent();
        writeLog(
          `${functionName}: 清除了 context ${contextId} (行 ${row}) 的 Sheet 歷史`
        );
      } else {
        writeLog(
          `${functionName}: 未找到 context ${contextId}，無需清除 Sheet 歷史`
        );
      }
      const cache = CacheService.getScriptCache();
      const cacheKey = buildHistoryKey_(contextId);
      cache.remove(cacheKey);
      writeLog(`${functionName}: 清除了 context ${contextId} 的歷史快取`);
    } catch (error) {
      writeLog(`${functionName}: 清除歷史錯誤 for ${contextId}: ${error}`);
    }
  });
}

// =========================================================================
// Provider / Model from Prompt!A1 / A2
// =========================================================================

function getProviderFromSheet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEYS.PROVIDER);
  if (hit) return hit;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET_NAMES.PROMPT
  );
  if (!sh) throw new Error("找不到工作表 Prompt");

  let provider = String(sh.getRange("A1").getValue() || "")
    .trim()
    .toUpperCase();
  if (provider !== "XAI" && provider !== "OPENROUTER") {
    provider = "XAI";
    try {
      sh.getRange("A1").setValue(provider);
    } catch (_) {}
  }
  cache.put(CACHE_KEYS.PROVIDER, provider, PROMPT_CACHE_EXPIRATION);
  return provider;
}

function getModelNameFromSheet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEYS.MODEL);
  if (hit) return hit;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET_NAMES.PROMPT
  );
  if (!sh) throw new Error("找不到工作表 Prompt");

  let model = String(sh.getRange("A2").getValue() || "").trim();
  const provider = getProviderFromSheet();
  if (!model) {
    model =
      provider === "OPENROUTER" ? "openai/gpt-4o-mini" : "x-ai/grok-3-beta";
    try {
      sh.getRange("A2").setValue(model);
    } catch (_) {}
  }
  cache.put(CACHE_KEYS.MODEL, model, PROMPT_CACHE_EXPIRATION);
  return model;
}

// =========================================================================
// AI API Call Handling (OpenRouter / xAI)
// =========================================================================

function callChatGPTWithRetry(messages) {
  let attempts = 0;
  let lastError = null;
  while (attempts < RETRY.MAX_ATTEMPTS) {
    attempts++;
    try {
      const response = callChatApi(messages);
      if (response && response.trim() !== "") {
        writeLog(`AI API 成功 (嘗試 ${attempts})`);
        return response;
      } else {
        lastError = new Error("API 回應無效或為空");
        writeLog(`API 回應空 (嘗試 ${attempts})`);
      }
    } catch (error) {
      lastError = error;
      writeLog(`AI API 失敗 (嘗試 ${attempts}): ${error}`);
      if (attempts < RETRY.MAX_ATTEMPTS) Utilities.sleep(RETRY.DELAY);
    }
  }
  writeLog(`AI API 失敗，達最大重試次數: ${lastError}`);
  return null;
}

function callChatApi(messages) {
  let provider = "XAI";
  let apiKey = null;
  let url = "";
  let specificHeaders = {};

  try {
    provider = getProviderFromSheet(); // ← A1
  } catch (propError) {
    writeLog(`讀取 Provider(A1) 錯誤，使用預設 XAI: ${propError}`);
    provider = "XAI";
  }

  if (provider === "XAI") {
    apiKey = getXaiApiKey();
    url = "https://api.x.ai/v1/chat/completions";
    if (!apiKey) throw new Error("xAI API key is missing.");
    specificHeaders = {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json",
    };
  } else {
    apiKey = getOpenRouterKey();
    url = "https://openrouter.ai/api/v1/chat/completions";
    if (!apiKey) throw new Error("OpenRouter API key is missing.");
    let siteUrl =
      PropertiesService.getScriptProperties().getProperty("YOUR_SITE_URL") ||
      "<YOUR_SITE_URL_DEFAULT>";
    let appName =
      PropertiesService.getScriptProperties().getProperty("YOUR_SITE_NAME") ||
      "<YOUR_APP_NAME_DEFAULT>";
    specificHeaders = {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json",
      "HTTP-Referer": siteUrl,
      "X-Title": appName,
    };
  }

  const modelName = getModelNameFromSheet(); // ← A2
  const payload = {
    model: modelName,
    messages: messages.map((m) => ({ role: m.role, content: m.content })),
    max_tokens: MAX_OUTPUT_TOKENS,
  };

  const options = {
    method: "post",
    headers: specificHeaders,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  var _t0 = Date.now();
  writeLog(`向 ${provider} (${modelName}) 發送 API 請求...`);
  let response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (fetchError) {
    var _t1 = Date.now();
    writeLog(provider + " APIFetch took " + (_t1 - _t0) + " ms");
    writeLog(`UrlFetchApp.fetch (${provider}) 失敗: ${fetchError}`);
    throw new Error(`Network error (${provider}): ${fetchError.message}`);
  }
  var _t1 = Date.now();
  writeLog(provider + " APIFetch took " + (_t1 - _t0) + " ms");

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  writeLog(`${provider} API 狀態碼: ${responseCode}`);

  if (responseCode === 200) {
    try {
      const json = JSON.parse(responseText);
      if (json.error) {
        writeLog(`${provider} API 錯誤: ${JSON.stringify(json.error)}`);
        throw new Error(
          json.error.message || `Unknown API error from ${provider}`
        );
      }
      if (
        json.choices &&
        json.choices[0] &&
        json.choices[0].message &&
        typeof json.choices[0].message.content === "string"
      ) {
        const result = json.choices[0].message.content.trim();
        if (result) {
          writeLog(`${provider} API 回應長度: ${result.length}`);
          return result;
        } else {
          writeLog(`${provider} API 回應空`);
          throw new Error("Empty content");
        }
      } else {
        writeLog(
          `${provider} API 回應格式錯誤: ${responseText.substring(0, 200)}...`
        );
        throw new Error(`Invalid response format from ${provider}`);
      }
    } catch (parseError) {
      writeLog(`解析 ${provider} API 回應錯誤: ${parseError}`);
      throw new Error(`Parse error (${provider}): ${parseError.message}`);
    }
  } else {
    writeLog(
      `${provider} API 失敗，狀態碼: ${responseCode}, 內容: ${responseText.substring(
        0,
        200
      )}...`
    );
    let errorMsg = `${provider} API Error ${responseCode}`;
    if (responseCode === 429)
      errorMsg += ": Rate limit exceeded or spending limit issue.";
    else if (responseCode === 401)
      errorMsg += ": Unauthorized (Check API Key).";
    else if (responseCode === 400)
      errorMsg += ": Bad Request (Check payload/model).";
    errorMsg += ` ${responseText.substring(0, 100)}...`;
    throw new Error(errorMsg);
  }
}

// =========================================================================
// Commands (/help, /reset, /clear, /p)
// =========================================================================

function isCommand(text) {
  return typeof text === "string" && text.trim().startsWith("/");
}

function handleCommand(command, userId, contextId) {
  let response = "";
  try {
    const commandClean = command.trim().toLowerCase();
    writeLog(
      `處理指令 '${commandClean}' from user ${userId} in context ${contextId}`
    );

    if (commandClean === "/help") {
      response = getHelpText();
    } else if (commandClean === "/reset") {
      writeLog(`用戶 ${userId} 在 context ${contextId} 執行 /reset`);
      clearHistorySheetAndCache(contextId);
      queueRecord({
        userId: userId,
        text: command.trim(),
        groupId: contextId,
        role: "user",
        resetFlag: "TRUE",
      });
      response = "對話歷史已重置。下次訊息將從新開始。\n(永久紀錄不受影響)";
    } else if (commandClean === "/clear") {
      const clearedKeys = clearPromptCache(); // 會連 provider/model 一起清
      bumpHistNs_();
      writeLog(
        `用戶 ${userId} 清除了 Prompt/Provider/Model 與歷史快取命名空間 in context ${contextId}. Cleared: ${
          clearedKeys.join(", ") || "None"
        }`
      );
      response = `Prompt/Provider/Model 快取已清除。下次讀取將從工作表重新載入。`;
    } else if (commandClean === "/p") {
      response = getPromptList();
    } else if (
      commandClean.startsWith("/p") &&
      (commandClean.length > 2 || commandClean.includes(" "))
    ) {
      response = handlePromptChange(commandClean, userId, contextId);

      // 修正判斷字串與 handlePromptChange 回傳內容一致
      if (response.startsWith("✅")) {
        // 從回應文字中提取新的人格名稱
        let promptName = "新的人格模式";
        const nameMatch = response.match(/模式：(.+?) \(/);
        if (nameMatch) promptName = nameMatch[1];

        const match = response.match(/編號 (\d+)/);
        if (match && match[1]) {
          const newPromptNumber = parseInt(match[1], 10);
          clearSpecificPromptCache(newPromptNumber);
        }
        bumpHistNs_();

        // 注入「導演紙條」轉場指令，解決上下文慣性
        injectSystemMarker(contextId, promptName);
      }
    } else {
      response = `未知指令：'${command.trim()}'。\n輸入 /help 查看可用指令。`;
    }

    if (!response) {
      response = `未知指令：'${command.trim()}'。\n輸入 /help 查看可用指令。`;
    }
    return response;
  } catch (error) {
    writeLog(`處理指令 '${command}' 錯誤: ${error}`);
    return "執行指令時發生錯誤，請檢查日誌。";
  }
}

function getHelpText() {
  return [
    "--- 指令說明 ---",
    "/help : 顯示此說明",
    "/reset : 重置當前對話歷史記憶（不影響個別模式設定）",
    "/p : 列出所有可用提示詞",
    "/p [編號] : 切換到指定編號的個別模式（保留對話記錄）",
    "/clear : 清除所有快取（不刪除個別模式記錄）",
    "-------------------",
    "💡 個別模式會記住每個對話（群組/個人）的專屬設定",
  ].join("\n");
}

function handlePromptChange(command, userId, contextId) {
  try {
    const promptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.PROMPT
    );
    if (!promptSheet) throw new Error(`找不到工作表 ${SHEET_NAMES.PROMPT}`);

    let promptNumberStr = "";
    if (command.startsWith("/p ")) {
      promptNumberStr = command.substring(3).trim();
    } else if (command.length > 2 && !command.includes(" ")) {
      promptNumberStr = command.substring(2).trim();
    } else {
      return "指令格式錯誤，請使用 /p[編號] 或 /p [編號]。";
    }

    const promptNumber = parseInt(promptNumberStr, 10);
    if (isNaN(promptNumber) || promptNumber <= 0) {
      writeLog(`用戶 ${userId} 輸入無效提示詞編號: '${promptNumberStr}'`);
      return `無效的提示詞編號 '${promptNumberStr}'，請輸入正整數。`;
    }

    const lastRow = promptSheet.getLastRow();
    if (lastRow < 4) {
      writeLog(`用戶 ${userId} 請求提示詞但工作表無資料`);
      return `找不到提示詞資料。\n請在 Prompt 工作表填寫提示詞（從第 4 行開始）。`;
    }

    const promptData = promptSheet.getRange("A4:B" + lastRow).getValues();
    let isValidNumber = false;
    let promptName = `編號 ${promptNumber}`;
    for (const row of promptData) {
      if (row[0] && !isNaN(Number(row[0])) && Number(row[0]) === promptNumber) {
        isValidNumber = true;
        promptName =
          row[1] && String(row[1]).trim() ? String(row[1]).trim() : promptName;
        break;
      }
    }
    if (!isValidNumber) {
      writeLog(`用戶 ${userId} 請求不存在的提示詞編號: ${promptNumber}`);
      return `找不到編號為 ${promptNumber} 的提示詞。\n請使用 /p 查看可用列表。`;
    }

    // ========== 寫入「個別模式」而非 Prompt!B1 ==========
    setIndividualMode(contextId, promptNumber, promptName);

    // ========== 不清除歷史（保留對話記錄） ==========
    // 清除快取以讀取新的 Prompt，但保留歷史
    clearPromptCache();

    writeLog(
      `用戶 ${userId} 在 context ${contextId} 切換個別模式為 #${promptNumber}: ${promptName}`
    );
    return `✅ 已切換至個別模式：${promptName} (編號 ${promptNumber})`;
  } catch (error) {
    writeLog(`切換提示詞錯誤 (指令: ${command}): ${error}`);
    return "切換提示詞時發生錯誤。";
  }
}

function getPromptList() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.PROMPT
    );
    if (!sheet) throw new Error(`找不到工作表 ${SHEET_NAMES.PROMPT}`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      writeLog("提示詞列表為空");
      return "目前沒有可用的提示詞。\n請在 Prompt 工作表的 A 欄填寫編號、B 欄填寫名稱（從第 4 行開始）。";
    }

    const data = sheet.getRange("A4:B" + lastRow).getValues();
    let prompts = [];
    for (const row of data) {
      if (
        row[0] &&
        !isNaN(Number(row[0])) &&
        Number(row[0]) > 0 &&
        row[1] &&
        String(row[1]).trim()
      ) {
        prompts.push(`${Number(row[0])}. ${String(row[1]).trim()}`);
      }
    }
    if (prompts.length === 0) {
      writeLog("提示詞列表為空");
      return "目前沒有可用的提示詞。\n請在 Prompt 工作表的 A 欄填寫編號、B 欄填寫名稱（從第 4 行開始）。";
    }
    writeLog(`獲取 ${prompts.length} 個可用提示詞`);
    return ["可用提示詞（使用 /p [編號] 切換）：", ...prompts].join("\n");
  } catch (error) {
    writeLog("讀取提示詞列表錯誤: " + error);
    return "無法讀取提示詞列表。";
  }
}

// =========================================================================
// Prompt Handling (Base + Specific)
// =========================================================================

function getGlobalBasePrompt() {
  const cache = CacheService.getScriptCache();
  const cachedPrompt = cache.get(CACHE_KEYS.GLOBAL_BASE_PROMPT);
  if (cachedPrompt) return cachedPrompt;

  try {
    const promptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.PROMPT
    );
    if (!promptSheet) throw new Error(`找不到工作表 ${SHEET_NAMES.PROMPT}`);
    let basePrompt = promptSheet.getRange("C1").getValue();
    if (
      !basePrompt ||
      typeof basePrompt !== "string" ||
      basePrompt.trim() === ""
    ) {
      writeLog("Prompt!C1 為空或無效，使用預設基礎提示詞。");
      basePrompt = "你是一個友善的 AI 助理。";
    }
    const promptToCache = basePrompt.trim();
    cache.put(
      CACHE_KEYS.GLOBAL_BASE_PROMPT,
      promptToCache,
      PROMPT_CACHE_EXPIRATION
    );
    writeLog("從工作表讀取並快取基礎提示詞 (C1)");
    return promptToCache;
  } catch (error) {
    writeLog("讀取基礎提示詞 (C1) 錯誤: " + error + "，使用預設。");
    return "你是一個友善的 AI 助理。";
  }
}

function getFullPrompt(contextId = null) {
  const customPromptNumber = getCurrentPromptNumber(contextId); // ← 傳入 contextId
  const cacheKey = `${CACHE_KEYS.SPECIFIC_PROMPT_PREFIX}${customPromptNumber}`;
  const cache = CacheService.getScriptCache();
  const cachedPrompt = cache.get(cacheKey);
  if (cachedPrompt) return cachedPrompt;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.PROMPT
    );
    if (!sheet) throw new Error(`找不到工作表 ${SHEET_NAMES.PROMPT}`);
    writeLog(`嘗試從工作表獲取特定提示詞，編號 #${customPromptNumber}`);

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      writeLog(`未找到提示詞資料（工作表少於 4 行），使用預設`);
      const defaultPrompt = "請根據對話上下文，以自然、友善的語氣回應。";
      cache.put(cacheKey, defaultPrompt, PROMPT_CACHE_EXPIRATION);
      return defaultPrompt;
    }

    const data = sheet.getRange("A4:C" + lastRow).getValues();
    let specificPromptContent = "";
    let promptName = `編號 ${customPromptNumber}`;

    for (const row of data) {
      if (row[0] && Number(row[0]) === customPromptNumber) {
        promptName =
          row[1] && String(row[1]).trim() ? String(row[1]).trim() : promptName;
        specificPromptContent = row[2] || "";
        break;
      }
    }

    let promptToCache = "";
    if (specificPromptContent.trim() !== "") {
      writeLog(`找到特定提示詞 #${customPromptNumber}: ${promptName}`);
      promptToCache = specificPromptContent.trim();
    } else {
      writeLog(
        `未找到提示詞 #${customPromptNumber} 的有效內容，使用預設特定提示詞。`
      );
      promptToCache = "請根據對話上下文，以自然、友善的語氣回應。";
    }

    cache.put(cacheKey, promptToCache, PROMPT_CACHE_EXPIRATION);
    return promptToCache;
  } catch (error) {
    writeLog("獲取特定提示詞錯誤: " + error + "，使用預設。");
    return "獲取提示詞時發生錯誤，請檢查 Prompt 試算表。";
  }
}

/**
 * 從「個別模式」頁讀取指定 contextId 的設定
 * @returns {Object|null} { promptNumber, modeName, lastUpdated } 或 null
 */
function getIndividualMode(contextId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.INDIVIDUAL_MODE
    );
    if (!sheet) return null;

    const textFinder = sheet
      .getRange("A:A")
      .createTextFinder(contextId)
      .matchEntireCell(true);
    const foundCell = textFinder.findNext();

    if (foundCell) {
      const row = foundCell.getRow();
      const data = sheet.getRange(row, 1, 1, 4).getValues()[0];
      return {
        contextId: data[0],
        promptNumber: Number(data[1]) || 1,
        modeName: data[2] || "",
        lastUpdated: data[3] || "",
      };
    }
    return null;
  } catch (error) {
    writeLog(`getIndividualMode 錯誤 for ${contextId}: ${error}`);
    return null;
  }
}

/**
 * 設定或更新「個別模式」頁的 contextId 記錄
 */
function setIndividualMode(contextId, promptNumber, modeName = "") {
  return withLock_(() => {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        SHEET_NAMES.INDIVIDUAL_MODE
      );
      if (!sheet) {
        writeLog(
          `setIndividualMode: 找不到 ${SHEET_NAMES.INDIVIDUAL_MODE} 工作表`
        );
        return;
      }

      const timestamp = formatDateTime(new Date());
      const textFinder = sheet
        .getRange("A:A")
        .createTextFinder(contextId)
        .matchEntireCell(true);
      const foundCell = textFinder.findNext();

      if (foundCell) {
        // 更新現有記錄
        const row = foundCell.getRow();
        sheet
          .getRange(row, 2, 1, 3)
          .setValues([[promptNumber, modeName, timestamp]]);
        writeLog(
          `已更新 ${contextId.substring(
            0,
            8
          )}*** 的個別模式為 #${promptNumber}: ${modeName}`
        );
      } else {
        // 新增記錄
        sheet.appendRow([contextId, promptNumber, modeName, timestamp]);
        writeLog(
          `已新增 ${contextId.substring(
            0,
            8
          )}*** 的個別模式 #${promptNumber}: ${modeName}`
        );
      }
    } catch (error) {
      writeLog(`setIndividualMode 錯誤 for ${contextId}: ${error}`);
    }
  });
}

/**
 * 取得當前 contextId 的 Prompt 編號
 * 優先讀取「個別模式」頁，無則 fallback 到 Prompt!B1
 */
function getCurrentPromptNumber(contextId = null) {
  try {
    // 若有提供 contextId，先查「個別模式」
    if (contextId) {
      const individualMode = getIndividualMode(contextId);
      if (individualMode && typeof individualMode.promptNumber === "number") {
        return individualMode.promptNumber;
      }
    }

    // Fallback 到全域預設 (Prompt!B1)
    const promptSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.PROMPT
    );
    if (!promptSheet) {
      writeLog("獲取當前提示編號時找不到 Prompt 工作表，使用預設 1");
      return 1;
    }
    const valB1 = promptSheet.getRange("B1").getValue();
    return typeof valB1 === "number" && valB1 > 0 && Number.isInteger(valB1)
      ? valB1
      : 1;
  } catch (e) {
    writeLog("讀取 Prompt 編號失敗，使用預設編號 1: " + e);
    return 1;
  }
}

function clearPromptCache() {
  let clearedKeys = [];
  try {
    const cache = CacheService.getScriptCache();

    cache.remove(CACHE_KEYS.GLOBAL_BASE_PROMPT);
    cache.remove(CACHE_KEYS.PROVIDER);
    cache.remove(CACHE_KEYS.MODEL);
    clearedKeys.push(
      CACHE_KEYS.GLOBAL_BASE_PROMPT,
      CACHE_KEYS.PROVIDER,
      CACHE_KEYS.MODEL
    );

    const currentPromptNumber = getCurrentPromptNumber();
    const currentSpecificCacheKey = `${CACHE_KEYS.SPECIFIC_PROMPT_PREFIX}${currentPromptNumber}`;
    cache.remove(currentSpecificCacheKey);
    clearedKeys.push(currentSpecificCacheKey);

    writeLog("Prompt/Provider/Model 快取已清除 (基礎 + 當前特定)");
    return clearedKeys;
  } catch (error) {
    writeLog("清除 Prompt 快取時出錯: " + error);
    return clearedKeys;
  }
}

function clearSpecificPromptCache(promptNumber) {
  if (
    typeof promptNumber !== "number" ||
    promptNumber <= 0 ||
    !Number.isInteger(promptNumber)
  ) {
    writeLog(`無效的提示編號提供給 clearSpecificPromptCache: ${promptNumber}`);
    return;
  }
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `${CACHE_KEYS.SPECIFIC_PROMPT_PREFIX}${promptNumber}`;
    cache.remove(cacheKey);
    writeLog(`清除了特定 Prompt 快取: ${cacheKey}`);
  } catch (error) {
    writeLog(`清除特定 Prompt 快取 #${promptNumber} 時出錯: ${error}`);
  }
}

// =========================================================================
// Utility Functions (LINE Reply, Logging, Record Saving, etc.)
// =========================================================================

function replyMessage(replyToken, text) {
  try {
    if (
      !replyToken ||
      !text ||
      typeof text !== "string" ||
      text.trim() === ""
    ) {
      writeLog(
        `空訊息或無 replyToken，跳過回覆 (Token: ${
          replyToken ? replyToken.substring(0, 5) + "..." : "N/A"
        })`
      );
      return;
    }
    const trimmedText = text.trim();
    writeLog(
      `準備回覆訊息 (Token: ${replyToken.substring(0, 5)}...)，長度: ${
        trimmedText.length
      }`
    );

    const token = getToken();
    const segments = splitMessage(trimmedText).slice(0, 5);
    const retryKey = buildRetryKey(replyToken + (segments[0] || ""));
    const url = "https://api.line.me/v2/bot/message/reply";
    const options = {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
        "X-Line-Retry-Key": retryKey,
      },
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: segments.map((msg) => ({ type: "text", text: msg })),
      }),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      writeLog(`LINE 回覆成功 (狀態碼 ${responseCode})`);
    } else {
      const responseContent = response.getContentText();
      writeLog(
        `LINE 回覆失敗，狀態碼: ${responseCode}, 內容: ${responseContent}`
      );
    }
  } catch (error) {
    writeLog("回覆 LINE 錯誤: " + error);
  }
}

function pushMessage(userId, text) {
  try {
    if (!ALLOW_PUSH) {
      writeLog("pushMessage: ALLOW_PUSH=false, 攔截");
      return;
    }
    if (!text || !text.trim()) {
      writeLog("pushMessage: 空內容，略過");
      return;
    }
    if (!underPushBudget()) {
      writeLog("PUSH 超過月度上限，已攔截");
      return;
    }

    const trimmedText = text.trim();
    writeLog(
      `準備 push 訊息 to user ${userId.substring(0, 6)}***，長度: ${
        trimmedText.length
      }`
    );

    const token = getToken();
    const segments = splitMessage(trimmedText).slice(0, 5);
    const retryKey = buildRetryKey(userId + (segments[0] || ""));
    const url = "https://api.line.me/v2/bot/message/push";
    const options = {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + token,
        "X-Line-Retry-Key": retryKey,
      },
      payload: JSON.stringify({
        to: userId,
        messages: segments.map((msg) => ({ type: "text", text: msg })),
      }),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      incPushCount(segments.length);
      writeLog(`LINE push 成功 (狀態碼 ${responseCode})`);
    } else {
      const responseContent = response.getContentText();
      writeLog(
        `LINE push 失敗，狀態碼: ${responseCode}, 內容: ${responseContent}`
      );
    }
  } catch (error) {
    writeLog("push LINE 錯誤: " + error);
  }
}

function splitMessage(text) {
  const MAX_LENGTH = LINE_TEXT_MAX;
  const messages = [];
  let currentText = text || "";
  while (currentText.length > 0) {
    if (currentText.length <= MAX_LENGTH) {
      messages.push(currentText);
      break;
    }
    let splitIndex = currentText.lastIndexOf("\n", MAX_LENGTH);
    if (splitIndex === -1 || splitIndex === 0) splitIndex = MAX_LENGTH;
    else splitIndex += 1;
    messages.push(currentText.substring(0, splitIndex).trim());
    currentText = currentText.substring(splitIndex).trim();
  }
  return messages.filter(Boolean);
}

function formatForLineMobile(text) {
  let processed = text;

  // 暴力修正：
  // 1. 只要遇到中文句號(。)、問號(？)、驚嘆號(！)，後面直接強制接兩個換行 (\n\n)
  // 2. (?!\n\n) 是為了避免原本就已經有空行的地方被重複加成超級大空行
  processed = processed.replace(/([。！？])(?!\n\n)/g, "$1\n\n");

  // 3. 清理一下：如果連續換行超過 3 個（變得太空），縮減回 2 個
  processed = processed.replace(/\n{3,}/g, "\n\n");

  return processed.trim();
}

// 可重現 UUID v4（X-Line-Retry-Key）
function buildRetryKey(seed) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed);
  var bytes = [];
  for (var i = 0; i < digest.length; i++) bytes.push((digest[i] + 256) % 256);
  while (bytes.length < 16) bytes.push(0);
  var b = bytes.slice(0, 16);
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  function hex(n) {
    return ("0" + (n & 0xff).toString(16)).slice(-2);
  }
  return (
    hex(b[0]) +
    hex(b[1]) +
    hex(b[2]) +
    hex(b[3]) +
    "-" +
    hex(b[4]) +
    hex(b[5]) +
    "-" +
    hex(b[6]) +
    hex(b[7]) +
    "-" +
    hex(b[8]) +
    hex(b[9]) +
    "-" +
    hex(b[10]) +
    hex(b[11]) +
    hex(b[12]) +
    hex(b[13]) +
    hex(b[14]) +
    hex(b[15])
  );
}

function getToken() {
  const token = PropertiesService.getScriptProperties().getProperty("TOKEN");
  if (!token) {
    writeLog("錯誤：未在 Script Properties 設定 LINE Token (TOKEN)");
    throw new Error("LINE Token not found in Script Properties.");
  }
  return token;
}

function getOpenRouterKey() {
  const key =
    PropertiesService.getScriptProperties().getProperty("OPENROUTER_KEY");
  if (!key) {
    writeLog("警告：未設定 OpenRouter API Key (OPENROUTER_KEY)");
    return null;
  }
  return key;
}

function getXaiApiKey() {
  const key =
    PropertiesService.getScriptProperties().getProperty("XAI_API_KEY");
  if (!key) {
    writeLog("警告：未設定 xAI API Key (XAI_API_KEY)");
    return null;
  }
  return key;
}

function writeLog(message) {
  const timestamp = formatDateTime(new Date());
  console.log(`[LOG] ${message}`);
  if (typeof LOG_BUFFER !== "undefined") {
    LOG_BUFFER.push([timestamp, message]);
  }
}

function flushLogs() {
  if (typeof LOG_BUFFER === "undefined" || LOG_BUFFER.length === 0) return;
  try {
    ensureSsAvailable_();
    const logSheet = ss.getSheetByName(SHEET_NAMES.LOG);
    if (logSheet) {
      logSheet
        .getRange(logSheet.getLastRow() + 1, 1, LOG_BUFFER.length, 2)
        .setValues(LOG_BUFFER);
    }
  } catch (e) {
    console.error("寫入日誌失敗: " + e);
  }
  LOG_BUFFER = [];
}

function formatDateTime(date) {
  try {
    return Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );
  } catch (e) {
    console.error("formatDateTime error:", e);
    return date.toISOString();
  }
}

// 暫存寫入（批次）
function queueRecord(recordData) {
  try {
    const cache = CacheService.getScriptCache();
    const key = `pendingRecords_${Utilities.getUuid()}`;
    cache.put(key, JSON.stringify(recordData), 600);
    const listKey = "pendingRecordKeys";
    let current = cache.get(listKey);
    let keys = current ? JSON.parse(current) : [];
    keys.push(key);
    cache.put(listKey, JSON.stringify(keys), 600);
    writeLog(
      `已加入暫存寫入隊列 (${recordData.role}): ${String(
        recordData.text
      ).substring(0, 30)}...`
    );
  } catch (e) {
    writeLog(`queueRecord 發生錯誤: ${e}`);
  }
}

function flushQueuedRecords() {
  return withLock_(() => {
    try {
      const cache = CacheService.getScriptCache();
      const listKey = "pendingRecordKeys";
      const current = cache.get(listKey);
      if (!current) {
        return;
      }
      const keys = JSON.parse(current);
      if (!Array.isArray(keys) || keys.length === 0) {
        return;
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SHEET_NAMES.RECORDS);
      const now = formatDateTime(new Date());
      const rows = [];

      keys.forEach((k) => {
        const val = cache.get(k);
        if (!val) return;
        try {
          const r = JSON.parse(val);
          if (r && r.text && r.userId && r.groupId) {
            rows.push([
              now,
              r.groupId,
              r.userId,
              r.text,
              r.role,
              r.resetFlag || "",
            ]);
          }
        } catch (err) {
          writeLog(`flushQueuedRecords: 解析 key ${k} 錯誤: ${err}`);
        }
        cache.remove(k);
      });

      if (rows.length > 0) {
        sheet
          .getRange(sheet.getLastRow() + 1, 1, rows.length, 6)
          .setValues(rows);
        writeLog(`flushQueuedRecords: 已批次寫入 ${rows.length} 筆紀錄`);
      }
      cache.remove(listKey);
    } catch (error) {
      writeLog(`flushQueuedRecords 發生錯誤: ${error}`);
    }
  });
}

// 1:1 Loading 動畫
function showLoadingAnimation(userId, seconds) {
  try {
    const duration = Math.max(5, Math.min(60, Number(seconds) || 10));
    const url = "https://api.line.me/v2/bot/chat/loading/start";
    const payload = { chatId: userId, loadingSeconds: duration };
    const options = {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + getToken() },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
    UrlFetchApp.fetch(url, options);
    writeLog(
      `Loading 動畫已啟動 ${duration}s for user ${userId.substring(0, 6)}***`
    );
  } catch (e) {
    writeLog("showLoadingAnimation 發生錯誤: " + e);
  }
}

// Push 配額
function canUsePush(contextId, userId) {
  return (
    ALLOW_PUSH &&
    contextId === userId &&
    underPushBudget() &&
    passUserCooldown(userId, 60)
  );
}

function underPushBudget() {
  const props = PropertiesService.getScriptProperties();
  const tz = Session.getScriptTimeZone();
  const ym = Utilities.formatDate(new Date(), tz, "yyyyMM");
  const curYm = props.getProperty("PUSH_MONTH") || "";
  if (curYm !== ym) {
    props.setProperty("PUSH_MONTH", ym);
    props.setProperty("PUSH_COUNT", "0");
  }
  const cap = Number(props.getProperty("PUSH_CAP") || "300");
  const used = Number(props.getProperty("PUSH_COUNT") || "0");
  return used < cap;
}
function incPushCount(n = 1) {
  const props = PropertiesService.getScriptProperties();
  const used = Number(props.getProperty("PUSH_COUNT") || "0") + n;
  props.setProperty("PUSH_COUNT", String(used));
}
function passUserCooldown(userId, sec = 60) {
  const c = CacheService.getScriptCache();
  const k = "pushCooldown_" + userId;
  if (c.get(k)) return false;
  c.put(k, "1", sec);
  return true;
}

// 命名空間 & 快取工具
function getHistNs_() {
  const props = PropertiesService.getScriptProperties();
  const v = props.getProperty(HIST_NS_PROP_KEY);
  if (!v) {
    props.setProperty(HIST_NS_PROP_KEY, "1");
    return "1";
  }
  return v;
}
function bumpHistNs_() {
  const props = PropertiesService.getScriptProperties();
  const v = Number(getHistNs_() || "1") + 1;
  props.setProperty(HIST_NS_PROP_KEY, String(v));
  return String(v);
}
function buildHistoryKey_(contextId) {
  const promptNum = getCurrentPromptNumber(contextId); // ← 使用個別模式編號
  return `${
    CACHE_KEYS.HISTORY_PREFIX
  }${getHistNs_()}:${promptNum}:${contextId}`;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    return fn();
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {}
  }
}
function safeJsonPutToCache_(cache, key, jsonStr, ttl) {
  const MAX_BYTES = 90 * 1024;
  let s = jsonStr;
  while (Utilities.newBlob(s).getBytes().length > MAX_BYTES) {
    try {
      const arr = JSON.parse(s);
      if (Array.isArray(arr) && arr.length > 2) {
        arr.shift();
        arr.shift();
        s = JSON.stringify(arr);
      } else break;
    } catch (_) {
      break;
    }
  }
  cache.put(key, s, ttl);
}

// 維護：每天清 LOG、建立批次器
function setupMaintenance() {
  const functionName = "cleanOldLogs";
  try {
    let triggerExists = false;
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === functionName) {
        if (!triggerExists) {
          writeLog(
            `找到現有的 ${functionName} 觸發器 (ID: ${trigger.getUniqueId()})`
          );
          triggerExists = true;
        } else {
          writeLog(
            `刪除重複的 ${functionName} 觸發器 (ID: ${trigger.getUniqueId()})`
          );
          ScriptApp.deleteTrigger(trigger);
        }
      }
    }
    if (!triggerExists) {
      ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyDays(1)
        .atHour(3)
        .create();
      writeLog(`已創建每日日誌清理任務 (${functionName} at ~3 AM)`);
      return `每日日誌清理任務 (${functionName}) 已創建。`;
    } else {
      return `每日日誌清理任務 (${functionName}) 已存在。`;
    }
  } catch (error) {
    writeLog(`設置維護任務 (${functionName}) 錯誤: ${error}`);
    return `設置維護任務 (${functionName}) 失敗。`;
  }
}

function cleanOldLogs() {
  const functionName = "cleanOldLogs";
  try {
    writeLog(`--- 開始執行每日日誌清理 (${functionName}) ---`);
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      SHEET_NAMES.LOG
    );
    if (!logSheet) {
      writeLog(`${functionName}: 找不到 LOG 工作表，無法清理。`);
      return;
    }

    const KEEP_ROWS = 200; // 保留最近 200 列
    const lastRow = logSheet.getLastRow();

    if (lastRow <= KEEP_ROWS + 1) {
      // +1 因為第 1 行是表頭
      writeLog(
        `${functionName}: LOG 表僅 ${lastRow} 行，無需清理（保留上限 ${
          KEEP_ROWS + 1
        }）。`
      );
      return;
    }

    const rowsToDelete = lastRow - KEEP_ROWS - 1; // 要刪除的行數（-1 排除表頭）

    // 從第 2 行開始刪除舊資料
    for (let i = 0; i < rowsToDelete; i++) {
      try {
        logSheet.deleteRow(2); // 每次都刪第 2 行（因為刪除後會自動上移）
      } catch (e) {
        writeLog(`${functionName}: 刪除第 2 行時出錯: ${e}`);
      }
    }

    writeLog(
      `${functionName}: 已刪除 ${rowsToDelete} 條舊日誌，保留最近 ${KEEP_ROWS} 列。`
    );
    writeLog(`--- 每日日誌清理完成 (${functionName}) ---`);
  } catch (error) {
    writeLog(`${functionName}: 清理日誌過程中發生錯誤: ${error}`);
  }
}

function initialize() {
  const functionName = "initialize";
  try {
    writeLog(`--- 開始執行初始化 (${functionName}) ---`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ensureSheet = (
      name,
      headerRow = [],
      initialData = [],
      frozenRows = 1
    ) => {
      let sheet = ss.getSheetByName(name);
      let created = false;
      if (!sheet) {
        sheet = ss.insertSheet(name);
        created = true;
        writeLog(`已創建工作表: ${name}`);
      }
      if (headerRow.length > 0) {
        if (
          created ||
          sheet.getLastRow() < 1 ||
          sheet.getRange(1, 1, 1, headerRow.length).getValues()[0].join("") ===
            ""
        ) {
          sheet
            .getRange(1, 1, 1, headerRow.length)
            .setValues([headerRow])
            .setFontWeight("bold");
          if (frozenRows > 0)
            try {
              sheet.setFrozenRows(frozenRows);
            } catch (_) {}
        }
      }
      if (created && initialData.length > 0)
        initialData.forEach((row) => sheet.appendRow(row));
      if (name === SHEET_NAMES.RECORDS && sheet.getMaxColumns() < 6) {
        try {
          sheet.insertColumnsAfter(
            sheet.getMaxColumns(),
            6 - sheet.getMaxColumns()
          );
          sheet.getRange("F1").setValue("Reset Flag").setFontWeight("bold");
        } catch (_) {}
      }
      return sheet;
    };

    ensureSheet(
      SHEET_NAMES.RECORDS,
      ["時間", "對話 ID", "用戶 ID", "內容", "角色", "Reset Flag"],
      [],
      1
    );
    ensureSheet(SHEET_NAMES.LOG, ["時間", "訊息"], [], 1);
    ensureSheet(
      SHEET_NAMES.LAST_CONVERSATION,
      ["對話 ID (Context)", "歷史紀錄 (JSON)"],
      [],
      1
    );

    // ========== 新增：個別模式工作表 ==========
    ensureSheet(
      SHEET_NAMES.INDIVIDUAL_MODE,
      ["Context ID", "Prompt 編號", "模式名稱", "最後更新時間"],
      [],
      1
    );

    const promptSheet = ensureSheet(SHEET_NAMES.PROMPT, [], [], 2);

    // ========== A1: Provider ==========
    if (!String(promptSheet.getRange("A1").getValue() || "").trim()) {
      promptSheet
        .getRange("A1")
        .setValue("XAI")
        .setNote("供應商：XAI 或 OPENROUTER");
    }

    // ========== A2: Model（只檢查一次）==========
    if (!String(promptSheet.getRange("A2").getValue() || "").trim()) {
      const provider = String(
        promptSheet.getRange("A1").getValue() || "XAI"
      ).toUpperCase();
      const defaultModel =
        provider === "OPENROUTER" ? "openai/gpt-4o-mini" : "grok-4-fast";
      promptSheet
        .getRange("A2")
        .setValue(defaultModel)
        .setNote("模型名稱：依供應商填入相容模型");
    }

    // ========== B1: 全域預設 Prompt 編號 ==========
    const b1 = promptSheet.getRange("B1").getValue();
    if (!(typeof b1 === "number" && b1 > 0 && Number.isInteger(b1))) {
      promptSheet
        .getRange("B1")
        .setValue(1)
        .setNote("全域預設提示詞編號（個別模式未設定時使用）");
    }

    // ========== C1: Base Prompt ==========
    if (!String(promptSheet.getRange("C1").getValue() || "").trim()) {
      promptSheet
        .getRange("C1")
        .setValue("你是一個友善的 AI 助理。")
        .setNote("通用的基礎提示詞");
    }

    // ========== A3:C3 提示詞列表表頭（避免與 A2 Model 衝突）==========
    if (
      promptSheet.getLastRow() < 3 ||
      promptSheet.getRange("A3:C3").getValues()[0].join("") === ""
    ) {
      promptSheet
        .getRange("A3:C3")
        .setValues([["提示詞編號", "提示詞名稱", "提示詞內容"]])
        .setFontWeight("bold");
    }

    // ========== A4 開始：預設提示詞範例 ==========
    if (promptSheet.getLastRow() < 4) {
      promptSheet.appendRow([
        1,
        "預設助理模式",
        "你是個友善且樂於助人的 AI 助理。",
      ]);
    }

    const maintResult = setupMaintenance();
    setupRecordFlusher();

    writeLog(`${functionName} 完成。${maintResult}`);
    writeLog(`--- 初始化完成 (${functionName}) ---`);
    return `${functionName} 完成。`;
  } catch (error) {
    const errorMsg =
      `${functionName} 過程中發生嚴重錯誤: ${error}` +
      (error.stack ? "\nStack: " + error.stack : "");
    try {
      writeLog(errorMsg);
    } catch (e) {
      console.error(errorMsg);
    }
    return errorMsg;
  }
}

function setupRecordFlusher() {
  const funcName = "flushQueuedRecords";
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some((t) => t.getHandlerFunction() === funcName);
  if (!exists) {
    ScriptApp.newTrigger(funcName).timeBased().everyMinutes(1).create();
    writeLog(`已建立每分鐘批次寫入觸發器 (${funcName})`);
  }
}

// =========================================================================
// LINE Webhook Entry Point (doPost)
// =========================================================================

/**
 * LINE Messaging API Webhook 接收函數
 * 當 LINE 伺服器向你的 Web App URL 發送 POST 請求時會呼叫此函數
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      writeLog("doPost: 收到空的 POST 請求，略過");
      return ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "Empty request" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const json = JSON.parse(e.postData.contents);

    // Webhook 簽名驗證（建議啟用以防偽造請求）
    // 若要啟用，取消下面註解並在 Script Properties 設定 CHANNEL_SECRET
    /*
    const signature = e.parameter['X-Line-Signature'] || (e.headers ? e.headers['X-Line-Signature'] || e.headers['x-line-signature'] : null);
    const channelSecret = PropertiesService.getScriptProperties().getProperty("CHANNEL_SECRET");
    if (channelSecret && signature) {
      const hash = Utilities.computeHmacSha256Signature(e.postData.contents, channelSecret);
      const expectedSignature = Utilities.base64Encode(hash);
      if (signature !== expectedSignature) {
        writeLog("doPost: Webhook 簽名驗證失敗，拒絕請求");
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid signature" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    */

    if (
      !json.events ||
      !Array.isArray(json.events) ||
      json.events.length === 0
    ) {
      writeLog("doPost: 無事件陣列，略過");
      return ContentService.createTextOutput(
        JSON.stringify({ status: "ok", message: "No events" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    json.events.forEach((event) => {
      try {
        // 事件去重（避免 LINE 重發導致重複處理）
        if (event.webhookEventId && isDuplicateEvent(event.webhookEventId)) {
          writeLog(`重複事件 ID ${event.webhookEventId}，略過`);
          return;
        }

        // 只處理訊息事件
        if (event.type !== "message") {
          writeLog(`略過非訊息事件: ${event.type}`);
          return;
        }

        // 只處理文字訊息
        if (event.message.type !== "text") {
          writeLog(`略過非文字訊息: ${event.message.type}`);
          return;
        }

        const userMessage = event.message.text;
        const userId = event.source.userId;
        const replyToken = event.replyToken;

        // contextId：群組/房間用 groupId/roomId，1:1 用 userId
        let contextId = userId;
        if (event.source.type === "group" && event.source.groupId) {
          contextId = event.source.groupId;
        } else if (event.source.type === "room" && event.source.roomId) {
          contextId = event.source.roomId;
        }

        writeLog(
          `收到訊息事件: userId=${userId}, contextId=${contextId}, text='${userMessage.substring(
            0,
            30
          )}...'`
        );

        // 呼叫核心處理函數
        handleMessage(userMessage, userId, replyToken, contextId);
      } catch (eventError) {
        writeLog(
          `處理事件錯誤: ${eventError}` +
            (eventError.stack ? `\nStack: ${eventError.stack}` : "")
        );
      }
    });

    return ContentService.createTextOutput(
      JSON.stringify({ status: "ok" })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    writeLog(
      `doPost 錯誤: ${error}` + (error.stack ? `\nStack: ${error.stack}` : "")
    );
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    flushLogs();
  }
}

/**
 * 事件去重：用快取記錄已處理的 webhookEventId（60 秒 TTL）
 */
function isDuplicateEvent(eventId) {
  const cache = CacheService.getScriptCache();
  const key = `event_${eventId}`;
  const exists = cache.get(key);
  if (exists) return true;
  cache.put(key, "1", 60);
  writeLog(`新事件 ID ${eventId}，加入快取 60 秒`);
  return false;
}
