function analyzeAndGenerateReport() {
  const ss = SpreadsheetApp.openById(<SpreadSheetID>); // 替換為你的 Google Sheet ID
  const sheet = ss.getActiveSheet();

  const data = sheet.getDataRange().getValues();
  const headerRow = data[0];
  const targetColumnIndex = headerRow.indexOf("[選填] 你期待後續活動上，想要聽什麼樣的主題內容？");

  if (targetColumnIndex === -1) {
    throw new Error("找不到目標欄位，請確認欄位名稱。");
  }

  const allResponses = [];
  for (let i = 1; i < data.length; i++) {
    const entry = data[i][targetColumnIndex];
    if (entry) {
      allResponses.push(entry.trim());
    }
  }

  const topTopics = requestGeminiSummarization(allResponses);
  if (topTopics.length === 0) {
    throw new Error("未能從 Gemini API 獲取主題摘要。");
  }

  const reportSheetName = "十大主題";
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(reportSheetName);
  } else {
    reportSheet.clear();
  }

  reportSheet.appendRow(["主題", "細節"]);
  topTopics.forEach(topic => {
    reportSheet.appendRow([topic.topic, topic.detail]);
  });
}

// 修正的 Gemini API 請求函數
function requestGeminiSummarization(responses) {
  const apiKey = <GoogleGeminitToken>; // 替換為你的 Gemini API Key
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;

  // 限制 prompt 長度並進行合併
  const promptText = responses.slice(0, 100).join("\n"); // 根據需要調整分批數量
  // Prepare the payload
  var payload = {
    contents: [{
      parts: [{
        text: "請從以下內容中，以台灣人習慣使用的繁體中文，統整最多人關注的十大主題，並輸出為JSON格式：[{\"topic\":\"...\", \"detail\":\"...\"}, {\"topic\":\"...\", \"detail\":\"...\"}, {\"topic\":\"...\", \"detail\":\"...\"}]： \n" + promptText
      }]
    }]
  };

  Logger.log('Gemini API 請求 Payload: ' + JSON.stringify(payload));

  // Make the POST request
  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const jsonResponse = JSON.parse(response.getContentText());
    const content = jsonResponse?.candidates?.[0]?.content?.parts?.[0]?.text;

    if (content) {
      const cleanJson = content.replace(/```json|```/g, "").trim();
      Logger.log('清理後的 JSON 字串: ' + cleanJson);
      const parsedJson = JSON.parse(cleanJson);
      return parsedJson;
    }
    return [];
  } catch (e) {
    Logger.log('Gemini API 請求失敗或解析錯誤: ' + e.message);
    return [];
  }
}
