// Google App Script to analyze and summarize form responses with Google Gemini API
function analyzeQuestionsWithGemini() {
  // Define the Spreadsheet and Sheet
  var spreadsheet = SpreadsheetApp.openById(<SpreadSheetID>); // Replace with your Google Sheet ID
  var sheet = spreadsheet.getSheets()[0]; // Access the first sheet

  // Locate the target column by header
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var targetHeader = '在今天的演講中，你想要詢問什麼問題？';
  var targetColumnIndex = headers.indexOf(targetHeader) + 1;

  if (targetColumnIndex === 0) {
    Logger.log('未找到目標欄位: ' + targetHeader);
    return;
  }

  // Collect all questions from the target column
  var lastRow = sheet.getLastRow();
  var questions = sheet.getRange(2, targetColumnIndex, lastRow - 1, 1).getValues()
                      .flat()
                      .filter(String); // Remove empty entries

  if (questions.length === 0) {
    Logger.log('目標欄位中沒有任何問題。');
    return;
  }

  Logger.log('提取的問題: ' + JSON.stringify(questions));

  // Request summarization from Google Gemini API
  var summarizedQuestions = requestGemini(questions);

  if (summarizedQuestions.length === 0) {
    Logger.log('Gemini API 未返回任何總結問題。');
    return;
  }

  Logger.log('總結的問題: ' + JSON.stringify(summarizedQuestions));

  // Create or access the target sheet for storing results
  var resultSheetName = '今日回答題目';
  var resultSheet = spreadsheet.getSheetByName(resultSheetName);
  if (!resultSheet) {
    resultSheet = spreadsheet.insertSheet(resultSheetName);
  } else {
    resultSheet.clear(); // Clear existing data
  }

  // Write results to the new sheet
  resultSheet.getRange(1, 1).setValue('建議問答題目');
  summarizedQuestions.forEach(function(question, index) {
    resultSheet.getRange(index + 2, 1).setValue((index + 1) + '. ' + question);
  });

  Logger.log('建議問答題目已記錄至 "' + resultSheetName + '" 工作表中。');
}

// Function to request summarization from Google Gemini API
function requestGemini(questions) {
  var apiKey = <GoogleGeminitID>; // Replace with your Gemini API key
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;

  // Prepare the payload
  var payload = {
    contents: [{
      parts: [{
        text: "根據以下聽者的問題，以台灣人習慣使用的繁體中文，統整最多人問到或最適合解決的三個問題，輸出格式是[{\"question\":\"...\"}, {\"question\":\"...\"}, {\"question\":\"...\"}]： " + questions.join(" | ")
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
    var response = UrlFetchApp.fetch(url, options);
    var jsonResponse = JSON.parse(response.getContentText());

    Logger.log('Gemini API 回應: ' + JSON.stringify(jsonResponse));

    // Parse the JSON returned in the content
    var contentParts = jsonResponse.candidates?.[0]?.content?.parts?.[0]?.text;
    if (contentParts) {
      // Remove the surrounding code block formatting if exists
      var cleanJson = contentParts.replace(/```json|```/g, "").trim();

      // Replace single quotes with double quotes to make JSON valid
      cleanJson = cleanJson.replace(/'/g, '"');

      // Parse the cleaned JSON
      var parsedJson = JSON.parse(cleanJson);

      // Extract top three questions
      var topQuestions = parsedJson.map(function(q) {
        return q.question;
      });

      return topQuestions; // Return the top three questions
    }
    return [];
  } catch (e) {
    Logger.log('Gemini API 請求失敗: ' + e.message);
    return [];
  }
}
