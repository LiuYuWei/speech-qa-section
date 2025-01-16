// Google App Script to create a Google Form
function createLectureFeedbackForm() {
  // Create a new form
  var form = FormApp.create('20250117 - Google 社群年終會 演講問卷調查表')
                  .setDescription('請幫忙填寫以下問卷，協助我們了解您對今天演講的看法和未來活動的期待！');

  // Add a paragraph text question for future topic suggestions
  form.addParagraphTextItem()
      .setTitle('在今天的演講中，你想要詢問什麼問題？')
      .setRequired(true);

  // Add a text question for additional comments
  form.addParagraphTextItem()
      .setTitle('[選填] 你期待後續活動上，想要聽什麼樣的主題內容？')
      .setRequired(false);

  // Add a text question for additional comments
  form.addParagraphTextItem()
      .setTitle('[選填] 有什麼想要回饋給我')
      .setRequired(false);

  // Log the URL of the created form
  Logger.log('表單已建立！表單連結：' + form.getEditUrl());

  // Display the link to the user
  return form.getPublishedUrl();
}
