function createTrigger() {
  ScriptApp.newTrigger('getDataFromOtherSite')
      .timeBased()
      .everyMinutes(1)  //毎分
      .create();   //トリガー
}

function deleteTriggerAll() {
  var allTriggers = ScriptApp.getScriptTriggers();
  for(var i=0; i < allTriggers.length; i++) {
      ScriptApp.deleteTrigger(allTriggers[i]);
  }
}

function getLatestRevenue() {
  createTrigger();
}

function getDataFromOtherSite() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var counter = sheet.getRange(3, 24);
  var count = counter.getValue() + 1;
  counter.setValue(count);
  var cell = sheet.getRange(count + 3, 4)
  var ticker = cell.getValue().toString();
  var financialsUrl = "https://www.marketwatch.com/investing/stock/" + ticker + "/financials";
  financials = UrlFetchApp.fetch(financialsUrl).getContentText();
  var revenue = getRevenue(financials);
  sheet.getRange(count + 3, 17).setValue(revenue);
  count++;
  if (cell.isBlank()) {
    deleteTriggerAll();
  }
}

function getRevenue(financials) {
  var htmlData = financials.match(/<tr class="partialSum">[\s\S]*<\/tr/g)[0];
  var revenueRegExp = /<div class="miniGraph" data-chart=".*(.*?)]}/;
  var revenueData = htmlData.match(revenueRegExp)[0];
  var latestRevenue = revenueData.match(/\d+/g)[4];
  return latestRevenue;
}