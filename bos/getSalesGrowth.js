function createTrigger() {
  ScriptApp.newTrigger('myFunction')
      .timeBased()
      .everyMinutes(2)  //毎分
      .create();   //トリガー
}

function getList() {
  createTrigger();
}


function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var count = sheet.getRange(2, 26).getValue() + 1;
  var ticker = sheet.getRange(count, 1).getValue().toString();
  var stockExchange = ['PINX','xnys', 'xnas'];
  var targetExchange = '';
  var byId = '';
  for (var i =0; i < stockExchange.length; i++) {
    byId = getById(ticker, stockExchange[i]);
    targetExchange = stockExchange[i];
    if (byId !== '') {
      break;
    }
  }
  if (byId !== '') {
    writeExchange(sheet, targetExchange, count);
    var fullKeyRatio = 'https://financials.morningstar.com/ratios/r.html?t='+ byId +'&culture=en&platform=sal';
    writeUrl(sheet, count, fullKeyRatio);
    var financialsUrl = "https://financials.morningstar.com/finan/financials/getFinancePart.html?&callback=jsonp1580192904311&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1580192905566"
    financials = UrlFetchApp.fetch(financialsUrl).getContentText();
    var keyRatio;
    var keyRatioUrl = "https://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=jsonp1579473219364&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1579473220658";
    keyRatio = UrlFetchApp.fetch(keyRatioUrl).getContentText();

    var revenueGrowth = getRevenueGrowth(keyRatio);
    var revenue = getRevenue(financials);
    var bps = getBookValue(financials);
    var grossMargin = getGrossMargin(financials);

    writeRevenueGrowth(sheet, revenueGrowth, count);
    writeRevenue(sheet, revenue, count);
    writeBPS(sheet, bps, count);
    writeGrossMargin(sheet, grossMargin, count);
  }
  sheet.getRange(2, 26).setValue(count);
}

function getGrossMargin(financials) {
  var grossMarginRegExp = /<td align=\\"right\\" headers=\\"Y9 i1\\">(.*?)</g;
  var grossMarginTags = financials.match(grossMarginRegExp);
  var grossMarginList = []
  for(var i = 0; i < grossMarginTags.length; i++) {
    grossMarginList.push(grossMarginTags[i].match(/>(.*?)</)[1]);
  }
  return grossMarginList;
}

function getRevenue(financials) {
  var revenueRegExp = /<td align=\\"right\\" headers=\\"Y9 i0\\">(.*?)</g;
  var revenueTags = financials.match(revenueRegExp);
  var revenueList = []
  for(var i = 0; i < revenueTags.length; i++) {
    revenueList.push(revenueTags[i].match(/>(.*?)</)[1]);
  }
  return revenueList;
}

function getBookValue(financials) {
  var bookValueRegExp = /<td align=\\"right\\" headers=\\"Y9 i8\\">(.*?)</g;
  var bookValueTags = financials.match(bookValueRegExp);
  var bookValueList = []
  for(var i = 0; i < bookValueTags.length; i++) {
    bookValueList.push(bookValueTags[i].match(/>(.*?)</)[1]);
  }
  return bookValueList;
}

function writeExchange(sheet, targetExchange, count) {
  sheet.getRange(count, 2).setValue(targetExchange);
}

function extractColName(htmlTag) {
  var colNames = htmlTag.match(/>(.*?)</g);
  var colName = "";
  for(var i = 0; i < colNames.length; i++) {
    colName += colNames[i];
  }
  colName = colName.replace(/>/g, "");
  colName = colName.replace(/</g, "");
  colName = colName.replace("&nbsp;", " ");
  return colName;
}

function replaceToDashForList(list) {
  var replacedList = [];
  for (var i = 0; i < list.length; i++) {
    replacedList.push(list[i].replace("&mdash;", "-"))
  }
  return replacedList;
}

function replaceToDash(string) {
  return string.replace("&mdash;", "-");
}

function replaceDashToZero(string) {
  return string.replace("&mdash;", "0");
}

function replaceToNumber(str) {
  return Number(str.replace(/,/g, ''));
}

function replaceToFloat(str) {
  return parseFloat(str.replace(/,/g, ''));
}

function getById(ticker, stockExchange) {
  var byId = new Array();
  try {
  var mStarQuoteUrl = "https://www.morningstar.com/stocks/"+ stockExchange +"/"+ ticker.toLowerCase() +"/quote";
  var mStarQuote = UrlFetchApp.fetch(mStarQuoteUrl).getContentText();
  var byIdRegExp = /byId:{\"(.+)\":.}/;
  var byId = mStarQuote.match(byIdRegExp)
  } catch (e) {
    targetExchange = '';
    byId = ['',''];
  }
  return byId[1];
}


function writeRevenueGrowth(sheet, revenueGrowth, count) {
  for (var i = 0; i < revenueGrowth.length; i++) {
    sheet.getRange(count, i + 5).setValue(revenueGrowth[i]);
  }
}

function writeRevenue(sheet, revenue, count) {
  for (var i = 0; i < revenue.length; i++) {
    sheet.getRange(count, 9).setValue(revenue[i]);
  }
}

function writeBPS(sheet, bps, count) {
  for (var i = 0; i < bps.length; i++) {
    sheet.getRange(count, 10).setValue(bps[i]);
  }
}

function writeGrossMargin(sheet, grossMargin, count) {
  for (var i = 0; i < grossMargin.length; i++) {
    sheet.getRange(count, 11).setValue(grossMargin[i]);
  }
}

function writeUrl(sheet, count, fullKeyRatio) {
  sheet.getRange(count, 12).setValue(fullKeyRatio);
}

function getRevenueGrowth(keyRatio) {
  var yoyRegExp = /<td align=\\"right\\" headers=\\"gr-Y9 gr-revenue i28\\">(.*?)</g;
  var threeYrRegExp = /<td align=\\"right\\" headers=\\"gr-Y9 gr-revenue i29\\">(.*?)</g;
  var fiveYrRegExp = /<td align=\\"right\\" headers=\\"gr-Y9 gr-revenue i30\\">(.*?)</g;
  var tenYrRegExp = /<td align=\\"right\\" headers=\\"gr-Y9 gr-revenue i31\\">(.*?)</g;

  var yoy = keyRatio.match(yoyRegExp)[0];
  var threeYr = keyRatio.match(threeYrRegExp)[0];
  var fiveYr = keyRatio.match(fiveYrRegExp)[0];
  var tenYr = keyRatio.match(tenYrRegExp)[0];

  var growthTags = [yoy, threeYr, fiveYr, tenYr];
  var growthPercentileList = [];
  for(var i = 0; i < growthTags.length; i++) {
    var replacedResult = replaceToDash(growthTags[i]);
    growthPercentileList.push(replacedResult.match(/>(.*?)</)[1]);
  }
  return growthPercentileList;
}
