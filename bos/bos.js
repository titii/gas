var no = "なし　No";
var yes = "ある　Yes";
var intCovRatioValues = [">  10", "> 4"];
var netProfitMarginValues = [">20%", ">10%"];
var searchError = "エラーです。次のことを順に確認してください。\\n TickerやExchangesを変更した後フォーカスを外しましたか？ \\n TickerやExchangesは正確ですが？ \\n モーニングスターにデータはありますか？"


function myFunction() {
  clearData();
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var ticker = sheet.getRange(2, 2).getValue().toString();
  var stockExchange = getStockExchange(sheet);
  var byId = getById(ticker, stockExchange);
  var financials;
  var keyRatio;
  var financialsUrl = "https://financials.morningstar.com/finan/financials/getFinancePart.html?&callback=jsonp1580192904311&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1580192905566"
  financials = UrlFetchApp.fetch(financialsUrl).getContentText();
  var keyRatioUrl = "https://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=jsonp1579473219364&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1579473220658";
  keyRatio = UrlFetchApp.fetch(keyRatioUrl).getContentText();
  

  var epsList = getEPS(financials);
  var freeCashFlowList = getFreeCashFlow(financials);
  var dividendsList = getDividends(financials);
  var roeList = getROE(keyRatio);
  var interestCoverageList = getInterestCoverage(keyRatio);
  var netProfitMarginList = getNetProfitMargin(keyRatio);
  var growthPercentile = getEPSGrowth(keyRatio);
  var bookValueList = getBookValue(financials);

  clearFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList);
  writeFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList);
  writeEPSGrowth(sheet, growthPercentile);
  writeBookValue(sheet, bookValueList);
  assessFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList);
  collectAssessedData(sheet);
}

function addShoppingList() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ticker = sheet.getRange(2, 2).getValue().toString();
  var result = sheet.getRange("A20:AQ20").getValues();
  var shoppingList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ShoppingList');
  var row = findRow(shoppingList, ticker);
  if (row) { // 行が見つかったら更新
    var range = shoppingList.getRange("A" + row + ":" + "AQ" + row);
    range.setValues(result);
  } else { // 行が見つからなかったら新しくデータを挿入
    var lastrow = shoppingList.getLastRow()+1;
    shoppingList.appendRow(result[0]);
  }
}

function clearAll() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var target0 = sheet.getRange("B2:B3");
  var target1 = sheet.getRange("A6:K11");
  var target2 = sheet.getRange("B14:D14");
  var target3 = sheet.getRange("A16:K16");
  var target4 = sheet.getRange("C20:C20");
  var target5 = sheet.getRange("G20:P20");
  target0.clearContent();
  target1.clearContent();
  target2.clearContent();
  target3.clearContent();
  target4.clearContent();
  target5.clearContent();
}

function updateDataForDbSheet(dbSheet, ticker, score) {
  var values = dbSheet.getDataRange().getValues();

  for (var i = values.length - 1; i > 0; i--) {
    if (values[i][0] === ticker) {
      var row = Number(i + 1);// 配列のキーは0から始まり、行数は1から始まるのでズレを直す
      var addedThreeToScore = parseFloat(score) + 3.0;
      var sum = parseFloat(values[i][1]) + addedThreeToScore;
      var count = values[i][2] + 1;
      var avg = sum / 2;
      return [ticker, avg, count, row];
    }
  }
  return [];
}


function collectAssessedData(sheet) {
  var db = SpreadsheetApp.openById("1jg6nrhMRbzNLoO90LzSogaqjlOm9CBaLaCgk9rpWG0k");
  var dbSheet = db.getSheetByName('hist');
  var dat = dbSheet.getDataRange().getValues();
  Utilities.sleep(1 * 1000);
  var tickerFromActiveSheet = sheet.getRange(2, 2).getValue().toString();
  var scoreFromActiveSheet = sheet.getRange(20, 2).getValue().toString();

  var targetRow = updateDataForDbSheet(dbSheet, tickerFromActiveSheet, scoreFromActiveSheet);

  if (targetRow.length > 0) {
    dbSheet.getRange(targetRow[3], 1, 1, targetRow.length).setValues([targetRow]);
  } else {
    var score = Number(scoreFromActiveSheet) + 3;
    dbSheet.appendRow([tickerFromActiveSheet, score, 1]);
  }
}

function getStockExchange(sheet) {
  return sheet.getRange(3, 2).getValue();
}

function clearFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList) {
  var fundamentals = [epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList];
  var rows = fundamentals.length;
  // sheet分を-1
  var cols = fundamentals[0].length;

  sheet.getRange(6,1,rows,cols).clearContent();
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
    Browser.msgBox(searchError);
  }
  return byId[1];
}

function writeEPSGrowth(sheet, growthPercentile) {
  for (var i = 0; i < growthPercentile.length; i++) {
    sheet.getRange(14, i + 2).setValue(growthPercentile[i]);
  }
}

function writeFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList) {
  var fundamentals = [epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList];
  var rows = fundamentals.length;
  var cols = fundamentals[0].length;

  var replacedfundamentals = [];
  for (var i = 0; i < fundamentals.length; i++) {
    var replacedResult = replaceToDashForList(fundamentals[i]);
    replacedfundamentals.push(replacedResult);
  }

  sheet.getRange(6,1,rows,cols).setValues(replacedfundamentals);
}

function writeBookValue(sheet, bookValueList) {
  var cols = bookValueList.length;

  var replacedBookValueList = [];
  for (var i = 0; i < bookValueList.length; i++) {
    var replacedResult = replaceToDash(bookValueList[i]);
    sheet.getRange(16,i + 1).setValue(replacedResult);
  }
}

function assessFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList) {
  var epsResult = assessEPS(epsList);
  var fcfResult = assessFreeCashFlow(freeCashFlowList);
  var dividendsResult = assessDividends(dividendsList);
  var roeResult = assessROE(roeList);
  var icResult = assessIC(interestCoverageList);
  var netProfitMarginResult = assessNPM(netProfitMarginList);

  sheet.getRange(20,7).setValue(epsResult);
  sheet.getRange(20,8).setValue(fcfResult);
  sheet.getRange(20,9).setValue(roeResult);
  sheet.getRange(20,10).setValue(icResult);
  sheet.getRange(20,12).setValue(netProfitMarginResult);
  sheet.getRange(20,13).setValue(dividendsResult);
}

function assessEPS(epsList) {
  for(var i = 1; i < epsList.length; i++) {
    if(replaceToFloat(epsList[i]) < 0) {
        return no;
      }
  }
  return yes;
}

function assessFreeCashFlow(freeCashFlowList) {
  for(var i = 1; i < freeCashFlowList.length; i++) {
    if(replaceToNumber(freeCashFlowList[i]) < 0) {
        return no;
      }
  }
  return yes;
}

function assessDividends(dividendsList) {
  var dividends = 0;
  for(var i = 1; i < dividendsList.length; i++) {
    var value = replaceToFloat(dividendsList[i]);
    if(value < 0) {
        dividends = 0;
      } else {
        dividends = value;
      }
  }

  if (dividends == 0) {
    return no;
  } 
  else if (0 < dividends) {
    return yes;
  }
}

function assessROE(roeList) {
  for(var i = 1; i < roeList.length; i++) {
    if(replaceToFloat(roeList[i]) < 15) {
        return no;
      }
  }
  return yes;
}

function assessIC(interestCoverageList) {
  var ic = ">  10";
  for(var i = 1; i < interestCoverageList.length; i++) {
    if(interestCoverageList[i] === "&mdash;" || replaceToFloat(interestCoverageList[i]) > 10) {
      ic = ">  10";
    }
    else if(interestCoverageList[i] > 4) {
      ic = "> 4";
    }
    else {
      ic = no;
    }

  }
  return ic;
}

function assessNPM(netProfitMarginList) {
  var npm = ">20%";

  for(var i = 1; i < netProfitMarginList.length; i++) {
    if(npm !== no && npm !== ">10%" && replaceToFloat(netProfitMarginList[i]) > 20) {
      npm = ">20%";
    }
    else if(npm !== no && replaceToFloat(netProfitMarginList[i]) > 10) {
      npm = ">10%";
    }
    else {
      npm = no;
    }
  }
  return npm;
}

function getEPS(financials) {
  var epsRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i5\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i5\\">(.*?)<\\\/th>/g;
  var epsTags = financials.match(epsRegExp);
  var epsList = [];
  var htmlTag = financials.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  epsList.push(colName);

  // Removed TTM by length - 1
  for(var i = 0; i < epsTags.length - 1; i++) {
    epsList.push(epsTags[i].match(/>(.*?)</)[1]);
  }
  return epsList;
}

function getEPSGrowth(keyRatio) {
  var threeYrRegExp = /<td align=\\"right\\" headers=\\"gr-Y9 gr-eps i37\\">(.*?)</g;
  var fiveYrRegExp = /<td align=\\"right\\" headers=\\"gr-Y9 gr-eps i38\\">(.*?)</g;
  var tenYrRegExp = /<td align=\\"right\\" headers=\\"gr-Y9 gr-eps i39\\">(.*?)</g;

  var threeYr = keyRatio.match(threeYrRegExp)[0];
  var fiveYr = keyRatio.match(fiveYrRegExp)[0];
  var tenYr = keyRatio.match(tenYrRegExp)[0];

  var growthTags = [threeYr, fiveYr, tenYr];
  var growthPercentileList = [];
  for(var i = 0; i < growthTags.length; i++) {
    var replacedResult = replaceToDash(growthTags[i]);
    growthPercentileList.push(replacedResult.match(/>(.*?)</)[1]);
  }
  return growthPercentileList;
}

function getFreeCashFlow(financials) {
  var fcfRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i11\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i11\\">(.*?)<\\\/th>/g;
  var fcfTags = financials.match(fcfRegExp);
  var fcfList = []
  var htmlTag = financials.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  fcfList.push(colName);
  // Removed TTM by length - 1
  for(var i = 0; i < fcfTags.length -1; i++) {
    fcfList.push(fcfTags[i].match(/>(.*?)</)[1]);
  }
  return fcfList;
}

function getDividends(financials) {
  var dividendsRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i6\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i6\\">(.*?)<\\\/th>/g;
  var dividendsTags = financials.match(dividendsRegExp);
  var dividendsList = [];
  var htmlTag = financials.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  dividendsList.push(colName);
  // Removed TTM by length - 1
  for(var i = 0; i < dividendsTags.length -1; i++) {
    dividendsList.push(dividendsTags[i].match(/>(.*?)</)[1]);
  }
  return dividendsList;
}

function getBookValue(financials) {
  var bookValueRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i8\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i8\\">(.*?)<\\\/th>/g;
  var bookValueTags = financials.match(bookValueRegExp);
  var bookValueList = [];
  var htmlTag = financials.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  bookValueList.push(colName);
  // Removed TTM by length - 1
  for(var i = 0; i < bookValueTags.length -1; i++) {
    bookValueList.push(bookValueTags[i].match(/>(.*?)</)[1]);
  }
  return bookValueList;
}

function getROE(keyRatio) {
  var roeRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i26\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i26\\">(.*?)<\\\/th>/g;
  var roeTags = keyRatio.match(roeRegExp);
  var roeList = []
  var htmlTag = keyRatio.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  roeList.push(colName);
  // Removed TTM by length - 1
  for(var i = 0; i < roeTags.length -1; i++) {
    roeList.push(roeTags[i].match(/>(.*?)</)[1]);
  }
  return roeList;
}

function getInterestCoverage(keyRatio) {
  var icRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i95\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i95\\">(.*?)<\\\/th>/g;
  var icTags = keyRatio.match(icRegExp);
  var icList = []
  var htmlTag = keyRatio.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  icList.push(colName);
  // Removed TTM by length - 1
  for(var i = 0; i < icTags.length -1; i++) {
    icList.push(icTags[i].match(/>(.*?)</)[1]);
  }
  return icList;
}

function getNetProfitMargin(keyRatio) {
  var npmRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i22\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i22\\">(.*?)<\\\/th>/g;
  var npmTags = keyRatio.match(npmRegExp);
  var npmList = []
  var htmlTag = keyRatio.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  npmList.push(colName);
  // Removed TTM by length - 1
  for(var i = 0; i < npmTags.length -1; i++) {
    npmList.push(npmTags[i].match(/>(.*?)</)[1]);
  }
  return npmList;
}

function findRow(sheet, ticker) {
  var values = sheet.getDataRange().getValues();

  for (var i = values.length - 1; i > 0; i--) {
    if (values[i][0] === ticker) {
      var row = i + 1;// 配列のキーは0から始まり、行数は1から始まるのでズレを直す
      return i + 1;
    }
  }
  return false;
}

function clearData() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var target1 = sheet.getRange("A6:K11");
  var target2 = sheet.getRange("B14:D14");
  var target3 = sheet.getRange("A16:K16");
  var target4 = sheet.getRange("C20:C20");
  var target5 = sheet.getRange("G20:P20");
  target1.clearContent();
  target2.clearContent();
  target3.clearContent();
  target4.clearContent();
  target5.clearContent();
}

function clearAll() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var target0 = sheet.getRange("B2:B3");
  var target1 = sheet.getRange("A6:K11");
  var target2 = sheet.getRange("B14:D14");
  var target3 = sheet.getRange("A16:K16");
  var target4 = sheet.getRange("C20:C20");
  var target5 = sheet.getRange("G20:P20");
  target0.clearContent();
  target1.clearContent();
  target2.clearContent();
  target3.clearContent();
  target4.clearContent();
  target5.clearContent();
}

// ShoppingList
function getCurrentPrices() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var tickers = sheet.getDataRange().getValues();
  for (var i = tickers.length - 2; i > 1; i--) {
    var ticker = tickers[i][0];
    var row = i + 2;
    sheet.getRange(row, 4).setFormula('=GOOGLEFINANCE(A' + row +', "price")');
  }
}

// BS
function writeBSItems() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var ticker = sheet.getRange(2, 2).getValue().toString();
  var stockExchange = getStockExchange(sheet);
  var byId = getById(ticker, stockExchange);
  var keyRatio;
  var keyRatioUrl = "https://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=jsonp1579473219364&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1579473220658";
  keyRatio = UrlFetchApp.fetch(keyRatioUrl).getContentText();
  
  getBSNestedList(sheet, keyRatio)
}

function getBSNestedList(sheet, keyRatio) {
  //assets
  var cashEtcList = getCashData(keyRatio);
  var accountsReceivableList = getAccountsReceivableData(keyRatio);
  var inventoryList = getInventoryData(keyRatio)
  var otherCurrentAssetList = getOtherCurrentAssetsData(keyRatio)
  var netPPAndEList = getNetPPAndEData(keyRatio)
  var intangiblesList = getIntangiblesData(keyRatio)
  var otherLongTermAssetsList = getOtherLongTermAssetsData(keyRatio)

  //Liabilities
  var accountsPayableList = getAccountsPayableData(keyRatio)
  var shortTermDebtList = getShortTermDebtData(keyRatio)
  var taxesPayableList = getTaxesPayableData(keyRatio)
  var accruedLiabilitiesList = getAccruedLiabilitiesData(keyRatio)
  var otherShortTermLiabilitiesList = getOtherShortTermLiabilitiesData(keyRatio)
  var longTermDebtList = getLongTermDebtData(keyRatio)
  var otherLongTermLiabilitiesList = getOtherLongTermLiabilitiesData(keyRatio)
  var totalStockholdersEquitysList = getTotalStockholdersEquitysData(keyRatio)

  var assetsItemName = [otherLongTermAssetsList[0], intangiblesList[0], netPPAndEList[0], otherCurrentAssetList[0], inventoryList[0], accountsReceivableList[0], cashEtcList[0]];
  var pastAssets = [otherLongTermAssetsList[1], intangiblesList[1], netPPAndEList[1], otherCurrentAssetList[1], inventoryList[1], accountsReceivableList[1], cashEtcList[1]];
  var latestAssets = [otherLongTermAssetsList[2], intangiblesList[2], netPPAndEList[2], otherCurrentAssetList[2], inventoryList[2], accountsReceivableList[2], cashEtcList[2]];

  var liabilitiesItemName =[totalStockholdersEquitysList[0],otherLongTermLiabilitiesList[0],longTermDebtList[0],otherShortTermLiabilitiesList[0],accruedLiabilitiesList[0],taxesPayableList[0],shortTermDebtList[0],accountsPayableList[0]];
  var pastLiabilities =[totalStockholdersEquitysList[1],otherLongTermLiabilitiesList[1],longTermDebtList[1],otherShortTermLiabilitiesList[1],accruedLiabilitiesList[1],taxesPayableList[1],shortTermDebtList[1],accountsPayableList[1]];
  var latestLiabilities =[totalStockholdersEquitysList[2],otherLongTermLiabilitiesList[2],longTermDebtList[2],otherShortTermLiabilitiesList[2],accruedLiabilitiesList[2],taxesPayableList[2],shortTermDebtList[2],accountsPayableList[2]];

  writeBSData(sheet, "past", assetsItemName,liabilitiesItemName,pastAssets,pastLiabilities);
  writeBSData(sheet, "latest", assetsItemName,liabilitiesItemName,latestAssets,latestLiabilities);

  var assetsColorScheme = ['#0041C2','#2554C7', '#1569C7', '#3090C7', '#659EC7', '#87AFC7', '#95B9C7'];
  generateGraph('資産', [15,2,0,0], 'left', 'D6:J7', assetsColorScheme);
  generateGraph('資産', [30,2,0,0], 'left', 'D10:J11', assetsColorScheme);

  var liabilitiesColorScheme = ['#728C00','#FFA62F','#FBB117','#E7A1B0','#F9A7B0','#FAAFBA','#FBBBB9','#FFDFDD']
  generateGraph('負債', [15,6,0,0], 'right', 'D8:K9', liabilitiesColorScheme);
  generateGraph('負債', [30,6,0,0], 'right', 'D12:K13', liabilitiesColorScheme);
}

function generateGraph(title, position, legendPosition, targetRange, colors) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange(targetRange)
  var chart = sheet.newChart()
    .addRange(range)
    .setPosition(position[0], position[1], position[2], position[3])
    .asColumnChart()
    .setNumHeaders(1)
    .setOption('title', title)
    .setOption('legend.position', legendPosition)
    .setOption('width', 400)
    .setOption('height', 300)
    .setOption('isStacked', 'percent' )
    .setOption('colors', colors)
    
  sheet.insertChart(chart.build());
}

function writeBSData(sheet, target, assetsItem, liabilitiesItem, assets, liabilities) {
  if (target === "past") {
    for (var i = 0; i < assets.length; i++) {
      sheet.getRange(6,i + 4).setValue(assetsItem[i]);
      sheet.getRange(7,i + 4).setValue(assets[i]);
    }
    for (var i = 0; i < liabilities.length; i++) {
      sheet.getRange(8,i + 4).setValue(liabilitiesItem[i]);
      sheet.getRange(9,i + 4).setValue(liabilities[i]);
    }
  } else {
    for (var i = 0; i < assets.length; i++) {
      sheet.getRange(10,i + 4).setValue(assetsItem[i]);
      sheet.getRange(11,i + 4).setValue(assets[i]);
    }
    for (var i = 0; i < liabilities.length; i++) {
      sheet.getRange(12,i + 4).setValue(liabilitiesItem[i]);
      sheet.getRange(13,i + 4).setValue(liabilities[i]);
    }
  }

}

function getCashData(keyRatio) {
  var cashRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i45\\">(.*?)</g;
  var cashTags = keyRatio.match(cashRegExp);
  var cashList = [];
  var itemName = "現金等";
  cashList.push(itemName);

  for(var i = 8; i < cashTags.length -1; i++) {
    cashList.push(replaceDashToZero(cashTags[i].match(/>(.*?)</)[1]));
  }
  return cashList;
}

function getAccountsReceivableData(keyRatio) {
  var accountsReceivableRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i46\\">(.*?)</g;
  var accountsReceivableTags = keyRatio.match(accountsReceivableRegExp);
  var accountsReceivableList = [];
  var itemName = "売掛金";
  accountsReceivableList.push(itemName);

  for(var i = 8; i < accountsReceivableTags.length -1; i++) {
    accountsReceivableList.push(replaceDashToZero(accountsReceivableTags[i].match(/>(.*?)</)[1]));
  }
  return accountsReceivableList;
}

function getInventoryData(keyRatio) {
  var inventoryRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i47\\">(.*?)</g;
  var inventoryTags = keyRatio.match(inventoryRegExp);
  var inventoryList = [];
  var itemName = "棚卸資産";
  inventoryList.push(itemName);

  for(var i = 8; i < inventoryTags.length -1; i++) {
    inventoryList.push(replaceDashToZero(inventoryTags[i].match(/>(.*?)</)[1]));
  }
  return inventoryList;
}

function getOtherCurrentAssetsData(keyRatio) {
  var otherCurrentAssetsRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i48\\">(.*?)</g;
  var otherCurrentAssetsTags = keyRatio.match(otherCurrentAssetsRegExp);
  var otherCurrentAssetsList = [];
  var itemName = "その他流動資産";
  otherCurrentAssetsList.push(itemName);

  for(var i = 9; i < otherCurrentAssetsTags.length; i++) {
    otherCurrentAssetsList.push(replaceDashToZero(otherCurrentAssetsTags[i].match(/>(.*?)</)[1]));
  }
  return otherCurrentAssetsList;
}

function getNetPPAndEData(keyRatio) {
  var otherCurrentAssetsRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i50\\">(.*?)</g;
  var otherCurrentAssetsTags = keyRatio.match(otherCurrentAssetsRegExp);
  var otherCurrentAssetsList = [];
  var itemName = "有形固定資産";
  otherCurrentAssetsList.push(itemName);

  for(var i = 8; i < otherCurrentAssetsTags.length -1; i++) {
    otherCurrentAssetsList.push(replaceDashToZero(otherCurrentAssetsTags[i].match(/>(.*?)</)[1]));
  }
  return otherCurrentAssetsList;
}

function getIntangiblesData(keyRatio) {
  var intangiblesRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i51\\">(.*?)</g;
  var intangiblesTags = keyRatio.match(intangiblesRegExp);
  var intangiblesList = [];
  var itemName = "無形資産";
  intangiblesList.push(itemName);

  for(var i = 8; i < intangiblesTags.length -1; i++) {
    intangiblesList.push(replaceDashToZero(intangiblesTags[i].match(/>(.*?)</)[1]));
  }
  return intangiblesList;
}

function getOtherLongTermAssetsData(keyRatio) {
  var otherLongTermAssetsRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i52\\">(.*?)</g;
  var otherLongTermAssetsTags = keyRatio.match(otherLongTermAssetsRegExp);
  var otherLongTermAssetsList = [];
  var itemName = "その他固定資産";
  otherLongTermAssetsList.push(itemName);

  for(var i = 8; i < otherLongTermAssetsTags.length -1; i++) {
    otherLongTermAssetsList.push(replaceDashToZero(otherLongTermAssetsTags[i].match(/>(.*?)</)[1]));
  }
  return otherLongTermAssetsList;
}

function getAccountsPayableData(keyRatio) {
  var accountsPayableRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i54\\">(.*?)</g;
  var accountsPayableTags = keyRatio.match(accountsPayableRegExp);
  var accountsPayableList = [];
  var itemName = "未払金";
  accountsPayableList.push(itemName);

  for(var i = 8; i < accountsPayableTags.length -1; i++) {
    accountsPayableList.push(replaceDashToZero(accountsPayableTags[i].match(/>(.*?)</)[1]));
  }
  return accountsPayableList;
}

function getShortTermDebtData(keyRatio) {
  var shortTermDebtRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i55\\">(.*?)</g;
  var shortTermDebtTags = keyRatio.match(shortTermDebtRegExp);
  var shortTermDebtList = [];
  var itemName = "短期借入金";
  shortTermDebtList.push(itemName);

  for(var i = 8; i < shortTermDebtTags.length -1; i++) {
    shortTermDebtList.push(replaceDashToZero(shortTermDebtTags[i].match(/>(.*?)</)[1]));
  }
  return shortTermDebtList;
}

function getTaxesPayableData(keyRatio) {
  var taxesPayableRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i56\\">(.*?)</g;
  var taxesPayableTags = keyRatio.match(taxesPayableRegExp);
  var taxesPayableList = [];
  var itemName = "短期借入金";
  taxesPayableList.push(itemName);

  for(var i = 8; i < taxesPayableTags.length -1; i++) {
    taxesPayableList.push(replaceDashToZero(taxesPayableTags[i].match(/>(.*?)</)[1]));
  }
  return taxesPayableList;
}

function getTaxesPayableData(keyRatio) {
  var taxesPayableRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i56\\">(.*?)</g;
  var taxesPayableTags = keyRatio.match(taxesPayableRegExp);
  var taxesPayableList = [];
  var itemName = "未払税金";
  taxesPayableList.push(itemName);

  for(var i = 8; i < taxesPayableTags.length -1; i++) {
    taxesPayableList.push(replaceDashToZero(taxesPayableTags[i].match(/>(.*?)</)[1]));
  }
  return taxesPayableList;
}

function getAccruedLiabilitiesData(keyRatio) {
  var accruedLiabilitiesRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i57\\">(.*?)</g;
  var accruedLiabilitiesTags = keyRatio.match(accruedLiabilitiesRegExp);
  var accruedLiabilitiesList = [];
  var itemName = "未払費用";
  accruedLiabilitiesList.push(itemName);

  for(var i = 8; i < accruedLiabilitiesTags.length -1; i++) {
    accruedLiabilitiesList.push(replaceDashToZero(accruedLiabilitiesTags[i].match(/>(.*?)</)[1]));
  }
  return accruedLiabilitiesList;
}

function getOtherShortTermLiabilitiesData(keyRatio) {
  var otherShortTermLiabilitiesRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i58\\">(.*?)</g;
  var otherShortTermLiabilitiesTags = keyRatio.match(otherShortTermLiabilitiesRegExp);
  var otherShortTermLiabilitiesList = [];
  var itemName = "その他流動負債";
  otherShortTermLiabilitiesList.push(itemName);

  for(var i = 8; i < otherShortTermLiabilitiesTags.length -1; i++) {
    otherShortTermLiabilitiesList.push(replaceDashToZero(otherShortTermLiabilitiesTags[i].match(/>(.*?)</)[1]));
  }
  return otherShortTermLiabilitiesList;
}

function getLongTermDebtData(keyRatio) {
  var longTermDebtRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i60\\">(.*?)</g;
  var longTermDebtTags = keyRatio.match(longTermDebtRegExp);
  var longTermDebtList = [];
  var itemName = "長期借入金";
  longTermDebtList.push(itemName);

  for(var i = 8; i < longTermDebtTags.length -1; i++) {
    longTermDebtList.push(replaceDashToZero(longTermDebtTags[i].match(/>(.*?)</)[1]));
  }
  return longTermDebtList;
}

function getOtherLongTermLiabilitiesData(keyRatio) {
  var otherLongTermLiabilitiesRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i61\\">(.*?)</g;
  var otherLongTermLiabilitiesTags = keyRatio.match(otherLongTermLiabilitiesRegExp);
  var otherLongTermLiabilitiesList = [];
  var itemName = "その他固定負債";
  otherLongTermLiabilitiesList.push(itemName);

  for(var i = 8; i < otherLongTermLiabilitiesTags.length -1; i++) {
    otherLongTermLiabilitiesList.push(replaceDashToZero(otherLongTermLiabilitiesTags[i].match(/>(.*?)</)[1]));
  }
  return otherLongTermLiabilitiesList;
}

function getTotalStockholdersEquitysData(keyRatio) {
  var totalStockholdersEquitysRegExp = /<td align=\\"right\\" headers=\\"fh-Y[0-9]{1,2} fh-balsheet i63\\">(.*?)</g;
  var totalStockholdersEquitysTags = keyRatio.match(totalStockholdersEquitysRegExp);
  var totalStockholdersEquitysList = [];
  var itemName = "資本金";
  totalStockholdersEquitysList.push(itemName);

  for(var i = 8; i < totalStockholdersEquitysTags.length -1; i++) {
    totalStockholdersEquitysList.push(replaceDashToZero(totalStockholdersEquitysTags[i].match(/>(.*?)</)[1]));
  }
  return totalStockholdersEquitysList;
}


// P/L
function writePLItems() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var ticker = sheet.getRange(2, 2).getValue().toString();
  var stockExchange = getStockExchange(sheet);
  var byId = getById(ticker, stockExchange);
  var keyRatio;
  var keyRatioUrl = "https://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=jsonp1579473219364&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1579473220658";
  keyRatio = UrlFetchApp.fetch(keyRatioUrl).getContentText();
  getPLNestedList(sheet, keyRatio)
}

function getPLNestedList(sheet, keyRatio) {
  //Profit, Cost
  var ebtMarginList = getEBTMargin(keyRatio);
  var otherCostList = getOtherCost(keyRatio);
  var rdList = getRD(keyRatio);
  var sgaList = getSGA(keyRatio)
  var cogsList = getCOGS(keyRatio)

  //Revenue
  var netIntIncAndOtherList = getNetIntIncAndOther(keyRatio);
  var revenue = calculateRevenue(netIntIncAndOtherList)

  var pcItemName = [ebtMarginList[0], otherCostList[0], rdList[0], sgaList[0], cogsList[0]];
  var pastPC = [ebtMarginList[1], otherCostList[1], rdList[1], sgaList[1], cogsList[1]];
  var latestPC = [ebtMarginList[2], otherCostList[2], rdList[2], sgaList[2], cogsList[2]];

  var salesItemName = [netIntIncAndOtherList[0],revenue[0]];
  var pastSales = [netIntIncAndOtherList[1],revenue[1]];
  var latestSales = [netIntIncAndOtherList[2],revenue[2]];

  writePLData(sheet, "past", pcItemName, salesItemName, pastPC, pastSales);
  writePLData(sheet, "latest", pcItemName,salesItemName,latestPC,latestSales);

  var pcColorScheme = ['#A1C935','#F88017', '#FF7F50', '#F88158', '#F9966B'];
  generateGraph('費用・利益', [15,2,0,0], 'left', 'D6:J7', pcColorScheme);
  generateGraph('費用・利益', [30,2,0,0], 'left', 'D10:J11', pcColorScheme);

  var salesColorScheme = ['#9CB071','#8BB381'];
  generateGraph('収益', [15,6,0,0], 'right', 'D8:K9', salesColorScheme);
  generateGraph('収益', [30,6,0,0], 'right', 'D12:K13', salesColorScheme);
}

function writePLData(sheet, target, pcItem, salesItem, pc, sales) {
  if (target === "past") {
    for (var i = 0; i < pc.length; i++) {
      sheet.getRange(6,i + 4).setValue(pcItem[i]);
      sheet.getRange(7,i + 4).setValue(pc[i]);
    }
    for (var i = 0; i < sales.length; i++) {
      sheet.getRange(8,i + 4).setValue(salesItem[i]);
      sheet.getRange(9,i + 4).setValue(sales[i]);
    }
  } else {
    for (var i = 0; i < pc.length; i++) {
      sheet.getRange(10,i + 4).setValue(pcItem[i]);
      sheet.getRange(11,i + 4).setValue(pc[i]);
    }
    for (var i = 0; i < sales.length; i++) {
      sheet.getRange(12,i + 4).setValue(salesItem[i]);
      sheet.getRange(13,i + 4).setValue(sales[i]);
    }
  }

}

function getEBTMargin(keyRatio) {
  var ebtMarginRegExp = /<td align=\\"right\\" headers=\\"pr-Y[0-9]{1,2} pr-margins i20\\">(.*?)</g;
  var ebtMarginTags = keyRatio.match(ebtMarginRegExp);
  var ebtMarginList = [];
  var itemName = "税引き前利益";
  ebtMarginList.push(itemName);

  for(var i = 8; i < ebtMarginTags.length -1; i++) {
    ebtMarginList.push(replaceDashToZero(ebtMarginTags[i].match(/>(.*?)</)[1]));
  }
  return ebtMarginList;
}

function getOtherCost(keyRatio) {
  var otherRegExp = /<td align=\\"right\\" headers=\\"pr-Y[0-9]{1,2} pr-margins i17\\">(.*?)</g;
  var otherTags = keyRatio.match(otherRegExp);
  var otherList = [];
  var itemName = "その他";
  otherList.push(itemName);

  for(var i = 8; i < otherTags.length -1; i++) {
    otherList.push(replaceDashToZero(otherTags[i].match(/>(.*?)</)[1]));
  }
  return otherList;
}

function getRD(keyRatio) {
  var rdRegExp = /<td align=\\"right\\" headers=\\"pr-Y[0-9]{1,2} pr-margins i16\\">(.*?)</g;
  var rdTags = keyRatio.match(rdRegExp);
  var rdList = [];
  var itemName = "研究開発費";
  rdList.push(itemName);

  for(var i = 8; i < rdTags.length -1; i++) {
    rdList.push(replaceDashToZero(rdTags[i].match(/>(.*?)</)[1]));
  }
  return rdList;
}

function getRD(keyRatio) {
  var rdRegExp = /<td align=\\"right\\" headers=\\"pr-Y[0-9]{1,2} pr-margins i16\\">(.*?)</g;
  var rdTags = keyRatio.match(rdRegExp);
  var rdList = [];
  var itemName = "研究開発費";
  rdList.push(itemName);

  for(var i = 8; i < rdTags.length -1; i++) {
    rdList.push(replaceDashToZero(rdTags[i].match(/>(.*?)</)[1]));
  }
  return rdList;
}

function getSGA(keyRatio) {
  var sgaRegExp = /<td align=\\"right\\" headers=\\"pr-Y[0-9]{1,2} pr-margins i15\\">(.*?)</g;
  var sgaTags = keyRatio.match(sgaRegExp);
  var sgaList = [];
  var itemName = "販管費";
  sgaList.push(itemName);

  for(var i = 8; i < sgaTags.length -1; i++) {
    sgaList.push(replaceDashToZero(sgaTags[i].match(/>(.*?)</)[1]));
  }
  return sgaList;
}

function getCOGS(keyRatio) {
  var cogsRegExp = /<td align=\\"right\\" headers=\\"pr-Y[0-9]{1,2} pr-margins i13\\">(.*?)</g;
  var cogsTags = keyRatio.match(cogsRegExp);
  var cogsList = [];
  var itemName = "売上原価";
  cogsList.push(itemName);

  for(var i = 8; i < cogsTags.length -1; i++) {
    cogsList.push(replaceDashToZero(cogsTags[i].match(/>(.*?)</)[1]));
  }
  return cogsList;
}

function getNetIntIncAndOther(keyRatio) {
  var netIntIncAndOtherRegExp = /<td align=\\"right\\" headers=\\"pr-Y[0-9]{1,2} pr-margins i19\\">(.*?)</g;
  var netIntIncAndOtherTags = keyRatio.match(netIntIncAndOtherRegExp);
  var netIntIncAndOtherList = [];
  var itemName = "営業外収益";
  netIntIncAndOtherList.push(itemName);

  for(var i = 8; i < netIntIncAndOtherTags.length -1; i++) {
    netIntIncAndOtherList.push(replaceDashToZero(netIntIncAndOtherTags[i].match(/>(.*?)</)[1]));
  }
  return netIntIncAndOtherList;
}

function calculateRevenue(netIntIncAndOtherList) {
  return ["売上高", 100 - netIntIncAndOtherList[1], 100 - netIntIncAndOtherList[2]]
}

function clearBSAllCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('貸借対照表');

  // This removes all the embedded charts from the spreadsheet
  var charts = sheet.getCharts();
  for (var i in charts) {
    sheet.removeChart(charts[i]);
  }
}

function clearPLAllCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('損益計算書');

  // This removes all the embedded charts from the spreadsheet
  var charts = sheet.getCharts();
  for (var i in charts) {
    sheet.removeChart(charts[i]);
  }
}


