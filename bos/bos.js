var no = "なし　No";
var yes = "ある　Yes";
var intCovRatioValues = [">  10", "> 4"];
var netProfitMarginValues = [">20%", ">10%"];
var searchError = "エラーです。次のことを順に確認してください。\\n TickerやExchangesを変更した後フォーカスを外しましたか？ \\n TickerやExchangesは正確ですが？ \\n モーニングスターにデータはありますか？"


function myFunction() {
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
  Logger.log(growthPercentile)
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
  for(var i = 0; i < epsList.length; i++) {
    if(replaceToFloat(epsList[i]) < 0) {
        return no;
      }
  }
  return yes;
}

function assessFreeCashFlow(freeCashFlowList) {
  for(var i = 0; i < freeCashFlowList.length; i++) {
    if(replaceToNumber(freeCashFlowList[i]) < 0) {
        return no;
      }
  }
  return yes;
}

function assessDividends(dividendsList) {
  var dividends = 0;
  for(var i = 0; i < dividendsList.length; i++) {
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
  for(var i = 0; i < roeList.length; i++) {
    if(replaceToFloat(roeList[i]) < 15) {
        return no;
      }
  }
  return yes;
}

function assessIC(interestCoverageList) {
  var ic = ">  10";
  for(var i = 0; i < interestCoverageList.length; i++) {
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
  for(var i = 0; i < netProfitMarginList.length; i++) {
    if(replaceToFloat(netProfitMarginList[i]) > 20) {
      npm = ">20%";
    }
    else if(netProfitMarginList[i] > 10) {
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