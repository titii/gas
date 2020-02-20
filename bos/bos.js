var no = "なし　No";
var yes = "ある　Yes";
function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var ticker = sheet.getRange(2, 2).getValue();;
  var byId = getById(ticker);
  // var stockMarket = getStockMarcket(ticker);

  var financialsUrl = "https://financials.morningstar.com/finan/financials/getFinancePart.html?&callback=jsonp1580192904311&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1580192905566"
  var financials = UrlFetchApp.fetch(financialsUrl).getContentText();
  var keyRatioUrl = "https://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=jsonp1579473219364&t="+ byId +"&region=usa&culture=en-US&cur=&order=asc&_=1579473220658";
  var keyRatio = UrlFetchApp.fetch(keyRatioUrl).getContentText();

  var epsList = getEPS(financials);
  var freeCashFlowList = getFreeCashFlow(financials);
  var dividendsList = getDividends(financials);
  var roeList = getROE(keyRatio);
  var interestCoverageList = getInterestCoverage(keyRatio);
  var netProfitMarginList = getNetProfitMargin(keyRatio);

  clearFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList);
  writeFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList);
  updateAssessSheet(epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList);
}

function getStockMarcket(ticker) {
  var url = "https://www.google.com/search?q=morningstar+" + ticker.toLowerCase() + "&num=" + 10;
  var response = UrlFetchApp.fetch(url).getContentText('UTF-8');
  var morningstarRegExp = /<cite class="iUh30 bc rpCHfe tjvcx">www.morningstar.com › stocks › .+ › quote</;
  var citeTag = response.match(morningstarRegExp)
  Logger.log(citeTag)
}

function clearFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList) {
  var fundamentals = [epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList];
  var rows = fundamentals.length;
  // sheet分を-1
  var cols = fundamentals[0].length;

  sheet.getRange(4,1,rows,cols).clearContent();
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

function getById(ticker) {
  var mStarQuoteUrl = "https://www.morningstar.com/stocks/xnys/"+ ticker.toLowerCase() +"/quote";
  var mStarQuote = UrlFetchApp.fetch(mStarQuoteUrl).getContentText();
  var byIdRegExp = /byId:{\"(.+)\":u}/;
  var byId = mStarQuote.match(byIdRegExp)
  return byId[1];
}

function writeFundamentals(sheet, epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList) {
  var fundamentals = [epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList];
  var rows = fundamentals.length;
  // sheet分を-1
  var cols = fundamentals[0].length;

  sheet.getRange(4,1,rows,cols).setValues(fundamentals);
}

function updateAssessSheet(epsList, freeCashFlowList, dividendsList, roeList, interestCoverageList, netProfitMarginList) {
	assessEPS(epsList);
	assessFreeCashFlow(freeCashFlowList);
	assessDividends(dividendsList);
	assessROE(roeList);
	assessIC(interestCoverageList);
}

function assessEPS(epsList) {
  for(var i = 0; i < epsList.length; i++) {
	 	if(epsList[i] < 0) {
	      return no;
	    }
	}
	return yes;
}

function assessFreeCashFlow(freeCashFlowList) {
  for(var i = 0; i < freeCashFlowList.length; i++) {
	 	if(freeCashFlowList[i] < 0) {
	      return no;
	    }
	}
	return yes;
}

function assessDividends(dividendsList) {
	var dividends = 0;
  for(var i = 0; i < dividendsList.length; i++) {
	 	if(dividendsList[i] < 0) {
	      dividends = 0;
	    } else {
	    	dividends = dividendsList[i];
	    }
	}
	if (dividends < 0) {
		return no;
	} else {
		return yes;
	}
}

function assessROE(roeList) {
  for(var i = 0; i < roeList.length; i++) {
	 	if(roeList[i] < 15) {
	      return no;
	    }
	}
	return yes;
}

function assessIC(interestCoverageList) {
	var ic = ">  10";
  for(var i = 0; i < interestCoverageList.length; i++) {
	 	if(interestCoverageList[i] > 10) {
	    ic = ">  10";
	  }
	  else if(interestCoverageList[i] > 4) {
	    ic = ">  4";
	  }
	  else {
	  	ic = no;
	  }

	}
	return no;
}

function getEPS(financials) {
  var epsRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i5\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i5\\">(.*?)<\\\/th>/g;
  var epsTags = financials.match(epsRegExp);
  var epsList = [];
  var htmlTag = financials.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  epsList.push(colName);


  for(var i = 0; i < epsTags.length; i++) {
    epsList.push(epsTags[i].match(/>(.*?)</)[1]);
  }
  return epsList;
}

function getFreeCashFlow(financials) {
  var fcfRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i11\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i11\\">(.*?)<\\\/th>/g;
  var fcfTags = financials.match(fcfRegExp);
  var fcfList = []
  var htmlTag = financials.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  fcfList.push(colName);
  for(var i = 0; i < fcfTags.length; i++) {
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
  for(var i = 0; i < dividendsTags.length; i++) {
    dividendsList.push(dividendsTags[i].match(/>(.*?)</)[1]);
  }
  return dividendsList;
}

function getROE(keyRatio) {
  var roeRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i26\\">(.*?)</g;
  var colNameRegExp = /<th class=\\"row_lbl\\" scope=\\"row\\" id=\\"i26\\">(.*?)<\\\/th>/g;
  var roeTags = keyRatio.match(roeRegExp);
  var roeList = []
  var htmlTag = keyRatio.match(colNameRegExp)[0];
  var colName = extractColName(htmlTag);
  roeList.push(colName);
  for(var i = 0; i < roeTags.length; i++) {
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
  for(var i = 0; i < icTags.length; i++) {
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
  for(var i = 0; i < npmTags.length; i++) {
    npmList.push(npmTags[i].match(/>(.*?)</)[1]);
  }
  return npmList;
}


