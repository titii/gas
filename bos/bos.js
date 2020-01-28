function myFunction() {
  var ticker = "GIS";
  var financialsUrl = "http://financials.morningstar.com/finan/financials/getFinancePart.html?&callback=jsonp1580192904311&t=XNYS:"+ ticker +"&region=usa&culture=en-US&cur=&order=asc&_=1580192905566"
  var financials = UrlFetchApp.fetch(financialsUrl).getContentText();
  var keyRatioUrl = "http://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=jsonp1579473219364&t=XNYS:"+ ticker +"&region=usa&culture=en-US&cur=&order=asc&_=1579473220658";
  var keyRatio = UrlFetchApp.fetch(keyRatioUrl).getContentText();

  var epsList = getEPS(financials);
  var freeCashFlowList = getFreeCashFlow(financials);
  var dividends = getDividends(financials);
  var roeList = getROE(keyRatio);
  var interestCoverageList = getInterestCoverage(keyRatio);
  var netProfitMarginList = getNetProfitMargin(keyRatio);
  
  
  Logger.log(roeList);
}

function getEPS(financials) {
  var epsRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i0\\">(.*?)</g;
  var epsTags = financials.match(epsRegExp);
  var epsList = []
  for(var i = 0; i < epsTags.length; i++) {
    epsList.push(epsTags[i].match(/>(.*?)</)[1]);
  }
  return epsList;
}

function getFreeCashFlow(financials) {
  var fcfRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i11\\">(.*?)</g;
  var fcfTags = financials.match(fcfRegExp);
  var fcfList = []
  for(var i = 0; i < fcfTags.length; i++) {
    fcfList.push(fcfTags[i].match(/>(.*?)</)[1]);
  }
  return fcfList;
}

function getDividends(financials) {
  var dividendsRegExp = /<td align=\\"right\\" headers=\\"Y[0-9]{1,2} i6\\">(.*?)</g;
  var dividendsTags = financials.match(dividendsRegExp);
  var dividendsList = []
  for(var i = 0; i < dividendsTags.length; i++) {
    dividendsList.push(dividendsTags[i].match(/>(.*?)</)[1]);
  }
  return dividendsList;
}

function getROE(keyRatio) {
  var roeRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i26\\">(.*?)</g;
  var roeTags = keyRatio.match(roeRegExp);
  var roeList = []
  for(var i = 0; i < roeTags.length; i++) {
    roeList.push(roeTags[i].match(/>(.*?)</)[1]);
  }
  return roeList;
}

function getInterestCoverage(keyRatio) {
  var icRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i95\\">(.*?)</g;
  var icTags = keyRatio.match(icRegExp);
  var icList = []
  for(var i = 0; i < icTags.length; i++) {
    icList.push(icTags[i].match(/>(.*?)</)[1]);
  }
  return icList;
}

function getNetProfitMargin(keyRatio) {
  var npmRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i22\\">(.*?)</g;
  var npmTags = keyRatio.match(icRegExp);
  var npmList = []
  for(var i = 0; i < npmTags.length; i++) {
    npmList.push(npmTags[i].match(/>(.*?)</)[1]);
  }
  return npmList;
}

