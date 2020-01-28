function myFunction() {
  var ticker = "GIS";
  var url = "http://financials.morningstar.com/finan/financials/getKeyStatPart.html?&callback=jsonp1579473219364&t=XNYS:"+ ticker +"&region=usa&culture=en-US&cur=&order=asc&_=1579473220658";
  var jsonp = UrlFetchApp.fetch(url).getContentText();
  var roeList = getROE(jsonp);
  
  
  Logger.log(roeList);
}

function getROE(jsonp) {
  var roeRegExp = /<td align=\\"right\\" headers=\\"pr-pro-Y[0-9]{1,2} pr-profit i26\\">(.*?)</g;
  var roeTags = jsonp.match(roeRegExp);
  var roeList = []
  for(var i = 0; i < roeTags.length; i++) {
    roeList.push(roeTags[i].match(/>(.*?)</)[1]);
  }
  return roeList;
}