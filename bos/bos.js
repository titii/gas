function myfunction() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var sht = ss.getSheetByName("BOS");
	var targetCell = sht.getActiveCell();
	var Ticker = targetCell.getValue();
	var html = UrlFetchApp.fetch('https://www.morningstar.com/stocks/xtks/' + Ticker + '/quote').getContentText();

	var parser = Parser.data(html);
	var trList = parser.from('<tr>').to('</tr>').iterate();
	for each(var tr in trList) {
 		var category = Parser.data(tr).from('<div class="category">').to('</div>').build();
		if (category.indexOf("技術") === -1) {
			continue;
  		}
  		var title = Parser.data(entry).from('<div class="title">').to('</div>').build();
  		titles.push(title);
	}
}