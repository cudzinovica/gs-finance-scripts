/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
 
var API_KEY="QJL37TBV56L03I7U";

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Add new stock...', functionName: 'addNewStock_'},
    {name: 'Add new historical data sheet...', functionName: 'addNewHistSheet_'}
  ];
  spreadsheet.addMenu('Stocks', menuItems);
}

/**
 * SET THE HISTORICAL DATA INTERVAL HERE
 */
function getInterval_(){
  return "1d";
}

/**
 * Adds new stock to yahoo links, generates historical stock data, 
 * and adds columns to financial sheet
 */
function addNewStock_(){
  var interval = getInterval_();
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  //prompt user for stock name
  var selectedStock = Browser.inputBox('Add new stock',
      'Please enter the stock symbol that you want to add' +
      ' (for example, "TLT"):',
      Browser.Buttons.OK_CANCEL);
  if (selectedStock == 'cancel') {
    return;
  }
  var stockTicker = String(selectedStock);
  
  //add name to yahoo links
  var stocksSheet = spreadsheet.getSheetByName('yahoo links');
  stocksSheet.activate();
  
  stocksSheet.appendRow([stockTicker]);
  
  //add historical data to data sheet
  copyYahooData();
  
  addHistSheet_(stockTicker, interval);
  
  //add new columns to financial sheet
  var workSheet = spreadsheet.getSheets()[2];
  workSheet.activate();
  
  //get num stocks for index
  var numStocks = stocksSheet.getLastRow()-1;
  
  //add ? and % cols
  var numInserted = 0;
  var col = numStocks + 1 + numInserted;
  workSheet.insertColumnBefore(col);
  numInserted++;
  var cell = workSheet.getRange(1,col);
  cell.setValue("=IFERROR(indirect(\"data!" + String.fromCharCode(97+numStocks) + "1\")&\"?\")");
  
  col = numStocks*2 + 1 + numInserted;
  workSheet.insertColumnBefore(col);
  numInserted++;
  cell = workSheet.getRange(1,col);
  cell.setValue("=IFERROR(indirect(\"data!" + String.fromCharCode(97+numStocks) + "1\")&\" %\")");
  
  //update cash % col
  col++;
  cell = workSheet.getRange(2,col);
  cell.setValue("=100-SUM(" + String.fromCharCode(97+col-2-numStocks+1) + "2:" + String.fromCharCode(97+col-2) + "2)");
  cell.copyTo(workSheet.getRange(3,col,100));
  
  //add data cols
  var padding = (numStocks+1)*2+4+(numStocks-1)*6-1;
  workSheet.insertColumnsAfter(padding,6);
  
  //format black separator col
  range = workSheet.getRange(getA1_(padding+1)+":"+getA1_(padding+1));
  range.setBackground("black");
  workSheet.setColumnWidth(padding+1,10);
  
  //format 5 col data range
  range = workSheet.getRange(getA1_(padding+2)+":"+getA1_(padding+6));
  range.setBorder(true,true,true,true,true,true);
  range = workSheet.getRange(getA1_(padding+3)+":"+getA1_(padding+5));
  range.setFontWeight("bold");
  
  //update headers of each col
  var firstCol = padding+2;
  var values = [
    ["=IFERROR(indirect(\"data!" + String.fromCharCode(97+numStocks) + "1\"))",
     "new shares", "total shares", "S price", "market $"]
  ];
  workSheet.getRange(1,firstCol,1,5).setValues(values);
  
  //update first data row for each col
  var values = [
    [0,"="+getA1Not_(firstCol,2,2),
    "=iferror("+ getA1Not_(firstCol,2,4) +"/" + getA1Not_(firstCol,2,3) + ")",
    "="+getA1Not_(firstCol,3,3), "="+getA1Not_(firstCol,2,0)]
  ];
  workSheet.getRange(2,firstCol,1,5).setValues(values);
  
  //update formulas for each col
  var values = [
    ["="+getA1Not_(firstCol,3,1)+"*"+getA1Not_(firstCol,3,3),"="+getA1Not_(firstCol,3,2)+"-"+getA1Not_(firstCol,2,2),
    "=iferror("+ getA1Not_(firstCol,3,4) +"/" + getA1Not_(firstCol,3,3) + ")",
    "=IFERROR(indirect(\"data!"+String.fromCharCode(97+numStocks)+"\"&match($A3,data!$A$1:$A$1739)))",
    "=if("+String.fromCharCode(97+numStocks)+"3=\"n\",0,IFERROR("+getA1_(numStocks*2+2)+"3/sumif($B3:$F3,\"y\",$H3:$L3)*$O3))"]
  ];
  range = workSheet.getRange(3,firstCol,1,5);
  range.setValues(values);
  range.copyTo(workSheet.getRange(4,firstCol,100,5));
  
  //update total money col
  var totalMoneyCell = workSheet.getRange(2,(numStocks+1)*2+3);
  totalMoneyCell.setFormula(totalMoneyCell.getFormula()+"+"+getA1Not_(firstCol,2,4));
  var totalMoneyCell = workSheet.getRange(3,(numStocks+1)*2+3);
  totalMoneyCell.setFormula(totalMoneyCell.getFormula()+"+"+getA1Not_(firstCol,2,2)+"*"+getA1Not_(firstCol,3,3));
  totalMoneyCell.copyTo(workSheet.getRange(4,(numStocks+1)*2+3,100));
  
}

/**
 * returns a1 notation relative to 5 col range
 */
function getA1Not_(firstCol, row, col){
  var workSheet = SpreadsheetApp.getActiveSheet();
  
  return workSheet.getRange(row,firstCol+col).getA1Notation()
}

/**
 * Returns a1 notation for integer
 */
function getA1_(number){
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var len = 1;
  if(number>26*26)
    len = 3;
  else if(number>26)
    len = 2;
  
  var a1 = sheet.getRange(1,number).getA1Notation();
  
  return a1.substr(0,len);
}

/**
 * Adds new historical data sheet for specified stock symbol
 */
function addNewHistSheet_(){
  var interval = getInterval_();
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  //prompt user for stock name
  var selectedStock = Browser.inputBox('Add new historical data sheet',
      'Please enter the stock symbol for which you want to add a historical data sheet' +
      ' (for example, "TLT"):',
      Browser.Buttons.OK_CANCEL);
  if (selectedStock == 'cancel') {
    return;
  }
  var stockTicker = String(selectedStock);
  
  addHistSheet_(stockTicker, interval);
}

/**
 * Creates a new historical data sheet
 */
function addHistSheet_(stockTicker,interval){
  var spreadsheet = SpreadsheetApp.getActive();
  
  spreadsheet.insertSheet(stockTicker + " " + interval,spreadsheet.getSheets().length);
  var tickerSheet = spreadsheet.getSheetByName(stockTicker+" "+interval);
  tickerSheet.activate();
  var closingPrices = getFinanceDataHist_(stockTicker);
  var range = tickerSheet.getRange(1,1,closingPrices.length,closingPrices[0].length);
  range.setValues(closingPrices);
}


/**
 * Returns closing price data for given stock symbol using alphavantage. Returns values no older than 2007-10-01
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of supplied stock symbol
 */
function getFinanceDataHist_(stockSymbol) {
  
  var closingPrices = [];
  
  var url = "https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=" + stockSymbol + "&outputsize=full&apikey=" + API_KEY;
  
  var response = UrlFetchApp.fetch(url);
  
  if (response.getResponseCode() == 200) {
    
    var resultsAsString = response.getContentText();
    
    var results = JSON.parse(resultsAsString);
    
    var timeSeriesDaily = results["Time Series (Daily)"];
    
    var MAX_ROWS = 2000000
    var counter = MAX_ROWS - 1000;
    
    var endDate = new Date('2007-10-01');
    for (key in timeSeriesDaily){
      counter--;
      if(counter < 0) break;
      
      var currDate = new Date(key)
      if (currDate < endDate) break;
      
      var data = timeSeriesDaily[key];
      
      var keyName = "4. close";
      
      var closingPrice = data[keyName];
      closingPrices.push([key, Number(closingPrice).toFixed(2)]);
    }
    
    Logger.log("time series daily:");
    Logger.log(timeSeriesDaily);
    
  } else {
    Logger.log("error getting finance history data");
  }
  
  return closingPrices;
  
}

/**
 * Returns closing price data for given stock symbol using query1.finance.yahoo.com
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of supplied stock symbol
 */
function getYahooFinanceDataHist_(stockSymbol,interval) {
  
  var closingPrices = [];
  
  var today = new Date();
  var begin = new Date();
  begin.setFullYear(2000,0,1);
  
  var url = "https://query1.finance.yahoo.com/v7/finance/download/" + stockSymbol + "?period1="
    + Math.floor(begin/1000) + "&period2=" + Math.floor(today/1000) + "&interval=" + interval + "&events=history";
  
  
  var response = UrlFetchApp.fetch(url.replace("^","%5E"), {muteHttpExceptions: true, 
      method: "post", payload: "{username:manyenc,password:3Cudzinovi*}"});
  
  if (response.getResponseCode()) {
    var textFile = response.getContentText();
    
    
    // If the URL is incorrect, Yahoo will return a 404 html page and not a CSV
    if (textFile.indexOf("<html>") == -1) {
      var data = textFile.split("\n");
      data.splice(0,1);
      for(var i = 0; i < data.length; i++){
        var dataLine = data[i].split(",");
        var currPrice = Number(dataLine[4]).toFixed(2);
        if(currPrice == 0)
          currPrice = null;
        closingPrices.push([dataLine[0],currPrice]);
      }
    }
  }
  
  return closingPrices;
  
}






