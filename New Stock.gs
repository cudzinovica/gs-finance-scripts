var API_KEY="PUT_API_KEY_HERE";

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
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
 * Adds new stock to stock ticker sheet and adds columns to financial sheet
 */
function addNewStock_() {
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
  
  //add name to stock tickers sheet
  var stocksSheet = spreadsheet.getSheetByName('stock tickers');
  stocksSheet.activate();
  
  stocksSheet.appendRow([stockTicker]);
  
  //add new columns to financial sheet
  addFinCols_()
  
}

/**
 * adds columns to financial sheet
 */
function addFinCols_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var stocksSheet = spreadsheet.getSheetByName('stock tickers');
  
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
 * Creates a new historical data sheet for given stock ticker
 */
function addHistSheet_(stockTicker,interval){
  var spreadsheet = SpreadsheetApp.getActive();
  
  spreadsheet.insertSheet(stockTicker + " " + interval,spreadsheet.getSheets().length);
  var tickerSheet = spreadsheet.getSheetByName(stockTicker+" "+interval);
  tickerSheet.activate();
  
  // tries to get historical data up to 5 times
  var closingPrices;
  var tries = 0;
  var success = false;
  while(!success){
    try {
      closingPrices = getFinanceDataHist_(stockTicker);
      success = true;
    } catch (e) {
      Logger.log(e);
      tries++;
      if (tries >= 5){
        success = true;
      }
    }
  }
  if(closingPrices){
    var range = tickerSheet.getRange(1,1,closingPrices.length,closingPrices[0].length);
    range.setValues(closingPrices);
  }
    
}


/**
 * Returns historical closing price data for given stock symbol using alpha vantageapi . Returns values no older than 2007-10-01
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of supplied stock symbol
 */
function getFinanceDataHist_(stockSymbol) {
  
  var closingPrices = [];
  
  var url = "https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=" + stockSymbol + "&outputsize=full&apikey=" + API_KEY;
  
  Logger.log(url);
  
  var response = UrlFetchApp.fetch(url);
  
  if (response.getResponseCode() == 200) {
    
    var resultsAsString = response.getContentText();
    
    var results = JSON.parse(resultsAsString);
    
    var timeSeriesDaily = results["Time Series (Daily)"];
    
    if(!timeSeriesDaily){
      Logger.log('failed to find data for: ' + stockSymbol)
      Logger.log('results:');
      Logger.log(results);
      throw new Error('no data in results!');
    }
    
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
