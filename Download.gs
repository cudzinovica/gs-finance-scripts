/**
 * Copies yahoo stock closing price to data sheet
 */

var API_KEY="QJL37TBV56L03I7U";
 
function copyYahooData() {
  var spreadsheet = SpreadsheetApp.getActive();
  var yahooSheet = spreadsheet.getSheetByName('yahoo links');
  yahooSheet.activate();

  // Get stock data
  var numStocks = numStocks_(yahooSheet);
  var stockValues = yahooSheet.getRange(2,1,numStocks).getValues();
  
  // Insert stock data as a new row in data sheet
  var dataSheet = spreadsheet.getSheetByName('data');
  dataSheet.activate();
  
  var today = new Date();
  //today.setDate(today.getDate() - 1);
  var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
  
  
  var header = [['Date']];
  var newRow = [date];
  
  try {
    for(var i = 0; i < numStocks; i++){
      header[0].push(stockValues[i][0]);
      newRow.push(getFinanceData_(stockValues[i][0], 1));
    }
  } catch (e) {
    Logger.log(e);
    return;
  }
  
  dataSheet.getRange(1,1,1,header[0].length).setValues(header);
  
  dataSheet.appendRow(newRow);
}

/**
 * returns number of stocks
 */
function numStocks_(sheet){
  var range = sheet.getRange(2,1,100);
  var values = range.getValues();
  
  var numStocks = 0;
  while(numStocks < values.length){
    if(values[numStocks][0]=="")
      break;
    
    numStocks++;
  }
  
  return numStocks;
}

/**
 * Returns closing price data for given stock symbol using query1.finance.yahoo.com
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of supplied stock symbol
 */
function getFinanceData_(stockSymbol) {
  
  var closingPrice;
  
  var url = "https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=" + stockSymbol + "&apikey=" + API_KEY;
  
  var response = UrlFetchApp.fetch(url);
  
  if (response.getResponseCode() == 200) {
    
    var resultsAsString = response.getContentText();
    
    var results = JSON.parse(resultsAsString);
    
    var timeSeriesDaily = results["Time Series (Daily)"];
    
    if ( !timeSeriesDaily ) {
      Logger.log('failed to find data for: ' + stockSymbol)
      Logger.log('results:');
      Logger.log(results);
      throw new Error('no data in results!');
    }
    
    var key = Object.keys(timeSeriesDaily)[0];
    
    var data = timeSeriesDaily[key];
    
    var keyName = "4. close";
    
    var closingPrice = Number(data[keyName]).toFixed(2);
  } else {
    Logger.log("There was an error downloading today's financial data");
  }
  
  return closingPrice;
  
}

/**
 * Returns closing price data for given stock symbol using finance.yahoo.com
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of supplied stock symbol
 */
function getYahooFinanceData_(stockSymbol) {
  
  var closingPrice;
  
  var url = "http://finance.yahoo.com/d/quotes.csv?s=" + stockSymbol + "&f=p";
  
  
  var response = UrlFetchApp.fetch(url.replace("^","%5E"), {muteHttpExceptions: true});
  
  if (response.getResponseCode()) {
    
    var textFile = response.getContentText();
    
    // If the URL is incorrect, Yahoo will return a 404 html page and not a CSV
    if (textFile.indexOf("<html>") == -1) {
      
      closingPrice = textFile;
      
    }
    
  }
  
  return closingPrice;
  
}

/**
 * Returns closing price data for given stock symbol using query1.finance.yahoo.com
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of supplied stock symbol
 */
function getYahooFinanceData2_(stockSymbol) {
  
  var closingPrice;
  
  var today = new Date();
  var yesterday = Math.floor(today/1000)-100000;
  
  var url = "https://query1.finance.yahoo.com/v7/finance/download/" + stockSymbol + "?period1=" + yesterday + 
      "&period2=" + Math.floor(today/1000) + "&interval=1d&events=history";
  
  
  var response = UrlFetchApp.fetch(url.replace("^","%5E"), {muteHttpExceptions: true, 
      method: "post", payload: "{username:manyenc,password:3Cudzinovi*}"});
  
  if (response.getResponseCode()) {
    
    var textFile = response.getContentText();
    
    // If the URL is incorrect, Yahoo will return a 404 html page and not a CSV
    if (textFile.indexOf("error") == -1) {
      
      var line = textFile.split("\n")[1];
      closingPrice = Number(line.split(",")[4]).toFixed(2);
      
    } else {
      Browser.msgBox("There was an error downloading today's yahoo data");
    }
    
  }
  
  return closingPrice;
  
}