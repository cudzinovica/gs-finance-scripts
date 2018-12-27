var API_KEY="PUT_API_KEY_HERE";

/**
 * appends most recent closing prices of stock tickers specified in stock tickers sheet to data sheet
 */
function addClosingPrices() {
  var spreadsheet = SpreadsheetApp.getActive();
  var stockTickerSheet = spreadsheet.getSheetByName('stock tickers');
  stockTickerSheet.activate();

  // Get stock data
  var numStocks = numStocks_(stockTickerSheet);
  var stockValues = stockTickerSheet.getRange(2,1,numStocks).getValues();
  
  // Insert stock data as a new row in data sheet
  var dataSheet = spreadsheet.getSheetByName('data');
  dataSheet.activate();
  
  var today = new Date();
  var date = today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate();
  
  var header = [['Date']];
  var newRow = [date];
  
  // for each stock ticker attempts to get closing data up to 5 times
  for(var i = 0; i < numStocks; i++){
    header[0].push(stockValues[i][0]);
    var closingPrice = 0;
    var tries = 0;
    var success = false;
    while(!success){
      try {
        closingPrice = getClosingPriceAlphaVantage_(stockValues[i][0], 1);
        success = true;
      } catch (e) {
        Logger.log(e);
        tries++;
        Utilities.sleep(5000);
        if (tries >= 5){
          success = true;
        }
      }
    }
    newRow.push(closingPrice);
  }
  
  dataSheet.getRange(1,1,1,header[0].length).setValues(header);
  dataSheet.appendRow(newRow);
}

/**
 * returns number of stock ticker symbols
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
 * Returns closing price data for given stock symbol using alpha vantage api
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of given stock symbol
 */
function getClosingPriceAlphaVantage_(stockSymbol) {
  
  var closingPrice;
  
  var url = "https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=" + stockSymbol + "&apikey=" + API_KEY;
  
  Logger.log(url);
  
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
 * Returns closing price data for given stock symbol using iex api
 * @param stockSymbol: the stock symbol we want closing price of
 * @return: the closing price of given stock symbol
 */
function getClosingPriceIex_(stockSymbol) {
  
  var closingPrice;
  
  var apiUrl = "https://api.iextrading.com/1.0";
  
  var url = apiUrl + "/stock/" + stockSymbol + "/ohlc";
  
  Logger.log(url);
  
  var response = UrlFetchApp.fetch(url);
  
  if (response.getResponseCode() == 200) {
    
    var resultsAsString = response.getContentText();
    
    var results = JSON.parse(resultsAsString);
    
    var close = results["close"];
    
    if ( !close ) {
      Logger.log('failed to find data for: ' + stockSymbol)
      Logger.log('results:');
      Logger.log(results);
      throw new Error('no data in results!');
    }
    
    var closingPrice = Number(close["price"]).toFixed(2);
    
  } else {
    Logger.log("There was an error downloading today's financial data");
  }
  
  return closingPrice;
  
}
    
