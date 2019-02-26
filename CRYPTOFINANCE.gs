/**
 * CODE LICENSED UNDER THE CREATIVE COMMON BY-NC-ND LICENSE.
 * https://creativecommons.org/licenses/by-nc-nd/4.0/
 * 
 * Copyright 2019 by cryptofinance.ai
 */


/**
 * @OnlyCurrentDoc
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CRYPTOFINANCE')
      .addItem('How to refresh rates', 'ShowRefreshInfo')
      .addSeparator()
      .addItem('Set Data Availability plan API Key', 'ShowAPIKeyDataAvaibilityPrompt')
      .addSeparator()
      .addItem('Set Historical Data plan API Key', 'ShowAPIKeyHistPlanPrompt')
      .addSeparator()
      .addItem('Documentation', 'ShowDoc')
      .addToUi();  
}


/**
 * @OnlyCurrentDoc
 */
function ShowAPIKeyDataAvaibilityPrompt() {
  var ui = SpreadsheetApp.getUi();
  var userProperties = PropertiesService.getUserProperties();
  var api_key = userProperties.getProperty("APIKEYDATAAVAIBILITYSERVICE")

  if (api_key) {
    var result = ui.prompt('Set API Key for Data Availability Service',
                           '✅ Your API '+ api_key +' key is already set.\n\nYour CRYPTOFINANCE calls are sent to the Data Availability Proxy API.\nYou can still re-enter it below to override its current value:',
                           ui.ButtonSet.OK_CANCEL);
  }
  else {
    var result = ui.prompt('Set API Key for Data Availability Service',
                           'The Data Availability Service offers unlimited data from CoinMarketCap and helps you avoid errors due to exchanges API ban and overload.\nExchanges ban Google Sheets servers IP addresses when too many requests originate from them.\nIt also provides you with more exchanges and markets data.\n\nMore info here: https://cryptofinance.ai/data-availability-service\n\nOnce subscribed, please enter your API Key below:',
                           ui.ButtonSet.OK_CANCEL);  
  }
  
  var button = result.getSelectedButton();
  var user_input = result.getResponseText().replace(/\s+/g, '');
  if (button == ui.Button.OK) {
    if (user_input && user_input == "__DELETE__") {
      userProperties.deleteProperty("APIKEYDATAAVAIBILITYSERVICE");
      ui.alert('API Key Removed',
               'Your API Key has been sucessfully removed.\nYour requests will not be sent to the Data Availability Proxy API anymore.'
               ,ui.ButtonSet.OK);
    }
    else if (user_input && user_input.length == 36) {
      userProperties.setProperty("APIKEYDATAAVAIBILITYSERVICE", user_input);
      ui.alert('API Key successfully saved',
               'Your requests will now be sent to the CRYPTOFINANCE Data Availability Service.\nWhenever an exchange API is overloaded you will keep getting data.\n\nBe sure to refresh the cells: Select cells calling CRYPTOFINANCE (or all with Cmd+A), hit Delete key, wait 3sec,\nand then undo the delete with Cmd+Z.\n(If you\'re on Windows use the Ctrl key instead of Cmd)\n\nYou can contact support@cryptofinance.ai if you have any question.'
               ,ui.ButtonSet.OK);
    }
    else if (user_input) {
      ui.alert('API Key not valid',
               'The API Key you entered appears to be not valid.\nIf you believe this is an error, contact support@cryptofinance.ai.'
               ,ui.ButtonSet.OK);
    }
  }
}


/**
 * @OnlyCurrentDoc
 */
function ShowAPIKeyHistPlanPrompt() {
  var ui = SpreadsheetApp.getUi();
  
  var userProperties = PropertiesService.getUserProperties();
  var api_key = userProperties.getProperty("APIKEY_HISTPLAN")
  if (api_key) {
    var result = ui.prompt('Set API Key for the Historical Data plan',
                           '✅ Your API '+ api_key +' key is already set.\n\nYou can now use the historical data syntaxes in your sheet.\n\nYou can still re-enter it below to override its current value:',
                           ui.ButtonSet.OK_CANCEL);
  }
  else {
    var result = ui.prompt('Set API Key for the Historical Data plan',
                        'The Historical Data plan gives you access to hourly historical data of 196 exchanges.\nIncluding hourly open, high, low, close and volume info.\nATH (All Time High) prices and volume per exchange and custom sparklines.\n\nMore info at https://cryptofinance.ai/crypto-historical-data\n\nOnce subscribed, please enter your API Key below:',
                        ui.ButtonSet.OK_CANCEL);
  }
  var button = result.getSelectedButton();
  var text = result.getResponseText().replace(/\s+/g, '');  
  if (button == ui.Button.OK) {
    if (text && text == "__DELETE__") {
      var userProperties = PropertiesService.getUserProperties();
      userProperties.deleteProperty("APIKEY_HISTPLAN");
      ui.alert('API Key Removed',
               'Your API Key has been sucessfully removed.\nYour requests will not be sent to the Data Availability Proxy API anymore.'
               ,ui.ButtonSet.OK);
    }
    else if (text && text.length == 36) {
      var userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty("APIKEY_HISTPLAN", text);
      ui.alert('API Key successfully saved',
               'You can now use the historical data syntaxes in your sheet.\n\nIt looks like this: =CRYPTOFINANCE("BTC/USD", "price", "2017-12-25@17:00")\n\nSee the doc for all options: https://cryptofinance.ai/docs/cryptocurrency-bitcoin-historical-prices/\n'
               ,ui.ButtonSet.OK);
    }
    else if (text) {
      ui.alert('API Key not valid',
               'The API Key you entered appears to be not valid.\nIf you believe this is an error, contact support@cryptofinance.ai.'
               ,ui.ButtonSet.OK);
    }
  }
}


/**
 * @OnlyCurrentDoc
 */
function ShowDoc() {
  var ui = SpreadsheetApp.getUi()
  ui.alert("Documentation and Info",
           'Official website: https://cryptofinance.ai\n\
            Documentation: https://cryptofinance.ai/docs/\n\
            Support email: support@cryptofinance.ai',
            ui.ButtonSet.OK)
}



/**
 * @OnlyCurrentDoc
 */
function ShowRefreshInfo() {
  var ui = SpreadsheetApp.getUi()
  ui.alert("How to refresh rates",
           'We recommend setting up a manual refresh trigger.\nSee the doc at this address for more info:\nhttps://cryptofinance.ai/docs/how-to-refresh-rates/',
            ui.ButtonSet.OK)
}


/**
 * Returns cryptocurrencies price and other info.
 *
 * @param {"EXCHANGE:BASE/QUOTE"} market The market data to fetch, default is BTC/USD. Default EXCHANGE is CoinMarketCap. When no QUOTE currency is set, default is USD.
 * @param {"price|marketcap|volume|change|name|rank|max_supply|total_supply|circulating_supply"} attribute What to return, default is price. Some exchanges provide more info than others. Refer to the documentation for the full list. 
 * @param {"Any supported period or date or empty string"} option Used to narrow down the attribute value by date or period. Use an empty string "" if you want to use a cell to force the refresh. Different attributes have different options. Refer to the doc for the supported syntaxes.
 * @param {"Empty cell reference"} refresh_cell ONLY on 4th argument. Reference an empty cell and change its content to force refresh the rates. See the doc for more info.
 * @return The current or historical price, volume, change, marketcap, name, rank, and more.
 * @customfunction
 */
function CRYPTOFINANCE(market, attribute, option, refresh_cell) {
  
  // Sanitize input
  var market = (market+"") || "";
  var attribute = (attribute+"") || "";
  var option = (option+"") || "";

  // Get user anonymous token (https://developers.google.com/apps-script/reference/base/session#getTemporaryActiveUserKey())
  // Mandatory to authenticate request origin
  var GSUUID = encodeURIComponent(Session.getTemporaryActiveUserKey());
  // Get Data Availability Service and Historical Plan API Keys, if any
  var userProperties = PropertiesService.getUserProperties();
  var APIKEYDATAAVAIBILITYSERVICE = userProperties.getProperty("APIKEYDATAAVAIBILITYSERVICE") || "";
  var APIKEY_HISTPLAN = userProperties.getProperty("APIKEY_HISTPLAN") || "";
  
  // Fetch data
  try {

    var data = {};
    var CACHE_KEY = "CF__"+ market.toLowerCase() + "_" + attribute.toLowerCase() + "_" + option.toLowerCase();
    // First check if we have a cached version
    var cache = CacheService.getUserCache();
    var cached = cache.get(CACHE_KEY);
    if (cached && cached != null && cached.length > 1) {
      data = JSON.parse(cached)
    }
    // Else, fetch it from API and cache it
    else {
      var url = "https://api.cryptofinance.ai/v1/data?histplanapikey=" + APIKEY_HISTPLAN + "&gsuuid=" + GSUUID + "&dataproxyapikey=" + APIKEYDATAAVAIBILITYSERVICE;
      url += "&m=" + market;
      url += "&a=" + attribute;
      url += "&o=" + option;
      // Send request
      var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true, validateHttpsCertificates: true})
      data = JSON.parse(response.getContentText());
      // Stop here if there is an error
      if (data["error"] != "") {
        throw new Error(data["error"])
      }
      // If everything went fine, cache the raw data returned
      else if (response && response.getResponseCode() == 200 && data.length > 1 && data.length < 99900) {
        cache.put(CACHE_KEY, response.getContentText(), data["caching_time"] || 60)
      }
    }

    // Return with the proper type
    var out = "-";
    if (data["type"] == "float") {
      out = parseFloat(data["value"]);
    }
    else if (data["type"] == "int") {
      out = parseInt(data["value"]);
    }
    else if (data["type"] == "csv") {
      out = Utilities.parseCsv(data["value"]);
    }
    else {
      out = data["value"]
    }
    return out;

  }

  catch (e) {
    throw new Error(e.message)
  }
  
}
