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
      .addItem('Enter Cryptowatch Public API Key', 'ShowAPIKeyCW')
      .addSeparator()
      .addItem('Documentation', 'ShowDoc')
      .addToUi();  
}


/**
 * @OnlyCurrentDoc
 */
function ShowAPIKeyCW() {
  var ui = SpreadsheetApp.getUi();
  var userProperties = PropertiesService.getUserProperties();
  var api_key = userProperties.getProperty("CWPUBLICAPIKEY")

  if (api_key) {
    var result = ui.prompt('Enter Cryptowatch Public API Key',
                           '✅ Your API '+ api_key +' key is already set.\n\nYou can still re-enter it below to override its current value:',
                           ui.ButtonSet.OK_CANCEL);
  }
  else {
    var result = ui.prompt('Enter Cryptowatch Public API Key',
                           'You can subscribe to different sets of data attributes from your Cryptowatch account page.\n\nOnce subscribed, please enter your Public API Key below:',
                           ui.ButtonSet.OK_CANCEL);  
  }
  
  var button = result.getSelectedButton();
  var user_input = result.getResponseText().replace(/\s+/g, '');
  if (button == ui.Button.OK) {
    if (user_input && user_input == "__DELETE__") {
      userProperties.deleteProperty("CWPUBLICAPIKEY");
      ui.alert('API Key Removed',
               'Your API Key has been sucessfully removed.'
               ,ui.ButtonSet.OK);
    }
    else if (user_input && (user_input.length == 20 || user_input.length == 36) {
      userProperties.setProperty("CWPUBLICAPIKEY", user_input);
      ui.alert('API Key successfully saved',
               'If you haven\t yet, you can subscribe to different sets of data attributes from your Cryptowatch account page.\n\nContact support@cryptofinance.zendesk.com if you have any question.'
               ,ui.ButtonSet.OK);
    }
    else if (user_input) {
      ui.alert('API Key not valid',
               'The API Key you entered appears to be not valid.\nIf you believe this is an error, contact support@cryptofinance.zendesk.com.'
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
            Guide: https://guides.cryptowat.ch/google-sheets-plugin#how-do-i-get-started\n\
            Support email: support@cryptofinance.zendesk.com',
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


function cast_matrix__(el) {
  if (el === "") {
    return "-"
  }
  else if (el.map) {return el.map(cast_matrix__);}
  try {
    var out = Number(el)
    if ((out === 0 || out) && !isNaN(out)) {
      if (el.length > 1 && el[1] == 'x') {
        return el
      }
      else {
        return out
      }
    }
    else {
      return el
    }
  }
  catch (e) {return el;}
}


/**
 * Returns cryptocurrencies price and other info.
 *
 * @param {"EXCHANGE:BASE/QUOTE"} market The exchange or asset data to fetch, default is BTC/USD. Default data source is Cryptowatch, default quote currency is USD.
 * @param {"price|volume|change|name|rank"} attribute What to return, default is price. Some exchanges provide more info than others. Refer to the documentation for the full list. 
 * @param {"Any supported period or date or empty string"} option Used to narrow down the attribute value by date or period. Use an empty string "" if you want to use a cell to force the refresh. Different attributes have different options. Refer to the doc for the supported syntaxes.
 * @param {"Empty cell reference"} refresh_cell ONLY on 4th argument. Reference an empty cell and change its content to force refresh the rates. See the doc for more info.
 * @return The latest or historical price, volume, change, marketcap, name, rank, and more.
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
  var APIKEYDATAAVAIBILITYSERVICE = userProperties.getProperty("CWPUBLICAPIKEY") || "";
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
      url += "&s=os";
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
      out = cast_matrix__(out);
    }
    else {
      out = data["value"]
    }
    return out;

  }

  catch (e) {
    var msg = e.message.replace(/https:\/\/api.*$/gi,'')
    throw new Error(msg)
  }
  
}
