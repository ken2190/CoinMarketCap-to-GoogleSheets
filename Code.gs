function onOpen(e) {
  // Add a custom 'Scripts' menu to the toolbar.
  SpreadsheetApp.getUi()
      .createMenu('Scripts')
      .addItem('Get CoinMarketCap data', 'getCMCData')
      .addToUi();
}

function getCMCData() {
  // --------
  // SETTINGS
  // --------
  //
  // Enter your API Key from CoinMarketCap here (just sign up for a free account)
  const API_KEY = 'your-api-key';
  //
  // Enter the number of cryptocurrencies you want to display ex. 100 for the top 100
  // It costs 3 credits per call per 200 cryptocurrencies and you get 10,000 call credits per month with the free plan
  const CRYPTOCURRENCIES = 200;
  //
  // This is a list of the supported data by this script. If you want to omit some data, comment it out by using //, ex. // 'cmc_rank'
  // You can also sort this list to display the data in the column order that you want.
  const DATA_TO_DISPLAY = [
    'cmc_rank',
    'name',
    'symbol',
    'market_cap',
    'price',
    'volume_24h',
    'num_market_pairs',
    'percent_change_1h',
    'percent_change_24h',
    'percent_change_7d',
    'circulating_supply',
    'total_supply',
    'max_supply',
    'date_added',
  ];
  //
  // This is the description of the data that will appear in bold on the first row.
  // Feel free to modify the description (right hand side only).
  // No need to comment out or sort anything in this list, it gets handled automatically.
  const DATA_DESCRIPTION = {
    'cmc_rank': '#',
    'name': 'Name',
    'symbol': 'Symbol',
    'market_cap': 'Market Cap',
    'price': 'Price',
    'volume_24h': 'Volume (24h)',
    'num_market_pairs': 'Market Pairs',
    'percent_change_1h': 'Change (1h)',
    'percent_change_24h': 'Change (24h)',
    'percent_change_7d': 'Change (7d)',
    'circulating_supply': 'Circulating Supply',
    'total_supply': 'Total Supply',
    'max_supply': 'Max Supply',
    'date_added': 'Date Added',
  };
  // ------------
  // SETTINGS END
  // ------------
  
  // Do the request to CoinMarketCap API
  const url = 'https://pro-api.coinmarketcap.com/v1/cryptocurrency/listings/latest?limit=' + CRYPTOCURRENCIES.toString();
  const headers = {
    'X-CMC_PRO_API_KEY': API_KEY
  };
  const options = {
     "headers": headers,
  };
  const response = UrlFetchApp.fetch(url, options);

//  Logger.log(response);
     
  // Parse JSON
  const json = response.getContentText();
  const data = JSON.parse(json);
    
  var description = [];
  for (var i in DATA_TO_DISPLAY) {
    description.push(DATA_DESCRIPTION[DATA_TO_DISPLAY[i]]);
  }
  
  var array = [];
  array.push(description);
  for (var i in data.data) {
    var asset = [];
    for (var j in DATA_TO_DISPLAY) {
      switch (DATA_TO_DISPLAY[j]) {
        case 'cmc_rank':
        case 'name':
        case 'symbol':
        case 'num_market_pairs':
        case 'circulating_supply':
        case 'total_supply':
        case 'max_supply':
          asset.push(data.data[i][DATA_TO_DISPLAY[j]]);
          break;
        case 'market_cap':
        case 'price':
        case 'volume_24h':
          asset.push(data.data[i].quote.USD[DATA_TO_DISPLAY[j]]);
          break;
        case 'percent_change_1h':
        case 'percent_change_24h':
        case 'percent_change_7d':
          asset.push(data.data[i].quote.USD[DATA_TO_DISPLAY[j]]/100);
          break;
        case 'date_added':
          asset.push(data.data[i].date_added.split('T')[0]);
      }
    }
    array.push(asset);
  }
  
  // Get spreadsheet and sheet (create one if it doesn't exist)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('CoinMarketCap');
  if (sheet === null) {
    sheet = spreadsheet.insertSheet('CoinMarketCap');
  }
  sheet.clear()
  
  // Write to sheet
  sheet.getRange(1, 1, array.length, DATA_TO_DISPLAY.length).setValues(array);
  
  // Formatting
  var rules = sheet.getConditionalFormatRules();
  const nameColumn = DATA_TO_DISPLAY.indexOf('name') + 1;
  if (nameColumn > 0) {
    sheet.getRange(2, nameColumn, CRYPTOCURRENCIES).setFontWeight('bold');
  }
  const marketCapColumn = DATA_TO_DISPLAY.indexOf('market_cap') + 1;
  if (marketCapColumn > 0) {
    sheet.getRange(2, marketCapColumn, CRYPTOCURRENCIES).setNumberFormat('$#,##0');
  }
  const priceIndex = DATA_TO_DISPLAY.indexOf('price');
  const priceColumn = priceIndex + 1;
  if (priceColumn > 0) {
    for (var assetIndex = 1; assetIndex < array.length; assetIndex++) {
      var assetRow = assetIndex + 1;
      var price = array[assetIndex][priceIndex];
      if (price >= 1) {
        sheet.getRange(assetRow, priceColumn).setNumberFormat('$#,##0.00');
      } else {
        sheet.getRange(assetRow, priceColumn).setNumberFormat('$#,##0.000000');
      }
    }
    sheet.getRange(2, priceColumn, CRYPTOCURRENCIES).setFontColor('#005ce6');    
  }
  const volumeColumn = DATA_TO_DISPLAY.indexOf('volume_24h') + 1;
  if (volumeColumn > 0) {
    const volumeRange = sheet.getRange(2, volumeColumn, CRYPTOCURRENCIES);
    volumeRange.setNumberFormat('$#,##0');
    volumeRange.setFontColor('#ff9900');
  }
  const change1hColumn = DATA_TO_DISPLAY.indexOf('percent_change_1h') + 1;
  if (change1hColumn > 0) {
    const change1hRange = sheet.getRange(2, change1hColumn, CRYPTOCURRENCIES)
    change1hRange.setNumberFormat('#,##0.00%');
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#cc0000').setRanges([change1hRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor('#009900').setRanges([change1hRange]).build());
  }
  const change24hColumn = DATA_TO_DISPLAY.indexOf('percent_change_24h') + 1;
  if (change24hColumn > 0) {
    const change24hRange = sheet.getRange(2, change24hColumn, CRYPTOCURRENCIES)
    change24hRange.setNumberFormat('#,##0.00%');
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#cc0000').setRanges([change24hRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor('#009900').setRanges([change24hRange]).build());
  }
  const change7dColumn = DATA_TO_DISPLAY.indexOf('percent_change_7d') + 1;
  if (change7dColumn > 0) {
    const change7dRange = sheet.getRange(2, change7dColumn, CRYPTOCURRENCIES)
    change7dRange.setNumberFormat('#,##0.00%');
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setFontColor('#cc0000').setRanges([change7dRange]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setFontColor('#009900').setRanges([change7dRange]).build());
  }
  const circulatingSupplyColumn = DATA_TO_DISPLAY.indexOf('circulating_supply') + 1;
  if (circulatingSupplyColumn > 0) {
    sheet.getRange(2, circulatingSupplyColumn, CRYPTOCURRENCIES).setNumberFormat('#,##');
  }
  const totalSupplyColumn = DATA_TO_DISPLAY.indexOf('total_supply') + 1;
  if (totalSupplyColumn > 0) {
    sheet.getRange(2, totalSupplyColumn, CRYPTOCURRENCIES).setNumberFormat('#,##');
  }
  const maxSupplyColumn = DATA_TO_DISPLAY.indexOf('max_supply') + 1;
  if (maxSupplyColumn > 0) {
    sheet.getRange(2, maxSupplyColumn, CRYPTOCURRENCIES).setNumberFormat('#,##');
  }
  const dateAddedColumn = DATA_TO_DISPLAY.indexOf('date_added') + 1;
  if (dateAddedColumn > 0) {
    sheet.getRange(2, dateAddedColumn, CRYPTOCURRENCIES).setNumberFormat('d MMM yyyy');
  }
  sheet.setConditionalFormatRules(rules);
  
  // Layout
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, DATA_TO_DISPLAY.length).setFontWeight('bold');
  sheet.autoResizeColumns(1, DATA_TO_DISPLAY.length);
}
