# CoinMarketCap to Google Sheets

A Google Apps Script to get CoinMarketCap API data into Google Sheets, nicely formatted and at the touch of a button.

![CoinMarketCap to Google Sheets screenshot](https://github.com/taorobot/CoinMarketCap-to-GoogleSheets/blob/master/cmc_googlesheets_screenshot.png?raw=true)

## Usage

This script adds a menu to the toolbar named 'Scripts' with a button named 'Get CoinMarketCap data'. You just need to click that button and the data from CoinMarketCap will appear nicely formatted in a new sheet named CoinMarketCap. If you already have a sheet named CoinMarketCap it will just update that sheet.

You can edit the Settings section of the script to modify the number of cryptocurrencies returned by the API and the columns of data that you want displayed. Read more below in Edit the script.

## Installation

1. Get your API key from https://coinmarketcap.com/api/. Just the free Basic plan will do.
2. Create a new Google Sheet or go to an existing one and click on Tools -> Script editor. This will open the editor on a new tab.
3. Delete all the code there and copy and paste all the code from Code.gs. You can also edit the name of the project if you want ex. name it 'CoinMarketCap to Google Sheets'.
4. Add your API key in the code at: const API_KEY = 'your-api-key';
5. Save the project from: File -> Save
6. Go back to your sheet and refresh the page. After a little bit the Scripts menu should appear on the toolbar.
7. Click on Scripts -> Get CoinMarketCap data. A popup should appear saying Authorization Required. Click on Continue.
8. Another popup should appear now asking you to choose your Google account. Click on your Google account.
9. Now a window should appear saying 'This app isn't verified: This app hasn't been verified by Google yet. Only proceed if you know and trust the developer.'
10. Assuming you trust the code here (if you don't, ask a developer friend of yours to have a look), click on Advanced at the bottom left.
11. Tap on 'Go to CoinMarketCap to Google Sheets (unsafe)'. If you used a different name for your app it will show here instead.
12. The next window will say 'CoinMarketCap to Google Sheets wants to access your Google Account'. Click on Allow.
13. The script should run now and the data should appear in the CoinMarketCap sheet :)

## Edit the script

### Edit the number of cryptocurrencies returned

You can go to the SETTINGS section that I marked in the script to edit the number of cryptocurrencies you want returned. The default is set to 200, meaning the top 200 cryptocurrencies by market cap will be returned. It costs 3 credits per call per 200 cryptocurrencies to get the data. You get 10,000 call credits per month with the free plan, which is more than enough for casual use.

### Edit the columns of data to display

To edit the columns that you want displayed, modify the DATA_TO_DISPLAY array. You can comment out the data types that you don't want returned by using //. ex. // 'cmc_rank'. You can also sort this array to change the order of the columns.

### Change the name of the columns

To change the names of the columns, modify the DATA_DESCRIPTION array. You should only change the right hand side of each data type here. You don't need to sort or comment out anything if you modified the DATA_TO_DISPLAY array, it will get handled automatically.
