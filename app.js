
const express = require('express');
const bodyParser = require('body-parser');
const http = require('http');

//Excel file creation
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");


const now = new Date();
const formattedDateTime = now.toLocaleString();
const headers = [["Symbol", "Price", "Stoploss", "Signal","Timeframe","CreatedAt"]];
const filepath = "Nifty.xlsx"
const sheetName ="nifty data"


const app = express();
const PORT = 80;
const HOSTNAME= 'ganesantrade'
// Middleware to parse incoming JSON requests
app.use(bodyParser.json());






app.post("/webhook", async (req, res) => {
  
  if (req.body !='')
  {
    var symbol = req.body['symbol']
   // console.log('Symbol: '+ symbol)
    var c_price = req.body['copen']

    var p_close = req.body['pclose']

    var p_open = req.body['popen']

    var p_high = req.body['phigh']

    var p_low = req.body['plow']
     var tc_price = req.body['tcprice']

    //console.log('Price: ' + price)
    //var stoploss = req.body['stoploss']
    var stoploss = p_low
    //console.log('Stoploss: ' + stoploss)
    var signal = req.body['signal']
    //console.log('Signal: ' + signal)
    var timeframe = req.body['timeframe']
    //console.log('Timeframe : ' + timeframe)
  };

 // { "symbol": "{{ticker}}", "copen":"{{open}}", "pclose":"{{close[1]}}", "popen":"{{open[1]}}", "phigh":"{{high[1]}}", "plow":"{{low[1]}}", "timeframe": "{{interval}}", "signal": "BUY" }
  
  var rows = [symbol, c_price, p_close, p_open, p_high, p_low, stoploss, tc_price, signal, timeframe, formattedDateTime];
  createdocumentExcel(sheetName,filepath,rows)
 res.send("OK");
});

async function createdocumentExcel(sheetName, filePath, drow) {

var workbook;
var worksheet;

 if (!filePath || !sheetName) {
    throw new Error("Invalid input parameters");
  }

  // 1. Create a new workbook and worksheet
  workbook = new ExcelJS.Workbook();
  if (fs.existsSync(filePath)) {
      await workbook.xlsx.readFile(filePath);    
      worksheet = workbook.getWorksheet(sheetName)
      
  }
  else
  {
      //var worksheet = workbook.getWorksheet(sheetName);
      worksheet = workbook.getWorksheet(sheetName) || workbook.worksheets[0];
      
      if (!worksheet) {
        worksheet = workbook.addWorksheet(sheetName);
      // 2. Define Columns 
      // (The 'key' allows you to add rows using objects)
      worksheet.columns = [
        { header: 'Symbol', key: 'symbol', width: 25 },
        { header: 'Price', key: 'c_price', width: 20 },
        { header: 'Previous Close', key: 'p_close', width: 50 },
        { header: 'Previous Open', key: 'p_open', width: 50 },
        { header: 'Previous High', key: 'p_high', width: 20 },
        { header: 'Previous Low', key: 'p_low', width: 20 },
        { header: 'Stoploss', key: 'stoploss', width: 20 },
        { header: 'Trend Change price', key: 'tc_price', width: 20 },
        { header: 'Signal', key: 'signal', width: 20 },
        { header: 'Timeframe', key: 'timeframe', width: 20 },
        { header: 'CreatedAt', key: 'date', width: 35 },
      ];
    // 5. Optional: Add some basic styling to the header row
      worksheet.getRow(1).font = { bold: true };            
    }

 }
     
    // 1. Force the row count calculation
    // actualRowCount ignores empty formatted rows
    const nextRow = worksheet.actualRowCount + 1;
  
    // 2. Directly set values at that row index
    const row = worksheet.getRow(nextRow);
    row.values = drow;
    row.commit;
    // 5. Save the workbook back to the same path
    await workbook.xlsx.writeFile(filePath);
    console.log(`Excel file "${filePath}" created successfully!`);
}


// Start the server
app.listen(PORT,HOSTNAME, () => {
    console.log(`Secure webhook server running on ${HOSTNAME} port ${PORT}`);
});


app.get("/sample", (req, res) => {
  var msg = 'Hello World';
  console.log(msg);
  // Call external REST API here
  // axios.post("https://api.broker.com/order", req.body);

  res.send(msg);
});