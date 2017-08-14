/**********************************************************************************
                               SIDEBAR CALCULATOR
                                                     
                                                     Developer: Jonathan Thornton
 Related:                                                    
 onOpen.gs
 onEdit.gs
 calcSidebar.html
 
 Opens Google Sheets side bar with various functions as well as an EMG Tools menu.
 *********************************************************************************/

/*******************************************************************************************************
                                            AUTHENTICATION
 Authentication for Admin console. Very secure!
 ******************************************************************************************************/
function getPassword(){
  var password_token = "test";
  return password_token;
}

function authenticate(input){
  var password_token = getPassword();
  if (input === password_token){
    return true;
  } else {
    return false;
  }
}

function calcSideBar(){
 var htmlCalc = HtmlService.createHtmlOutputFromFile('calcSidebar')
   .setTitle('Calculator').setWidth(300);
 
 SpreadsheetApp.getUi().showSidebar(htmlCalc);
}

function caseStepsWindow(){  
  var html = HtmlService.createHtmlOutputFromFile('caseStepsDialog')
    .setWidth(625)
    .setHeight(425);
  SpreadsheetApp.getUi()
    .showModalDialog(html, ' ');
}

function adminDeniedWindow(){
  var html = HtmlService.createHtmlOutputFromFile('adminDenied')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi()
    .showModalDialog(html, ' ');
}
function openAdminWindow(){
    var html = HtmlService.createHtmlOutputFromFile('adminConsole')
    .setWidth(600)
    .setHeight(400);
    SpreadsheetApp.getUi()
    .showModalDialog(html, ' ');
}

function authenticationPrompt(){
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.prompt('Please login:',
    ui.ButtonSet.OK_CANCEL);
  
  var button = result.getSelectedButton();
  var input = result.getResponseText();
  var passCheck = getPassword();
  Logger.log("button = " + button);
  Logger.log("input = " + input);
  
  if(button == ui.Button.OK){
    var passCheck = authenticate(input);
    if(passCheck == true){
      Logger.log("pass " + passCheck);
      openAdminWindow();
    } else {
      adminDeniedWindow();
    }
  } else if(button = ui.Button.CANCEL){
    return;
  }
}

function adminSetCalls(inbound, outbound) {
  Logger.log('adminSetCalls! ' + inbound + ' ' + outbound);
  /*var curDate = formatDate();
  var row = findCell(curDate);
  Logger.log('row = ' + row);
  
  if(cell === -1)
    return false;
  else {
    var outboundCell = sheet.getRange(row, 1);
    var inboundCell = sheet.getRange(row, 2);
    outboundCell.setValue(outbound);
    inboundCell.setValue(inbound);
    return true;
  }*/
}
    
function subCalc(waste, recy, del, adm, frf, erf, county, frfObj, bool){
  //TODO STATIC Global Variables pulled from sheets
  var admFee = Number('5.25');
  
  if(bool===true){
    var frFee = frfObj;
    } else if(frf===true){
      var frFee = Number('10.28');
  }
   
  var erFee = Number('15.00');
  var waste = Number(waste);
  var recy = Number(recy);
  var selCounty = county;

  var del = Number(del);
  
  Logger.log('waste= ' + waste);
  Logger.log('recy= ' + recy);
  Logger.log('del= ' + del);
  Logger.log('adm= ' + adm);
  Logger.log('frf= ' + frf);
  Logger.log('erf= ' + erf);
  Logger.log('admFee= ' + admFee);
  Logger.log('frFee= ' + frFee);
  Logger.log('erFee= ' + erFee);
  Logger.log('county = ' + selCounty);
  Logger.log('bool = ' + bool);
  
  //Calculation for First Quarter
  var total = (waste + recy + del);
  if(frf === true){
    total = total * percentify(frFee);
    Logger.log('frf total ' + total);
  }
  if(erf === true){
    total = total * percentify(erFee);
    Logger.log('erf total ' + total);
  }
  if(adm === true){
    total = (total + admFee);
    Logger.log('adm total ' + total);
  }
  total = total.toFixed(2);
  Logger.log('total= ' + total);
  
  //If Minnesota calculate tax
  if(selCounty === "Not Applicable") {
    var addTax = 0;
  } else {
     var addTax = taxMN(total,selCounty);
     Logger.log('added tax return val = ' + addTax);
  }
  Logger.log('added tax ' + addTax);
  total = (Number(addTax)+ Number(total));
  Logger.log('FINAL TOTAL = ' + total);
  
  return Number(total);
 
 }
 
//Calculate Tax rates for select MN counties
function taxMN(total, county){
  var selCounty = county;
  
  //Tax rate for MN counties
  var henTax = .09;
  var washTax = .35;
  var ramTax = .28;
  var sternTax = .1;
  var stateTax = .0975;
  var stTax =  total * stateTax;
  var addTax = 0;
  
  //calculate tax for each county
  if (selCounty === "Hennepin") {
     var coTax = (total*henTax);
     var stTax = (total*stateTax);
     var addTax = (coTax + stTax);
     Logger.log('Hen addTax = ' + addTax);
  } else {
      if(selCounty === "Washington") {
        var coTax = (total*washTax);
        var stTax = (total*stateTax);
        var addTax = (coTax + stTax);
        Logger.log('Wash addTax = ' + addTax);
      } else {
        if (selCounty === "Ramsey") {
         var coTax = (total*ramTax);
         var stTax = (total*stateTax);
         var addTax = (coTax + stTax);
         Logger.log('Ram addTax = ' + addTax);
        } else {
          if (selCounty === "Sterns") {
           var coTax = (total*sternTax);
           var stTax = (total*stateTax);
           var addTax = (coTax + stTax); 
           Logger.log('Sterns addTax = ' + addTax);
        }
      }
    }
  }
  Logger.log('FINAL ADDED TAX = ' + addTax);
  return Math.round((addTax*100)/100);
}

//Reverse Percentage Calculator
function reversePercentage(rate, adm, frf, erf, frfObj, bool){
  var cAdm = getAdmin();
  //var cFrf = getFRF();
  var cErf = getERF();
  
  Logger.log("rate = " + rate);
  Logger.log("adm = " + adm + " | cAdm = " + cAdm);
  Logger.log("frf = " + frf + " | cFrf = " + cFrf);
  Logger.log("erf = " + erf + " | cErf = " + cErf);
  
  if(bool===true){
    var cFrf = frfObj;
    } else if(frf===true){
      var cFrf = Number('10.28');
  }
  
  if(frf===true && erf===true){
    var perc = cFrf + cErf;
  } else {
    if(frf===true){
      var perc = cFrf;
    } else {
      if(erf===true){
        var perc = cErf;
      } else {
        var perc = 0;
      }
    }
  }
  Logger.log("perc = " + perc);
  
  if(adm===true){
    var rate = rate-cAdm;
  }  
  var total = Math.round((rate/((perc/100)+1)*100)/100);
  Logger.log("total = " + total);
  return total; 
}

//Calculate 2 year contract rates
function contractCalc(basePrice){
  var increase = 10;
  var moPrice = Math.round(((basePrice/3)*100)/100);
  var yr2Price = Math.round((moPrice*percentify(increase)));
  
  Logger.log('moPrice = ' + moPrice);
  Logger.log('yr2Price = ' + yr2Price);
  
  var a = new Array(moPrice, yr2Price);
  return a;
}



function frfPrompt(){
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.prompt('Enter the FRF override percentage:',
    ui.ButtonSet.OK_CANCEL);
  
  var button = result.getSelectedButton();
  var rate = result.getResponseText();
  
  var frfObj = new Object();
  
  if(button == ui.Button.OK){
    Logger.log('rate = ' + rate);
    frfObj.rate = rate;
    frfObj.bool = true;
    return frfObj; 
  } else if(button = ui.Button.CANCEL) {
    frfObj.bool = false;
    return frfObj;
  }
}

function adminWindow(){
  var html = HtmlService.createHtmlOutputFromFile('caseStepsDialog')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi()
    .showModalDialog(html, ' ');
}


/*********************************************************************************************
                                            CASE STEPS
                                            
--DEPRECATED 8/10/2017 - deprecated in favor of Calculator Case Steps Modal Window
*********************************************************************************************/
/*
function caseStepsPriceChange(){
  var caseStepsTab = "Case Steps";
  var sheet = SpreadsheetApp.getActive().getSheetByName(caseStepsTab);
  SpreadsheetApp.setActiveSheet(sheet);
  var range = sheet.getRange(3,1);
  SpreadsheetApp.setActiveRange(range);
}

function caseStepsSignedContract(){
  var caseStepsTab = "Case Steps";
  var sheet = SpreadsheetApp.getActive().getSheetByName(caseStepsTab);
  SpreadsheetApp.setActiveSheet(sheet);
  var range = sheet.getRange(21,1);
  sheet.hideRows(2,20);
  SpreadsheetApp.setActiveRange(range);
  sheet.showRows(2,20);
}

function caseStepsCancelResi(){
  var caseStepsTab = "Case Steps";
  var sheet = SpreadsheetApp.getActive().getSheetByName(caseStepsTab);
  SpreadsheetApp.setActiveSheet(sheet);
  var range = sheet.getRange(43,1);
  sheet.hideRows(2,42);
  SpreadsheetApp.setActiveRange(range);
  sheet.showRows(2,42);
}

function caseStepsCredit(){
  var caseStepsTab = "Case Steps";
  var sheet = SpreadsheetApp.getActive().getSheetByName(caseStepsTab);
  SpreadsheetApp.setActiveSheet(sheet);
  var range = sheet.getRange(55,1);
  sheet.hideRows(2,54);
  SpreadsheetApp.setActiveRange(range);
  sheet.showRows(2,54);
}

function caseStepsTBICredit(){
  var caseStepsTab = "Case Steps";
  var sheet = SpreadsheetApp.getActive().getSheetByName(caseStepsTab);
  SpreadsheetApp.setActiveSheet(sheet);
  var range = sheet.getRange(75,1);
  sheet.hideRows(2,74);
  SpreadsheetApp.setActiveRange(range);
  sheet.showRows(2,74);
}

function caseStepsServiceChange(){
  var caseStepsTab = "Case Steps";
  var sheet = SpreadsheetApp.getActive().getSheetByName(caseStepsTab);
  SpreadsheetApp.setActiveSheet(sheet);
  var range = sheet.getRange(85,1);
  sheet.hideRows(2,87);
  SpreadsheetApp.setActiveRange(range);
  sheet.showRows(2,87);
}*/

/*******************************************************************************************************************
                                              UTILITY FUNCTIONS
*******************************************************************************************************************/
/* 
Using 24 hour clock, stamps the spreadsheet with the current HOUR the entry was made in
*/
function timeStamp(event, activeSheet){ 
  // TODO: Add functionality for daylight savings time...
  
  var timezone = "GMT-5";
  var timestamp_format = "HH"; // Timestamp Format. 
  var updateColName = "SALES REP ID";
  var timeStampColName = "TIME STAMP";
  var sheet = event.source.getSheetByName(activeSheet);
  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues(); // METHOD SYNTAX: getrange(row, column, numRows, numColumns)
  var dateCol = headers[0].indexOf(timeStampColName);
  var updateCol = headers[0].indexOf(updateColName); 
  
  updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol) { // only timestamp if 'Last Updated' header exists, but not in the header row itself
    var cell = sheet.getRange(index, dateCol + 1);
    var date = Utilities.formatDate(new Date(), timezone, timestamp_format);
    cell.setValue(date);
  }
}

/*
  Sets Date cell to current Date 
*/
function dateStamp(event, activeSheet){
  var updateColName = "SALES REP ID";
  var dateColName = "DATE";
  var sheet = event.source.getSheetByName(activeSheet);
  var actRng = event.source.getActiveRange();
  var editColumn = actRng.getColumn();
  var index = actRng.getRowIndex();
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues(); // METHOD SYNTAX: getrange(row, column, numRows, numColumns)
  var dateCol = headers[0].indexOf(dateColName);
  var updateCol = headers[0].indexOf(updateColName); 
  
  updateCol = updateCol+1;
  if (dateCol > -1 && index > 1 && editColumn == updateCol) { // only timestamp if 'Last Updated' header exists, but not in the header row itself
    var cell = sheet.getRange(index, dateCol + 1);
    cell.setValue(formatDate());
  }
}

/*
  Brings user to the first blank row
*/
  
function jumpToLastLine(){
  var sheet = SpreadsheetApp.getActiveSheet();
  return sheet.setActiveRange(sheet.getRange(sheet.getDataRange().getHeight() + 1, 1));
}

/*
  returns current Date formatted as YYYYMMDD
*/

function formatDate() {
  var date = new Date();
  return date.getFullYear() + ("0" + (date.getMonth() + 1)).slice(-2) + ("0" + date.getDate()).slice(-2);
}

//Utility to change number to a multiplicable value 
 function percentify(n){
     var p = Number((n/100)+1);
     Logger.log('percent n= ' + n + ' percent p ' + p);
    
     return p;
  }

//CURRENT FEES 
// ***** TODO: NOTE GET THESE VALUES FROM SHEET ADMIN ONLOAD


function feeObj(){
  var ss = SpreadsheetApp.openById("1374546442");
  SpreadsheetApp.setActiveSheet(ss);
  var admCell = ss.getRange("B1");
  var frfCell = ss.getRange("B2");
  var erfCell = ss.getRange("B3");
  
  var feeO = new Fees(admCell, frfCell, erfCell);
  return feeO;
}

function getAdmin(){
    return 5.25;
}

function getFRF(){
  return 10.28;
}

function getERF(){
  return 15.00;
}

function findCell(str){
  var search = str.toString();
  var ss = SpreadsheetApp.openById(41457026);
  SpreadsheetApp.setActiveSheet(ss);
  var range = "A:A";
  var column = sheet.getRange(range);
  var values = column.getValues();
  var row = 0;
  
  while(values[row] && values[row][0] !== search){
    row++;
  }
  
  if (values[row][0] === search)
    return row + 1;
  else
    return -1;
  
}



















  
  






































