/*
Google Apps Script is a scripting platform for applications in the G Suite environment of applications.
It is remarkably powerful as it has the capability to extend the functionality of ALL Google Apps applications, 
or all applications in the GSuite. Being built on JavaScript, the syntax and functionality is remarkably similar 
to that of JavaScript. Any regular JavaScript developer with a flair for utilising the Google Apps APIs will find
it very easy to do a great deal more than what Google Apps provides to the end user by default.

A simple example would be to extend the functionality of Google sheets. Sheets as an application already provides 
developers with APIs that allows a developer to consume them. However, using Apps Script, creating additional 
functionality along with menus, their icons and dialog windows can be done - pretty easily. A similar analogy can be
the VBSript of Microsoft Office Applications that was used to extend the Office apps.

This is a simple Apps Script parser of the AMFI Mutual Fund NAVs. It will give you a clear picture how you can 
quickly read from a file, process the information and then display the same in Google Sheets that users can then 
see and act upon.
*/

function myFunction() {
  
  /*
  setting the headers information for easily assigning values to array
  Scheme Code;ISIN Div Payout/ ISIN Growth;ISIN Div Reinvestment;Scheme Name;Net Asset Value;Date*/
  var _SCHEMECODE   = 0;
  var _ISIN         = 1;
  var _REINVESTMENT = 2;
  var _SCHEMENAME   = 3;
  var _NAV          = 4;
  var _DATE         = 5;
  
  /*
  starting row and column index for the sheet where we print the values
  the separator is not being used as of now. Except in the logger, ie the Console equivalent
  for AppScript
  */
  var startRow  = 2;
  var startCol  = 2;
  var separator = ' | ';
  
  //capture the current spreadsheet's active sheet and assign to 'sheet'
  var sheet = SpreadsheetApp.getActiveSheet();
  var url = 'https://www.amfiindia.com/spages/NAVOpen.txt?t=';
  //capture the value of date and append it to url. fetch result url to get the latest results
  url += sheet.getRange(1,1).getValue();
  var html = UrlFetchApp.fetch(url).getContentText();
  
  //set some variables to default values
  var htmlLength = html.length;
  var rowStart   = 0;
  var start      = 0;
  var counter    = 0;
  var rowEnd     = html.length;
  var _arr       = {};  // this is our temporary array that stores the individual lines from the document
  var remainingText = html.slice(start, html.length);
  
  sheet.getRange(rowStart+1, 2).setValue(url);
  
  //set sheet header text. see how legible the array keys are since they are set this way
  sheet.getRange(startRow + counter,startCol + _SCHEMECODE).setValue("SCHEME CODE");
  sheet.getRange(startRow + counter,startCol + _ISIN).setValue("ISIN");
  sheet.getRange(startRow + counter,startCol + _REINVESTMENT).setValue("REINVESTMENT");
  sheet.getRange(startRow + counter,startCol + _SCHEMENAME).setValue("SCHEMENAME");
  sheet.getRange(startRow + counter,startCol + _NAV).setValue("NAV");
  sheet.getRange(startRow + counter,startCol + _DATE).setValue("DATE");
  
  /*
  If you open the file and check the line with the least number of possible characters, you will find 
  it to be around 120 characters. You can capture this in a variable by iterating the text once. But I 
  have avoided the same because that adds another almost unnecessary iteration - considering the fact that
  it already takes up about a couple of minutes to fetch the data.
  start a loop covering the entire text till the last text segment is at least 120 characters in length
  if you check the document outputed with url = 'url', the minimum segment length is at least 120+
  therefore, if there is any existing line with a length less than 120, it means we have covered all the 
  relevant rows.
  */
  
  while(remainingText.length > 120){
    
    /* set the start and end search strings in regex. this pair gives us each row
    the advantage in using regex is largely to reduce the amount of code required to 
    set the acceptable range of rows. You can take the help of this website for making 
    this easier. https://regex101.com
    */
    rowStart = remainingText.search('[0-9]{6};');
    rowEnd   = remainingText.search('2019|2018|2017|2016|2015|2014|2013|2012|2011|2010|2009')+4;
    
    var _row = remainingText.substring(rowStart, rowEnd);
    
    // split it up
    _arr = _row.split(";");
    
    sheet.getRange(1 + startRow + counter,startCol + _SCHEMECODE).setValue(_arr[_SCHEMECODE]);
    sheet.getRange(1 + startRow + counter,startCol + _ISIN).setValue(_arr[_ISIN]);
    sheet.getRange(1 + startRow + counter,startCol + _REINVESTMENT).setValue(_arr[_REINVESTMENT]);
    sheet.getRange(1 + startRow + counter,startCol + _SCHEMENAME).setValue(_arr[_SCHEMENAME]);
    sheet.getRange(1 + startRow + counter,startCol + _NAV).setValue(_arr[_NAV]);
    sheet.getRange(1 + startRow + counter,startCol + _DATE).setValue(_arr[_DATE]);
    
    /* this is typical iteration. each time, we can just start from where we covered in the 
    previous iteration. the text is substringed accordingly. */
    start = rowEnd; counter +=1;
    
    remainingText = returnRestofText(remainingText, start);
    
    /*
    this section is entirely a log of the processing that occurs. if you are debugging, 
    you can uncomment this entire section
    */
    Logger.log("---------------------------------------------------------");
    Logger.log("Start - " + rowStart + separator + "End - " + rowEnd);
    Logger.log(html.length + separator + remainingText.length);
    Logger.log(_arr[_SCHEMECODE]+separator+_arr[_ISIN]+separator+_arr[_REINVESTMENT]+
             separator+_arr[_SCHEMENAME]+separator+_arr[_NAV]+separator+_arr[_DATE]);
  }
  
};

/* 
returnRestofText function is to provide some abstraction. I could have gone with the default substring directly
but i feel i will require this function later as I extend the functionality and set rules
in here directly.
*/
function returnRestofText( _str, _start){
  var _text = _str.substring(_start);
  return _text;
};

/*
If you notice carefully, it calls the SpreadsheetApp method getUI(). 
This is then used to call the createMenu function with the parent text 'Fetch Data' along with 
sub menu 'Fetch all rows'. Additionally it binds the function to the sub menu item. This allows
the menu to be added to the sheet as it is loaded.
*/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Fetch Data')
      .addItem('Fetch all rows','myFunction')
      .addToUi();
};
