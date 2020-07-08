function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Credit Card Sums', 'creditCardSums')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
          .addItem('Second item', 'menuItem2'))
      .addToUi();
};

function creditCardSums() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Credit Card Statements'), true);
  var startRow = 2
  start_dates = spreadsheet.getRangeByName("Start Date")
  end_dates = spreadsheet.getRangeByName("End Date")
  Logger.log(start_dates.getNumColumns());
  
  //SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  //   .alert(spreadsheet.getSheetName());
};

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
};

function Duplicate_Budget() {
  var spreadsheet = SpreadsheetApp.getActive();
  var default_style = spreadsheet.getRange('A1').getTextStyle();
  var curSheet = spreadsheet.getActiveSheet();
  var curSheet_name = curSheet.getSheetName();

  // var category = "-Budget";
  // if (curSheet_name.includes(category)){
  // if (curSheet_name.includes('Budget') || curSheet_name.includes('Spending')){
  if (curSheet_name.indexOf('Budget') !== -1 || curSheet_name.indexOf('Spending') !== -1){
    var sheet_index = curSheet.getIndex();
    var text_as_array = curSheet_name.split("-");
    var category = text_as_array[2];
    // var date_string = '20'.concat(text_as_array[0],'-', text_as_array[1]);
    var date_str = '2020-Mar';
    var date = new Date(date_str);
    var newDate = new Date(date.setMonth(date.getMonth()+1));
    var new_date_formatted = date.getFullYear().toString().slice(-2).concat('-', newDate.toString().slice(4,7),'-', category);
    Logger.log(date);
    Logger.log(date_str);
    Logger.log(newDate);
    Logger.log(date.getFullYear().toString().slice(-2))
    Logger.log(new_date_formatted);
    spreadsheet.duplicateActiveSheet().setName(new_date_formatted);
    ss = SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(sheet_index);
  };
};
