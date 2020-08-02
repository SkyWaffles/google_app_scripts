function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('18-Dec-Spending'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('19-Jan-Spending'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('19-Jan-Budget'), true);
  spreadsheet.getRange('C13').activate();
  spreadsheet.getCurrentCell().setValue('write something here');
  spreadsheet.getRange('C14').activate();
};

function get_date_string(curSheet_name){
  var text_as_array = curSheet_name.split("-");
  var category = text_as_array[2];
  var date_string = '20'.concat(text_as_array[0],'-', text_as_array[1]);
  return [date_string, category]
}

function Duplicate_Budget() {
  var spreadsheet = SpreadsheetApp.getActive();
  var default_style = spreadsheet.getRange('A1').getTextStyle()
  var curSheet = spreadsheet.getActiveSheet()
  var curSheet_name = curSheet.getSheetName()
  
  if (curSheet_name.includes('Budget') || curSheet_name.includes('Spending')){
    var sheet_index = curSheet.getIndex();
    var parsed_name = get_date_string(curSheet_name)
    var date_string = parsed_name[0]
    var date = new Date(date_string);
    var category = parsed_name[1]
    var newDate = new Date(date.setMonth(date.getMonth()+1))
    var new_date_formatted = date.getFullYear().toString().slice(-2).concat('-', newDate.toString().slice(4,7),'-', category);
    if (curSheet_name.includes('Budget')){
      spreadsheet.duplicateActiveSheet().setName(new_date_formatted);
      SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(sheet_index);
    } else {
      var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();

      // check to see that Budget Sheet exists before creating Spending Sheet
      // and also store the current sheet name at the same time (to get sheet index)
      var budget_name = null; // store name of sheet if it exists
      for (var i=0 ; i<sheets.length ; i++){
        var curName = sheets[i].getName();
        var curName_ls = get_date_string(curName);
        const budget_date_string = curName_ls[0].slice(2);
        const budget_category = curName_ls[1];
        if (budget_category == 'Budget' && budget_date_string == new_date_formatted.slice(0,6)){
          budget_name = curName;
          break;
        }
      }

      if (budget_name != null) {
        // 'Spending' sheet should always come before 'Budget' sheet for that month
        NewSheetSpending(newDate, new_date_formatted, sheet_index-1);
      } else {
        throw ("Error: Budget Sheet not yet created for this month. Please create Budget Sheet first")
      }
    }
  };
};

function DeleteNumbers() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B3').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('B3:D52').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
};

function link_cell() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setFormula('=\'20-Mar-Budget\'!E2');
};

function getDaysInMonth(year_month) {
  const month = year_month.getMonth()
  var date = new Date(year_month.getFullYear(), month, 1);
  var days = [];
  while (date.getMonth() === month) {
    days.push(new Date(date));
    date.setDate(date.getDate() + 1);
  }
  return days;
}

total_records = 500;
header_rows = 2;
budget_header = [["Day","Date","Time","Spent","Item","Category","Brand","Store","Store City", "Verified","Note","Daily Spending","Residual"]]
col_date = budget_header[0].indexOf('Date') + 1;
col_date_letter = String.fromCharCode(budget_header[0].indexOf('Date') + 65);
col_day = budget_header[0].indexOf('Day') + 1;
col_category = budget_header[0].indexOf('Category') + 1;
col_verified = budget_header[0].indexOf('Verified') + 1;
col_spent = budget_header[0].indexOf('Spent');
col_spent_letter = String.fromCharCode(budget_header[0].indexOf('Spent') + 65);
col_daily_spending = budget_header[0].indexOf('Daily Spending') + 1;
col_daily_spending_letter = String.fromCharCode(budget_header[0].indexOf('Daily Spending') + 65);
col_residual = budget_header[0].indexOf('Residual') + 1;
col_residual_letter = String.fromCharCode(budget_header[0].indexOf('Residual') + 65);
function NewSheetSpending(new_month, new_name, new_index) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.duplicateActiveSheet().setName(new_name);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(new_index);
  
  const curSheet = spreadsheet.getActiveSheet();
  // clear sheet
  curSheet.getRange(header_rows, 1, header_rows+total_records, budget_header[0].length).clearContent();
  var default_style = spreadsheet.getRange('A1').getTextStyle()
  // set daily budget
  spreadsheet.getRange("A1").setValue('Daily Budget').setTextStyle(default_style);
  // setup sum for 'Daily Spending' column
  spreadsheet.getRange(col_spent_letter+"1").setValue('=SUM('+col_spent_letter+(header_rows+1)+':'+col_spent_letter+total_records+')')
  // write 'Running Residual' label above 'Item' column
  curSheet.getRange(1, budget_header[0].indexOf('Item')+1).setValue('Running Residual');
  // setup sum for 'Residual' column
  curSheet.getRange(1, col_category).setValue('=SUM('+col_residual_letter+(header_rows+1)+':'+col_residual_letter+total_records+')');
  
  // link Daily Spending limit to Budget Sheet
  const new_yy_mmm_str = new_month.getFullYear().toString().slice(-2).concat('-', new_month.toString().slice(4,7));
  spreadsheet.getRange("B1").setFormula('=\'' + new_yy_mmm_str +'-Budget\'!E2').setTextStyle(default_style);
  // fill out column headers
  var row = 2;
  var col = 1;
  curSheet.getRange(row, col, 1, budget_header[0].length).setValues(budget_header);
  // spreadsheet.getRange("A2").setValue(budget_header[0].length).setTextStyle(default_style);
  // loop through days of the month
  row = 3;
  const days = getDaysInMonth(new_month);
  const num_rows_per_day = 2;

  days.forEach(function (day, index) {
    
    const start_row = row;
    for (i = 0; i < num_rows_per_day; i++){
      // date
      curSheet.getRange(row, col_date).setValue(day)
      // day of the week
      curSheet.getRange(row, col_day).setFormula('=Text(WEEKDAY('+ col_date_letter + row + '),"DDDD")')
      
      // fill in dropdowns
      // 'category' dropdown column
      curSheet.getRange(row, col_category).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(spreadsheet.getRange('\'List of options\'!$B$2:$B$100'), true)
      .build());
      
      // 'credit card verification' dropdown column
      curSheet.getRange(row, col_verified).setDataValidation(SpreadsheetApp.newDataValidation()
      .setAllowInvalid(true)
      .requireValueInRange(spreadsheet.getRange('\'List of options\'!$C$2:$C$100'), true)
      .build());

      row += 1;
    }
    const end_row = row;
    
    // fill in dailies
    row -= 1;
    curSheet.getRange(row, col_daily_spending).setFormula('=SUM('+col_spent_letter+start_row+':'+col_spent_letter+end_row+')')
    curSheet.getRange(row, col_residual).setFormula('=B1-'+get_address(row, col_daily_spending))

    row += 1;
  });

  // // add row colors
  // var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  // spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  // conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  // conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  // .setRanges([spreadsheet.getRange('A3:M51')])
  // .whenFormulaSatisfied('=$A$1:$A$808="Sunday"')
  // .setBackground('#FCE8B2')
  // .build());
};

function get_address(row_num, col_num){
  return String.fromCharCode(col_num + 65) + row_num
}


function calculate_residuals() {
  const spreadsheet = SpreadsheetApp.getActive();
  const curSheet = spreadsheet.getActiveSheet();
  const curSheet_name = curSheet.getSheetName();

  if (curSheet_name.includes('Spending')){
    // clear old values
    const col_daily_spending = budget_header[0].indexOf('Daily Spending') + 1;
    const col_residual = budget_header[0].indexOf('Residual') + 1;
    const row_data_start = 3;
    const width_col = col_residual - col_daily_spending + 1;
    curSheet.getRange(row_data_start, col_daily_spending, total_records, width_col).clear({contentsOnly: true, skipFilteredRows: true});

    var date_str = null;
    var prev_date = null;
    var row = header_rows;
    var stored_row = 0;
    for (var i = 0; i < total_records; i++){
      date_str = curSheet.getRange(row, col_date).getValue();
      date = new Date(date_str);
      // check that date is valid
      if (date instanceof Date && !isNaN(date)){
        console.log(date);
        if (prev_date == null){
          prev_date = date;
          stored_row = row;
        } else if ( prev_date.getTime() !== date.getTime() && stored_row < row ){
          const prev_row = row-1;
          var formula_sum = '=SUM('+col_spent_letter+stored_row+':'+col_spent_letter+prev_row+')';
          var formual_residual = '=B1-' + col_daily_spending_letter + prev_row;
          curSheet.getRange(prev_row, col_daily_spending).setFormula(formula_sum);
          curSheet.getRange(prev_row, col_residual).setFormula(formual_residual);
          prev_date = date;
          stored_row = row;
          // curSheet.getRange(stored_row, col_spent, row-1, 0)
        };
      } else if (prev_date){
        // deal with last row
        const prev_row = row-1;
        var formula_sum = '=SUM('+col_spent_letter+stored_row+':'+col_spent_letter+prev_row+')';
        var formual_residual = '=B1-' + col_daily_spending_letter + prev_row;
        curSheet.getRange(prev_row, col_daily_spending).setFormula(formula_sum);
        curSheet.getRange(prev_row, col_residual).setFormula(formual_residual);
        date = null;
        break;
      }
      row += 1;
    }

    // spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  } else {
    throw ("Error: Calculating Residuals can only be performed on 'Spending' sheets.")
  }
};

function conditional_formatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  // clear background colors
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRangeList().setBackground(null);
  // clear any old conditional formatting
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  // set Sunday rows to have gold background
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AB787')])
  .whenFormulaSatisfied('=$A$1:$A$792="Sunday"')
  .setBackground('#FCE8B2')
  .build());
  // set Monday rows to have green background
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AB787')])
  .whenFormulaSatisfied('=$A$1:$A$792="Monday"')
  .setBackground('#B7E1CD')
  .build());
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AB787')])
  .whenFormulaSatisfied('=$A$1:$A$792="Wednesday"')
  .setBackground('#B7E1CD')
  .build());
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AB787')])
  .whenFormulaSatisfied('=$A$1:$A$792="Friday"')
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);

  // conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  // conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  // .setRanges([spreadsheet.getRange('A1:AB787')])
  // .whenFormulaSatisfied('=$A$1:$A$792="Friday"')
  // .setBackground('#B7E1CD')
  // .build());
  // spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};

function clear_background() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveRangeList().setBackground(null);
};

function verified_with_card() {
  // for a card to be verified, it just needs to have the parantheses around it removed
  var spreadsheet = SpreadsheetApp.getActive();
  // apply parantheses removal to all active cells
  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  var to_write_ls = []
  for (var i = 1; i <= range.getNumRows(); i++) {
    var row = []
    for (var j = 1; j <= range.getNumColumns(); j++) {
      var cur_value = range.getCell(i,j).getValue();
      // remove parantheses
      var new_value = cur_value.replace("(","").replace(")","");
      row.push(new_value)
    }
    to_write_ls.push(row)
  }
  range.setValues(to_write_ls)
};