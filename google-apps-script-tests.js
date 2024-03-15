function CreateNewSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.insertSheet(1);
};

function MoveSheet3times() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.moveActiveSheet(3);
  spreadsheet.moveActiveSheet(4);
  spreadsheet.moveActiveSheet(5);
};

function Autofill10Cells() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue('1');
  spreadsheet.getRange('A2').activate();
  spreadsheet.getCurrentCell().setValue('2');
  spreadsheet.getRange('A1:A2').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A1:A10'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A1:A10').activate();
};

function MoveRangeFromAtoC() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  spreadsheet.getRange('A1:A10').moveTo(spreadsheet.getActiveRange());
};

function FillColorRedThenBlue() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  spreadsheet.getActiveRangeList().setBackground('#ff0000')
  .setBackground('#0000ff');
};

function RemoveFillColor() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  spreadsheet.getActiveRangeList().setBackground(null);
};

function BoldThenItalicThenStrikethroughThenRemoveEach() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold')
  .setFontStyle('italic')
  .setFontLine('line-through')
  .setFontWeight(null)
  .setFontStyle(null)
  .setFontLine(null);
};

function ChangeFont3TimesThenBackToDefaultArial() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  spreadsheet.getActiveRangeList().setFontFamily('Arial')
  .setFontFamily('Comic Sans MS')
  .setFontFamily('Lora')
  .setFontFamily(null);
};

function ChangeNumberFormatsThenBackToAutomatic() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0.00')
  .setNumberFormat('0.00%')
  .setNumberFormat('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)')
  .setNumberFormat('M/d/yyyy')
  .setNumberFormat('M/d/yyyy H:mm:ss')
  .setNumberFormat('@')
  .setNumberFormat('General');
};

function ChartInsertThenResizeThenMove() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('C1:C10'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setPosition(2, 2, 27, 2)
  .build();
  sheet.insertChart(chart);
  var charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('C1:C10'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setOption('height', 238)
  .setOption('width', 384)
  .setPosition(8, 4, 41, 10)
  .build();
  sheet.insertChart(chart);
  charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('C1:C10'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setOption('height', 238)
  .setOption('width', 384)
  .setPosition(1, 4, 60, 13)
  .build();
  sheet.insertChart(chart);
};

function InsertChart() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('C1:C10'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setPosition(2, 2, 27, 2)
  .build();
  sheet.insertChart(chart);
};

function UpdateChartSize() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('C1:C10'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setOption('height', 238)
  .setOption('width', 385)
  .setPosition(8, 4, 41, 9)
  .build();
  sheet.insertChart(chart);
};

function UpdateChartPosition() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('C1:C10'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setOption('height', 238)
  .setOption('width', 385)
  .setPosition(1, 5, 32, 14)
  .build();
  sheet.insertChart(chart);
};

function UpdateChartRange() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C5').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asColumnChart()
  .addRange(spreadsheet.getRange('C1:C5'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('height', 238)
  .setOption('width', 385)
  .setPosition(1, 5, 32, 14)
  .build();
  sheet.insertChart(chart);
};

function UpdateChartType() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C5').activate();
  var sheet = spreadsheet.getActiveSheet();
  var charts = sheet.getCharts();
  var chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
  .asLineChart()
  .addRange(spreadsheet.getRange('C1:C5'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', false)
  .setOption('isStacked', 'false')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#1a1a1a')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setOption('hAxis.textStyle.color', '#000000')
  .setOption('vAxes.0.textStyle.color', '#000000')
  .setOption('height', 238)
  .setOption('width', 385)
  .setPosition(1, 5, 32, 14)
  .build();
  sheet.insertChart(chart);
};

function CreateLink() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C15').activate();
  spreadsheet.getRange('C15').activate();
  spreadsheet.getRange('C15').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('TextoDoLinkParaOGoogle')
  .setTextStyle(0, 22, SpreadsheetApp.newTextStyle()
  .setForegroundColor('#1155cc')
  .setUnderline(true)
  .build())
  .build());
};

function AddConditionalFormatting() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C1:C10').activate();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('C1:C10')])
  .whenCellNotEmpty()
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('C1:C10')])
  .whenNumberGreaterThanOrEqualTo(5)
  .setBackground('#B7E1CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('C1:C10')])
  .whenNumberGreaterThanOrEqualTo(5)
  .setBackground('#F4C7C3')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};

function SelectRow1AndAdd1RowAbove() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
};

function SortEverything1CriteriaWithHeader() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  //Select the whole worksheet
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  //The offset is there because there is a header
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 1, ascending: true});
};


function SortEverything1CriteriaWithoutHeader() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().sort(1, true);
};

function SortRange2CriteriaWithHeader() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:C11').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('C11'));
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 1, ascending: true}, {column: 3, ascending: true}]);
};

function CopyCreateSheetAndPaste() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:C11').activate();
  spreadsheet.insertSheet(5);
  spreadsheet.getRange('Sheet25!A1:C11').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};

function CopyCreateSheetAndPasteValues() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:C11').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Sheet26'), true);
  spreadsheet.getRange('A1:C11').activate();
  spreadsheet.getRange('Sheet25!A1:C11').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
};

function FilterAllCells2CriteriaSpecificRecordsOnAAndDateOnB() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).createFilter();
  spreadsheet.getRange('A1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['', 'F', 'G', 'H', 'I', 'J'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(1, criteria);
  spreadsheet.getRange('B1').activate();
  criteria = SpreadsheetApp.newFilterCriteria()
  .whenDateAfter(new Date(1950, 0, 2))
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(2, criteria);
};

function FilterRangeByValues() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:C11').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B3'));
  spreadsheet.getRange('A1:C11').createFilter();
  spreadsheet.getRange('C1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['6', '7', '8', '9', '10'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(3, criteria);
};

function FilterRangeByCriteriaGreaterThan5() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:C11').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('C4'));
  spreadsheet.getRange('A1:C11').createFilter();
  spreadsheet.getRange('C1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .whenNumberGreaterThan(5)
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(3, criteria);
};

function RemoveFilters() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().getFilter().remove();
};

function RemoveFiltersOnly1CellInsideRangeSelected() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B7').activate();
  spreadsheet.getActiveSheet().getFilter().remove();
};

//UNDO is not recordable!
function WriteIn2CellsThenUndo2Times() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E4').activate();
};

//PRINT is not recordable!
function PrintSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E2').activate();
};

//INSERT IMAGE IN or OVER CELLS is not recordable
function InsertImageInCell() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G6').activate();
};

//INSERT IMAGE IN or OVER CELLS is not recordable
function InsertImageOverCells() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G6').activate();
};

function MoveImage2Times() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G6').activate();
  var sheet = spreadsheet.getActiveSheet();
  var images = sheet.getImages();
  var image = images[images.length - 1];
  image.setAnchorCell(spreadsheet.getRange('D2'))
  .setAnchorCellXOffset(71)
  .setAnchorCellYOffset(13);
  images = sheet.getImages();
  image = images[images.length - 1];
  image.setAnchorCell(spreadsheet.getRange('H2'))
  .setAnchorCellXOffset(73)
  .setAnchorCellYOffset(6);
};

function ResizeImage() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G6').activate();
  var sheet = spreadsheet.getActiveSheet();
  var images = sheet.getImages();
  var image = images[images.length - 1];
  image.setAnchorCell(spreadsheet.getRange('I8'))
  .setAnchorCellXOffset(97)
  .setAnchorCellYOffset(5)
  .setHeight(101)
  .setWidth(101);
};

function InsertDataValidationDropDown() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F2').activate();
  spreadsheet.getRange('F2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Option 1', 'Option 2'], true)
  .build());
  spreadsheet.getRange('C2:C11').activate();
  spreadsheet.getRange('F2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(spreadsheet.getRange('Sheet26!$C$2:$C$11'), true)
  .build());
  spreadsheet.getRange('F2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(spreadsheet.getRange('Sheet26!$C$2:$C$11'), true)
  .build());
  spreadsheet.getRange('F2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(spreadsheet.getRange('Sheet26!$C$2:$C$11'), false)
  .build());
  spreadsheet.getRange('F2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(spreadsheet.getRange('Sheet26!$C$2:$C$11'), true)
  .build());
};

function InsertDataValidationDropDown1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('F2').activate();
  spreadsheet.getRange('F2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInList(['Option 1', 'Option 2'], true)
  .build());
  spreadsheet.getRange('C2:C11').activate();
  spreadsheet.getRange('F2').setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireValueInRange(spreadsheet.getRange('Sheet26!$C$2:$C$11'), true)
  .build());
};

//Center alignment, then right alignment, then wrap text, then clip text
function FormatAlignmentAndTextWrapping() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('B:B').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center')
  .setHorizontalAlignment('right')
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
};