const COLUMNS = ["D", "E", "F", "G", "H"];
const START_LINE = 2;
const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const SHEET = SPREADSHEET.getSheets()[0];

function main() {
  const numberOfRows = getNumberOfRows();
  
  for(let i = START_LINE; i <= numberOfRows; i++) {
    Logger.log(`Converting line: ${i}`);
    COLUMNS.forEach((column) => {
      convert(`${column}${i}`);
    });
  }
}

function convert(cell) {
  const range = SHEET.getRange(cell);
  const formula = range.getFormula();

  if (formula != "") {
    Logger.log(`Converting: ${cell}`);

    const link = convertFormulaToLink(formula)

    try {
      const cellImage = SpreadsheetApp
                .newCellImage()
                .setSourceUrl(link)
                .setAltTextTitle(link)
                .build();

      range.setValue(cellImage);
    } catch(e) {
      Logger.log(e);
    }
  }
}

function getNumberOfRows() {
  return SHEET.getDataRange().getValues().length;
}

function convertFormulaToLink(formula) {
  return formula
          .replace("image", "")
          .replace("IMAGE")
          .replace("=(\"", "")
          .replace("\")", "");
}
