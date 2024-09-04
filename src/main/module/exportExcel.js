import FileSaver from "file-saver";
import excel from "exceljs";

function numberToColumnLetter(columnNumber) {
  let temp,
    letter = "";
  while (columnNumber > 0) {
    temp = (columnNumber - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    columnNumber = (columnNumber - temp - 1) / 26;
  }
  return letter;
}

export default async function exportExcel(
  fileName = "file1.xlsx",
  sheets = [
    {
      name: "Sheet1",
      columns: [],
      data: [],
    },
  ],
  dataDropdown = [],
  returnCallback = false,
) {
  let workbook = new excel.Workbook(); // Creating workbook

  sheets.forEach((sheet, index) => {
    let properties = {};

    if (sheet.views) {
      properties.views = sheet.views;
    }

    let worksheet = workbook.addWorksheet(sheet.name, properties);
    worksheet.columns = sheet.columns;
    worksheet.addRows(sheet.data);

    if (sheet.callback && typeof sheet.callback == "function")
      sheet.callback({ worksheet });

    if (!sheet.state) worksheet.state = "visible";
    else worksheet.state = sheet.state;

    // dropdown old version
    if (dataDropdown.length > 0) {
      dataDropdown.forEach((item) => {
        let keyColumnIndex = sheet.columns.findIndex(
          (column) => column.header === item.keyColumn,
        );
        if (keyColumnIndex !== -1) {
          const keyColumnLetter = numberToColumnLetter(keyColumnIndex + 1);
          const startRow = 2;

          Array.from({ length: 50 }).map((_, index) => {
            worksheet.getCell(
              `${keyColumnLetter}${startRow + index}`,
            ).dataValidation = {
              type: "list",
              allowBlank: true,
              formulae: [
                `"${item.arrSelect.map(
                  (subItemSelect) => subItemSelect.name,
                )}"`,
              ],
              showErrorMessage: true,
              errorStyle: "error",
              errorTitle: "Error",
              error: "Value must be in the list",
            };
          });
        }
      });
    }
  });

  const buffer = await workbook.xlsx.writeBuffer();

  //console.log("asdasd")

  if (returnCallback) {
    return { buffer, fileName };
  } else {
    FileSaver.saveAs(new Blob([buffer]), fileName);
  }
}
