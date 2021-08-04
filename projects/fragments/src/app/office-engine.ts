export class OfficeEngine {

  static getVisibleColumns(sheet: Excel.Worksheet): Promise<any[]> {
    return Excel.run(context => {
      const worksheet = context.workbook.worksheets.getItem(`${sheet}`);
      worksheet.load(["items"]);
      const arrColumns: Array<OfficeExtension.ClientResult<Excel.ColumnProperties[]>> = [];
      let range: Excel.Range;
      range = worksheet.getUsedRange();
      range.load(["address"]);
      console.log('range', range)
      arrColumns.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
      return context.sync().then(() => {
        let visibleArr: any[] = [];
        arrColumns.forEach(el => {
          const visibleColumns: Excel.ColumnProperties[] = el.value.filter(column => column.columnHidden === false);
          visibleColumns.forEach(column => {
            if (column.columnIndex != undefined) {
              visibleArr.push({index: column.columnIndex, value: this.fromNumToChar(column.columnIndex + 1)});
            }
          })
        })
        return visibleArr;
      })
    })
  }

  static getVisibleRows(sheet: Excel.Worksheet): Promise<any[]> {
    return Excel.run(context => {
      const worksheet = context.workbook.worksheets.getItem(`${sheet}`);
      worksheet.load(["items"]);
      const arrRows: Array<OfficeExtension.ClientResult<Excel.RowProperties[]>> = [];
      let range: Excel.Range;
      range = worksheet.getUsedRange();
      range.load(["address"]);
      console.log('range', range)
      arrRows.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
      return context.sync().then(() => {
        let visibleArr: any[] = [];
        arrRows.forEach(el => {
          const visibleRows: Excel.RowProperties[] = el.value.filter(row => row.rowHidden === false);
          visibleRows.forEach(row => {
              visibleArr.push(row.rowIndex);
          })
        })
        return visibleArr;
      })
    })
  }

  static getVisibleSheets(): Promise<Array<string>> {
    return Excel.run(context => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items"]);
      let sheetArr: string[] = [];
      return context.sync().then(() => {
        sheets.items.forEach(sheet => {
          sheet.load(["name", "visibility"])
        })
        return context.sync().then(() => {
          sheets.items.forEach(sheet => {
            if (sheet.visibility === Excel.SheetVisibility.visible) {
              sheetArr.push(sheet.name);
            }
          })
          return sheetArr;
        })
      })
    })
  }

  static fromNumToChar(num: number) {
    let letterAddress;
    let secondLetter, firstLetter: string;
    if (num > 26) {
      if (num % 26) {
        firstLetter = String.fromCharCode(64 + (num - (num % 26)) / 26);
        secondLetter = String.fromCharCode(64 + (num % 26));
      } else {
        firstLetter = String.fromCharCode(64 + (num - (num % 26)) / 26 - 1);
        secondLetter = String.fromCharCode(64 + (num % 26) + 26);
      }
      letterAddress = firstLetter + secondLetter;
    } else {
      letterAddress = String.fromCharCode(64 + num);
    }
    return letterAddress;
  }
}
