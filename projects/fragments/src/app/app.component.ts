import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  getVisibleRanges() {
    Excel.run(context => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items"]);
      let rangeArr: Array <Excel.Range> = [];
      let hiddenColumns: Array<OfficeExtension.ClientResult<Excel.ColumnProperties[]>> = [];
      let hiddenRows: Array<OfficeExtension.ClientResult<Excel.RowProperties[]>> = [];
      let range: Excel.Range
      return context.sync().then(() => {
        sheets.items.forEach(sheet => {
          sheet.load(["names", "namedSheetViews", "pageLayout", "pivotTables", "name", "visibility", "tables", "context.workbook"])
          range = sheet.getUsedRange();
          range.load(["address", "cellCount", "columnCount", "rowCount", "hidden", "rowHidden", "worksheet", "values"])
          rangeArr.push(range);
          hiddenColumns.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
          hiddenRows.push(range.getRowProperties({rowHidden: true, rowIndex: true}))
        })
        return context.sync().then(() => {
          console.log('hiddenCol', hiddenColumns)
          console.log('hiddenRow', hiddenRows)
          hiddenRows.forEach(el => {
            let visibleRowsArr: Array<any> = [];
            let visibleRows = Object.values(el.value).filter(row => row.rowHidden === false)
            visibleRows.forEach(row => {
              visibleRowsArr.push(row.rowIndex);
            })
            console.log('rowsNotHidden', visibleRowsArr)
          })
          hiddenColumns.forEach(el => {
            let visibleColumnsArr: Array<any> = [];
            let visibleColumns = Object.values(el.value).filter(column => column.columnHidden === false)
            visibleColumns.forEach(column => {
              //@ts-ignore
              visibleColumnsArr.push(column.columnIndex, this.fromNumToChar(column.columnIndex + 1));
            })
            console.log('columnsNotHidden', visibleColumnsArr)
          })
          rangeArr.forEach(el => {
            //console.log("address, column, row, values", el.address, el.columnCount, el.rowCount, el.values)
            //console.log("rowHidden", el.rowHidden)
            console.log("address", el.address)
            //console.log("")
          })
        })
      })
    })
  }

  fromNumToChar(num: number) {
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
