import {Component, OnInit} from '@angular/core';

@Component({
  selector: 'app-merger',
  templateUrl: './merger.component.html',
  styleUrls: ['./merger.component.scss']
})
export class MergerComponent implements OnInit {
  sheetArr: Array<string> = []

  visibleColumnsArr: Array<any> = [];

  constructor() {
  }

  ngOnInit(): void {
    this.getSheets().then((arr) => {
      this.sheetArr = arr;
    })
    this.getVisibleColumns().then((arr) => {
      this.visibleColumnsArr = arr;
    })
  }

  getVisibleColumns(): Promise <any[]> {
    return Excel.run(context => {
      const sheet = context.workbook.worksheets.getItem('Sheet1');
      sheet.load(["items"]);
      const hiddenColumns: Array<OfficeExtension.ClientResult<Excel.ColumnProperties[]>> = [];
      let range: Excel.Range;
      range = sheet.getUsedRange();
      range.load(["address", "values"]);
      console.log('range', range)
      hiddenColumns.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
      return context.sync().then(() => {
        let visibleArr: Array<any> = [];
        hiddenColumns.forEach(el => {
          const visibleColumns: Excel.ColumnProperties[] = el.value.filter(column => column.columnHidden === false);
          console.log(visibleColumns, '')
          visibleColumns.forEach(column => {
            if (column.columnIndex) {
              visibleArr.push({index: column.columnIndex, value: this.fromNumToChar(column.columnIndex + 1)});
            }
          })
        })
        return visibleArr;
      })
    })
  }

  getSheets(): Promise<Array<string>> {
    return Excel.run(context => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items"]);
      let sheetArr: Array<string> = [];
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

  getCheckProperties() {
    let checkedSheetsArr: Array<any> = [];
    let checkedColumnsArr: Array<any> = [];
    let sheetCheckboxes = document.querySelectorAll("input[name=sheet]");
    let columnCheckboxes = document.querySelectorAll("input[name=column]");
    //console.log('sheetCheckboxes', sheetCheckboxes)
    columnCheckboxes.forEach(column => {
      // @ts-ignore
      if (column.checked === false) {
        checkedColumnsArr.push(column.id);
      }
    })
    sheetCheckboxes.forEach(sheet => {
      // @ts-ignore
      if (sheet.checked === true) {
        checkedSheetsArr.push(sheet.id);
      }
    })
    checkedSheetsArr.forEach(sheet => {
      this.getChooseProperties(sheet, checkedColumnsArr)
    })
  }

  getChooseProperties(sheet: Excel.Worksheet, columns: Excel.ColumnProperties[]) {
    Excel.run(context => {
      const worksheet = context.workbook.worksheets.getItem(`${sheet}`);
      const hiddenRows: Array<OfficeExtension.ClientResult<Excel.RowProperties[]>> = [];
      let range: Excel.Range;

      return context.sync().then(() => {
        range = worksheet.getUsedRange();
        range.load(["address", "values"])
        hiddenRows.push(range.getRowProperties({rowHidden: true, rowIndex: true, address: true, addressLocal: true}))
        const visibleRowsArr: Array<any> = [];
        return context.sync().then(() => {
          console.log('hiddenRows', hiddenRows)
          const getValues = range.values;
          hiddenRows.forEach(el => {
            const visibleRows: Excel.RowProperties[] = Object.values(el.value).filter(row => row.rowHidden === true)
            visibleRows.forEach(row => {
              visibleRowsArr.push(row.rowIndex);
            })
            console.log('rowsNotHidden', visibleRowsArr)
          })
          console.log('getValues', getValues)
          //console.log(columns)
          /*for (let el of visibleRowsArr) {
            Object.values(getValues).forEach(row => {
              if (row === el) {

              }
            })
          }*/

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
