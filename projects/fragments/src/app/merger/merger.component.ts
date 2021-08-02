import {Component, OnInit} from '@angular/core';

@Component({
  selector: 'app-merger',
  templateUrl: './merger.component.html',
  styleUrls: ['./merger.component.scss']
})
export class MergerComponent implements OnInit {

  rangeArr: Array<Excel.Range> = [];

  visibleColumnsArr = new Map<Number, Number>()

  constructor() {
  }

  ngOnInit(): void {
  }

  getVisibleRanges() {

    Excel.run(context => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items"]);
      let hiddenColumns: Array<OfficeExtension.ClientResult<Excel.ColumnProperties[]>> = [];
      let hiddenRows: Array<OfficeExtension.ClientResult<Excel.RowProperties[]>> = [];
      let range: Excel.Range;
      return context.sync().then(() => {
        sheets.items.forEach(sheet => {
          sheet.load(["names", "name", "tables"])
          range = sheet.getUsedRange();
          range.load(["address", "values"])
          console.log('range', range)
          this.rangeArr.push(range);
          hiddenColumns.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
          hiddenRows.push(range.getRowProperties({rowHidden: true, rowIndex: true}))
        })
        return context.sync().then(() => {
          /*console.log('hiddenCol', hiddenColumns)
          console.log('hiddenRow', hiddenRows)*/
          hiddenRows.forEach(el => {
            let visibleRowsArr: Array<any> = [];
            let visibleRows = Object.values(el.value).filter(row => row.rowHidden === false)
            visibleRows.forEach(row => {
              visibleRowsArr.push(row.rowIndex);
            })
            console.log('rowsNotHidden', visibleRowsArr)
          })
          hiddenColumns.forEach(el => {
            let visibleColumns: Excel.ColumnProperties[] = Object.values(el.value).filter(column => column.columnHidden === false);
            visibleColumns.forEach(column => {
              // @ts-ignore
              this.visibleColumnsArr.set(column.columnIndex, this.fromNumToChar(column.columnIndex + 1));
              /*// @ts-ignore
              console.log(column.columnIndex, this.fromNumToChar(column.columnIndex + 1))*/
            })
            //console.log('columnsNotHidden', hiddenColumns)
          })
          console.log('visibleColumnsArr', this.visibleColumnsArr)
        })
      })
    })
  }

  getCheckProperties() {
    let checkedSheetsArr: Array<any>;
    let checkedColumnsArr: Array<any>;
    let sheetCheckboxes = document.querySelectorAll("input[name=sheet]");
    let columnCheckboxes = document.querySelectorAll("input[name=column]");
    sheetCheckboxes.forEach(el=> {
      /*if (el.checked === true) {

      }*/
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
