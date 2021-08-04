import {Component, OnInit} from '@angular/core';
import {OfficeEngine} from '../office-engine'

@Component({
  selector: 'app-merger',
  templateUrl: './merger.component.html',
  styleUrls: ['./merger.component.scss']
})
export class MergerComponent implements OnInit {
  sheetArr: string[] = []
  visibleColumnsArr: any[] = [];

  constructor() {
  }

  ngOnInit(): void {
    OfficeEngine.getVisibleSheets().then((arr) => {
      this.sheetArr = arr;
    })
    OfficeEngine.getVisibleColumns('Sheet1').then((arr) => {
      this.visibleColumnsArr = arr;
    })
  }

  getCheckProperties() {
    let checkedSheetsArr: any[] = [];
    let uncheckedColumnsArr: any[] = [];
    let sheetCheckboxes = document.querySelectorAll("input[name=sheet]");
    let columnCheckboxes = document.querySelectorAll("input[name=column]");
    columnCheckboxes.forEach(column => {
      // @ts-ignore
      if (column.checked === true) {
        uncheckedColumnsArr.push(+column.id);
      }
    })
    sheetCheckboxes.forEach(sheet => {
      // @ts-ignore
      if (sheet.checked === true) {
        checkedSheetsArr.push(sheet.id);
      }
    })
    checkedSheetsArr.forEach(sheet => {
      this.getChooseProperties(sheet, uncheckedColumnsArr)
    })
  }

  getChooseProperties(sheet: Excel.Worksheet, columns: Excel.ColumnProperties[]) {
    Excel.run(context => {
      const worksheet = context.workbook.worksheets.getItem(`${sheet}`);
      const arrRows: Array<OfficeExtension.ClientResult<Excel.RowProperties[]>> = [];
      let range: Excel.Range;
      return context.sync().then(() => {
        range = worksheet.getUsedRange();
        range.load(["address"])
        arrRows.push(range.getRowProperties({rowHidden: true, rowIndex: true}))
        const invisibleRowIndexArr: any[] = [];
        return context.sync().then(() => {
          console.log('arrRows', arrRows)
          arrRows.forEach(el => {
            const invisibleRowsArr: Excel.RowProperties[] = el.value.filter(row => row.rowHidden === true)
            invisibleRowsArr.forEach(row => {
              invisibleRowIndexArr.push(row.rowIndex);
            })
            console.log('rowsHidden', invisibleRowIndexArr)
          })
          //this.splitBySquares(invisibleRowIndexArr, invisibleColumnIndexArr);

        })

      })
    })
  }

  splitBySquares(rows: number[], columns: number[]) {
    let i: number;
    let j: number;
    let arrRows = [];
    let row = [];
      for (i = 0, i < rows.length; i++;) {
        row.push()
    }
  }


}
