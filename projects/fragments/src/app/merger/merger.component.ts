import {Component, OnInit} from '@angular/core';
import {Bound, OfficeEngine} from '../office-engine'

@Component({
  selector: 'app-merger',
  templateUrl: './merger.component.html',
  styleUrls: ['./merger.component.scss']
})
export class MergerComponent implements OnInit {
  sheetArr: string[] = [];
  visibleColumnsArr: any[] = [];
  invisibleRowsArr: number[] = [];

  constructor() {
  }

  ngOnInit(): void {
    OfficeEngine.getVisibleSheets().then((arr) => {
      this.sheetArr = arr;
    })

    OfficeEngine.getVisibleColumns('Sheet1').then((arr) => {
      this.visibleColumnsArr = arr;
    })
    OfficeEngine.getInvisibleRows('Sheet1').then((arr) => {
      console.log('inv', arr)
      this.invisibleRowsArr = arr;
    })
  }

  getCheckProperties() {
    let checkedSheetsArr: any[] = [];
    let checkedColumnsArr: any[] = [];
    let sheetCheckboxes = document.querySelectorAll("input[name=sheet]");
    let columnCheckboxes = document.querySelectorAll("input[name=column]");
    columnCheckboxes.forEach(column => {
      // @ts-ignore
      if (column.checked === true) {
        checkedColumnsArr.push(+column.id);
      }
    })
    sheetCheckboxes.forEach(sheet => {
      // @ts-ignore
      if (sheet.checked === true) {
        checkedSheetsArr.push(sheet.id);
      }
    })
    checkedSheetsArr.forEach(sheet => {
      //console.log('sheet', sheet)
      /*OfficeEngine.getInvisibleRows(sheet).then((arr) => {
        this.invisibleRowsArr = arr;
      })*/
      this.splitBySquares(this.invisibleRowsArr, checkedColumnsArr)
      console.log('rows+ columns', this.invisibleRowsArr, checkedColumnsArr)
    })
  }

  splitBySquares(rows: number[], cols: number[]): Bound[][] {
    let i: number;
    let j: number;
    let arrRows: Bound[][] = [];
    let row: Bound[] = [];
    let colCount = 1;

    for (i = 0, i < rows.length; ;) {
      if (rows[i] === rows[rows.length - 1]) {
        break;
      }
      for (j = 0, j < cols.length; ;) {
        console.log('row', row)
        if (rows[i] === 0) {
          if (cols[j] + colCount === cols[j + 1]) {
            colCount += 1;
            row.push(new Bound(cols[j], rows[i] + 1, colCount, rows[i + 1]));

            j += 2;
          } else {
            colCount = 1;
            row.push(new Bound(cols[j], rows[i] + 1, colCount, rows[i + 1] + 1));
            j += 1;
          }
        } else {
          if (cols[j] + colCount === cols[j + 1]) {
            colCount += 1;
            row.push(new Bound(cols[j], 0, colCount, rows[i + 1] - rows[i]));
            console.log('jhlkjl', new Bound(cols[j], 0, colCount, rows[i + 1] - rows[i]))
            j += 2;
          } else {
            colCount = 1;
            row.push(new Bound(cols[j], 0, colCount, rows[i + 1] - rows[i]));
            j += 1;
          }

        }

        if (cols[j] === cols[cols.length]) {
          arrRows.push(row);
          i++;
          break;
        }
      }
    }
    console.log('arrRows', arrRows)
    return arrRows;
  }

}
