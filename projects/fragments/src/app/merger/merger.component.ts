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
    let b: number;
    let arrRows: Bound[][] = [];
    let row: Bound[] = [];
    let colCount: number = 1;
    let startRow: number = 0;
    //let startCol: number = 0;
    let rowCount: number = 0;
    //console.log('row 0', row)
    //console.log('rows, cols', rows, cols)
    for (i = 0, i < rows.length + 1; ;) { debugger;
      b = i;

      if (i === 0) {
        startRow = 0;
        rowCount = rows[i];
      } else {
        if (rows[i] === 0 && rows[i + 1] === 1) {
          startRow = rows[i + 1] + 1;
          rowCount = rows[i + 1] - startRow - 1;
        } else if (rows[i] === 0) {
          startRow = rows[i] + 1;
          rowCount = rows[i + 1] - startRow - 1;
        } /*else if (i === rows.length + 1) {
        rowCount = rows[i] - rows[i - 1];
      }*/ else {
          rowCount = rows[i] - rows[i - 1] - 1;debugger;
        }
      }
      while (rows[b] + 1 === rows[b + 1]) {
        rowCount += 1;debugger;
        b += 1;
      }

      for (j = 0, j < cols.length; ;) {debugger;
        let a = j;

        while (cols[a] + 1 === cols[a + 1]) {
          colCount += 1;
          a += 1;
        }

        if (rowCount > 1 && b + 1 > rows.length - 1) {debugger;
          rowCount = 1;
        }

        row.push(new Bound(cols[j], startRow, colCount, rowCount));debugger;
        console.log('jhlkjl', new Bound(cols[j], startRow, colCount, rowCount))

        if (colCount > 1) {
          j = a + 1;
          colCount = 1;
        } else {
          j += 1;
        }

        if (j > cols.length) {
          arrRows.push(row);

          if(rowCount > 1) {
            i = b + 1;
          } else {
            i += 1;
          }

          startRow += rowCount + 1;
          colCount = 1;
          rowCount = 1;
          break;
        }
        /*{if (rows[i] === 0) {
         if (cols[j] + colCount === cols[j + 1]) {
           colCount += 1;
           row.push(new Bound(cols[j], rows[i] + 1, colCount, rows[i + 1]));
           j += 2;
         } else {
           colCount = 1;
           row.push(new Bound(cols[j], rows[i] + 1, colCount, rows[i + 1] + 1));
           j += 1;
         }
         if (cols[j] === cols[cols.length]) {
           console.log('j break', j)
           arrRows.push(row);
           i++;
           break;
         }
       } else*/

      }
      if (i >= rows.length || b + 1 > rows.length - 1) {
        break;
      }
    }
    console.log('arrRows', arrRows)
    return arrRows;
  }

}
