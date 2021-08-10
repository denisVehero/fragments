import {Component, OnInit} from '@angular/core';
import {Bound, OfficeEngine} from '../office-engine'
import context = Office.context;

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

  }

  getCheckProperties(): void {
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

    OfficeEngine.createWorksheet(['Sheet100']).then((name) => {
      debugger;
      let sheetName: string = '';

      let startRow: number = 0;
      let startCol: number = 0;
      let rowCount: number = 0;
      let colCount: number = 0;
      let boundName: string;
      let nameCurSheet: string;
      sheetName = name[0];//debugger;

      checkedSheetsArr.forEach(sheet => {

        let arrOfBounds: Bound[] = [];
        let arrOfNewBounds: Bound[] = [];
        OfficeEngine.getInvisibleRows(sheet).then((arr) => {
          console.log('sheet', sheet)
          this.invisibleRowsArr = arr;
          arrOfBounds = this.splitByVisibleBounds(this.invisibleRowsArr, checkedColumnsArr, sheet);
          console.log('rows+ columns', this.invisibleRowsArr, checkedColumnsArr);
          arrOfBounds.forEach(bound => {
            if (bound.sheetName != nameCurSheet && nameCurSheet !== undefined) {
              debugger;
              startRow += rowCount - 1;//debugger;
              startCol = bound.col;
            } else {
              if (bound.col === 0) {//debugger;
                startCol = 0;
                startRow += rowCount;//debugger;
              } else if (startCol === 0 || startCol + colCount < bound.col) {//debugger;
                startCol += colCount;
                startRow = startRow;//debugger;
              } /*else if (startCol === bound.col + startCol) {debugger;
              startCol += colCount;
            }*/
              /*else if () {debugger;
                startRow = startRow;debugger;
              }else {debugger;
                startRow += rowCount;debugger;
              }*/
              /*if (bound.row === 0) {
                startRow = 0;
              } else if (startRow === 0) {debugger;
                startRow += rowCount;
              } else if (bound.row != startRow) {
                startRow += rowCount;
              } else {
                startRow = bound.row;
                startCol = bound.col;
              }*/
            }
            rowCount = bound.rowCount;
            colCount = bound.colCount;
            boundName = sheetName;
            debugger;
            nameCurSheet = bound.sheetName;
            debugger;
            //finalRow = startRow + rowCount + 1;
            arrOfNewBounds.push(new Bound(startCol, startRow, colCount, rowCount, boundName));
            debugger;

          })
          console.log('arrOfNewBounds', arrOfNewBounds)
          OfficeEngine.copyValues(arrOfBounds, arrOfNewBounds).then()
        })

      })
    })
  }

  splitByVisibleBounds(rows: number[], cols: number[], sheet: string): Bound[] {
    let i: number;
    let j: number;
    let b: number;
    let row: Bound[] = [];
    let colCount: number = 1;
    let startRow: number = 0;
    let rowCount: number = 0;
    for (i = 0, i < rows.length + 1; ;) {
      //debugger
      b = i;

      if (i === 0) {
        startRow = 0;
        rowCount = rows[i];
      } else {
        if (rows[i] === 0 && rows[i + 1] === 1) {
          //debugger
          startRow = rows[i + 1] + 1;
          rowCount = rows[i + 1] - startRow - 1;
        } else if (rows[i] === 0) {
          //debugger
          startRow = rows[i] + 1;
          rowCount = rows[i + 1] - startRow - 1;
        } /*else if (i === rows.length + 1) {
        rowCount = rows[i] - rows[i - 1];
      }*/ else {
          //debugger
          rowCount = rows[i] - rows[i - 1] - 1;
        }
      }
      if (i === row.length - 1) {
        rowCount = rows[i] - startRow;
      }
      while (rows[b] + 1 === rows[b + 1] && i !== 0) {
        rowCount += 1;
        b += 1;
      }
      while (rows[b] + 1 === rows[b + 1] && i - 1 === 0) {
        startRow += 1;
        b += 1;
      }
      if (i - 1 === 0) {
        startRow += rowCount + 1;
        rowCount = rows[b + 1] - rows[b] - 1;
      }

      for (j = 0, j < cols.length; ;) {
        let a = j;

        while (cols[a] + 1 === cols[a + 1]) {
          colCount += 1;
          a += 1;
        }

        if (rowCount > 1 && b + 1 > rows.length - 1) {
          rowCount = 1;
        }

        row.push(new Bound(cols[j], startRow, colCount, rowCount, sheet));
        //debugger;
        console.log('jhlkjl', new Bound(cols[j], startRow, colCount, rowCount, sheet))

        if (colCount > 1) {
          j = a + 1;
          colCount = 1;
        } else {
          j += 1;
        }

        if (j >= cols.length) {

          if (rowCount > 1) {
            i = b + 1;
          } else {
            i += 1;
          }

          startRow += rowCount + 1;
          colCount = 1;
          rowCount = 1;
          break;
        }

      }
      if (i > rows.length || b + 1 > rows.length - 1 || startRow > rows[rows.length - 1]) {
        break;
      }
    }
    console.log('row', row)
    return row;
  }

}
