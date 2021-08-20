import {Component, OnInit} from '@angular/core';
import {Bound, OfficeEngine} from '../office-engine'
import context = Office.context;

@Component({
  selector: 'app-merger',
  templateUrl: './merger.component.html',
  styleUrls: ['./merger.component.scss'],
})

export class MergerComponent implements OnInit {
  sheetArr: { index: number, name: string, checked: boolean }[] = [];
  invisibleRowsArr: number[] = [];
  visibleColumnsArr: { index: number, name: string, checked: boolean }[] = []
  headersRightOrder = {};
  columnString: { value: string } = {value: 'A-AX'};

  constructor() {
  }

  ngOnInit(): void {
    OfficeEngine.getVisibleSheets().then((arr) => {
      this.sheetArr = arr.map((val) => {
        val.checked = false;
        return val;
      });
      this.sheetArr[0].checked = true;
    })

    OfficeEngine.getVisibleColumns('Sheet1').then((arr) => {
      this.visibleColumnsArr = arr.map((val) => {
        val.checked = false;
        return val;
      })
      //console.log('visible', this.visibleColumnsArrRightOrder)
    })

    OfficeEngine.getHeaders('Sheet1').then(arr => {
      this.headersRightOrder = arr;
    })

    debugger;
    /*console.log('AAA', OfficeEngine.fromNumToChar(703))
    console.log('AAA', OfficeEngine.fromNumToChar(704))
    console.log('AAA', OfficeEngine.fromNumToChar(705))*/
  }

  fillWithSomething() {
    OfficeEngine.fillWithSomething([new Bound(0, 0, 50, 3000, 'Sheet3')]).then()
  }

  getTable() {
    OfficeEngine.getTable().then()
  }

  getCheckProperties(): void {
    let t = performance.now();
    let checkedSheetsArr: any[] = [];
    let checkedColumnsArr = OfficeEngine.getChooseColumns(this.columnString);
    /*this.visibleColumnsArr.forEach(column => {
      if (column.checked) {
        checkedColumnsArr.push({index: column.index});
      }
    })*/
    this.sheetArr.map(sheet => {
      if (sheet.checked) {
        // @ts-ignore
        checkedSheetsArr.push(sheet.value);
      }
    })

    OfficeEngine.createWorksheet().then((name) => {
      console.log('name', name)
      let sheetName: string = '';

      let startRow: number = 0;
      let startCol: number = 0;
      let rowCount: number = 0;
      let colCount: number = 0;
      let boundName: string;
      let nameCurSheet: string;
      sheetName = name[0];//debugger;
      let arrPromises: any[] = [];
      checkedSheetsArr.forEach(sheet => {

        //let headers = [];

        let arrHeadersNew: any[] = [];
        /*OfficeEngine.getHeaders('Sheet2').then(arr => {
          Object.entries(this.headersRightOrder).forEach((elHeader) => {
            Object.entries(arr).map((elValue) => {
              if (elHeader[1] === elValue[1]) {
                /!* console.log('elValue', elValue);
                 console.log('elHeader', elHeader)*!/
                let value = elValue;
                // @ts-ignore
                //arrHeadersNew.push({index: arr[elHeader[0]], value: this.headersRightOrder[elValue[1]]})
                //arr[elHeader[0]] = arr[elHeader[0]];
                // @ts-ignore
                arr[elHeader[1]] = elValue;
                // @ts-ignore
                arrHeadersNew.push(elValue);
                //console.log('arr[elHeader[1]] = elValue', arr[elHeader[1]], elValue)
                //console.log('arr[elValue[1]] = value', arr[elValue[1]], value)
                //arr[elValue[0]] = arr[elValue[0]];
                // @ts-ignore
                //arr[elValue[1]] = value;
              }

            })

          })
          console.log('arr', arrHeadersNew)
        })*/

        let arrOfBounds: Bound[] = [];
        let arrOfNewBounds: Bound[] = [];
        arrPromises.push(OfficeEngine.getInvisibleRows(sheet).then((arr): Promise<any> => {
          console.log('sheet', sheet)
          this.invisibleRowsArr = arr;
          arrOfBounds = this.splitByVisibleBounds(this.invisibleRowsArr, checkedColumnsArr, sheet);
          console.log('rows + columns', this.invisibleRowsArr, checkedColumnsArr);
          arrOfBounds.forEach(bound => {
            debugger;
            if (bound.sheetName != nameCurSheet && nameCurSheet !== undefined) {
              debugger;
              startRow += rowCount;
              debugger;
              startCol = bound.col;
            } else {
              console.log('arrOfBounds[0]', arrOfBounds[0])
              if (bound.col === 0) {
                debugger;
                startCol = 0;
                startRow += rowCount;
                debugger;
              } /*else if (startCol === 0 && ) {

              } */ else if (startCol === 0 || startCol + colCount < bound.col) {
                debugger;
                startCol += colCount;
                startRow = startRow;
                debugger;
              }
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
            //debugger;
            nameCurSheet = bound.sheetName;
            //debugger;
            arrOfNewBounds.push(new Bound(startCol, startRow, colCount, rowCount, boundName));
            debugger;

          })
          console.log('arrOfNewBounds', arrOfNewBounds)
          return OfficeEngine.copyValues(arrOfBounds, arrOfNewBounds).then()
        }))

      })
      Promise.all(arrPromises).then(() => {
        console.log('performance', performance.now() - t);
      })
    })
  }

  splitByVisibleBounds(rows: number[], cols: number[], sheet: string): Bound[] {
    let i: number;
    let j: number;
    let row: Bound[] = [];
    let colCount: number = 1;
    let startRow: number = 0;
    let rowCount: number = 0;
    for (i = -1, i < rows.length; ;) {
      //debugger

      if (i === -1) {
        startRow = 0;
        rowCount = rows[0];
        i = -1;
      } else if (i === rows.length - 1) {
        rowCount = rows[i] - rows[i - 1] - 1;
      } else {
        /*if (rows[i] === 0) {
          debugger
          startRow = rows[i] + 1;
          rowCount = rows[i + 1] - startRow - 1;
        } /!*else if (i === rows.length + 1) {
        rowCount = rows[i] - rows[i - 1];
      }*!/ else*/
        {
          //debugger
          rowCount = rows[i + 1] - rows[i] - 1;
        }
      }

      while (rows[i] + 1 === rows[i + 1] && i !== 0) {
        startRow += 1;
        i += 1;
        rowCount = rows[i + 1] - rows[i] - 1;
      }

      if (i >= rows.length || startRow >= rows[rows.length - 1]) {
        break;
      }

      for (j = 0, j < cols.length; ;) {

        let a = j;

        while (cols[a] + 1 === cols[a + 1]) {
          colCount += 1;
          a += 1;
        }

        if (rowCount) {
          row.push(new Bound(cols[j], startRow, colCount, rowCount, sheet));
          //debugger;
        }
        console.log('jhlkjl', new Bound(cols[j], startRow, colCount, rowCount, sheet))

        if (colCount > 1) {
          j = a + 1;
          colCount = 1;
        } else {
          j += 1;
        }

        if (j >= cols.length) {

          i++;
          startRow += rowCount + 1;
          colCount = 1;
          rowCount = 1;
          break;
        }

      }

    }
    console.log('row', row)
    return row;
  }

}
