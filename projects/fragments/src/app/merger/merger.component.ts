import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-merger',
  templateUrl: './merger.component.html',
  styleUrls: ['./merger.component.scss']
})
export class MergerComponent implements OnInit {

  constructor() { }

  ngOnInit(): void {
  }

  getVisibleRanges() {

    Excel.run(context => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items"]);
      let rangeArr: Array<Excel.Range> = [];
      let hiddenColumns: Array<OfficeExtension.ClientResult<Excel.ColumnProperties[]>> = [];
      let hiddenRows: Array<OfficeExtension.ClientResult<Excel.RowProperties[]>> = [];
      let range: Excel.Range;
      return context.sync().then(() => {
        sheets.items.forEach(sheet => {
          sheet.load(["names", "name", "tables"])
          range = sheet.getUsedRange();
          range.load(["address", "values"])
          rangeArr.push(range);
          hiddenColumns.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
          hiddenRows.push(range.getRowProperties({rowHidden: true, rowIndex: true}))
        })
        const sheetsList = document.getElementById('sheetsList');
        const columnsList = document.getElementById('columnsList')
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
            let visibleColumnsArr: Array<any> = [];
            let visibleColumns = Object.values(el.value).filter(column => column.columnHidden === false)
            visibleColumns.forEach(column => {
              const columnItem = document.createElement('li');
              //@ts-ignore
              columnItem.innerText = this.fromNumToChar(column.columnIndex + 1);
              columnItem.id = `${column.columnIndex}`;
              // @ts-ignore
              columnsList.append(columnItem);
              //visibleColumnsArr.push([column.columnIndex, this.fromNumToChar(column.columnIndex + 1)]);
            })
            // @ts-ignore
            console.log('columnsNotHidden', Object.fromEntries(visibleColumnsArr))
          })
          rangeArr.forEach(el => {
            const sheetItem = document.createElement('li')
            const checkbox = document.createElement('input')
            sheetItem.innerText = el.address;
            checkbox.type = `checkbox`;
            checkbox.id = `${el.address}`;
            // @ts-ignore
            sheetsList.append(sheetItem);
            // @ts-ignore
            sheetsList.append(checkbox);
            //console.log("address", el.address)
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
