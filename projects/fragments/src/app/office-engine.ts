import {range} from "rxjs";
import {ProgressStatus} from "./progress-statuses";

export class OfficeEngine {

  static maxCells = 5000;

  /**
   * copies data. Each bound must be the same size
   * @param sourceInd indices of source columns
   * @param sourceSheet name of source sheet
   * @param toInd indices of target columns
   * @param toSheet indices of target sheet
   */
  static copyValues(sourceInd: Bound[], toInd: Bound[], progress?: ProgressStatus): Promise<any> {

    if (sourceInd.length != toInd.length) {
      throw new Error("columns do not match");
    }
    let task: TruckBounds[] = [];
    for (let i = 0; i < sourceInd.length; i++) {
      if (sourceInd[i].colCount * sourceInd[i].rowCount != toInd[i].colCount * toInd[i].rowCount) {
        console.error(sourceInd[i], toInd[i]);
        throw new Error("bound sizes do not match");
      }
      if (sourceInd[i].colCount * sourceInd[i].rowCount > this.maxCells) {
        let tmp1 = Bound.splitBound(sourceInd[i], this.maxCells, this.maxCells);
        let tmp2 = Bound.splitBound(toInd[i], this.maxCells, this.maxCells);
        for (let j = 0; j < tmp1.length; j++) {
          task.push(new TruckBounds(tmp1[i], tmp2[i]));
        }
      } else {
        task.push(new TruckBounds(sourceInd[i], toInd[i]))
      }
    }
    return Excel.run(async (ctx) => {

      let r1: Excel.Range, r2: Excel.Range;
      let counter = 0;
      for (let i = 0; i < task.length; i++) {
        r1 = this.getRange(ctx.workbook, task[i].source);
        ;
        r2 = this.getRange(ctx.workbook, task[i].destination);
        counter += task[i].source.colCount * task[i].source.rowCount;
        r2.copyFrom(r1);
        if (progress) progress.complited++;
        if (counter > this.maxCells) {
          await ctx.sync();
          counter = 0;
        }
      }
      await ctx.sync();
    })
  }


  /**
   *
   * @param worksheet worksheet, containing range
   * @param adr Bounds of requested range
   * @returns range object
   */
  static getRange(workbook: Excel.Workbook, adr: Bound): Excel.Range {
    return workbook.worksheets.getItem(adr.sheetName).getRangeByIndexes(adr.row, adr.col, adr.rowCount, adr.colCount);
  }

  static fillWithSomething(rangeAdr: Bound) {
    console.log(rangeAdr)
    return Excel.run((ctx) => {
      let w = ctx.workbook.worksheets.getItem(rangeAdr.sheetName);
      let r = this.getRange(ctx.workbook, rangeAdr);
      let color = "#" + ("00000" + Math.floor(Math.random() * 16581375).toString(16)).slice(-6);
      r.format.fill.color = color;
      return ctx.sync();
    }).then(() => console.log("adas33"));
  }

  static createWorksheet(workSheetName: String) {
    return Excel.run(async (ctx) => {
      let t = ctx.workbook.worksheets.getItemOrNullObject(String(workSheetName)), i = 0;
      await ctx.sync();
      while (!t.isNullObject) {
        i++;
        t = ctx.workbook.worksheets.getItemOrNullObject(String(workSheetName + "(" + i.toString() + ")"));
        await ctx.sync();
      }
      if (i != 0) ctx.workbook.worksheets.add(String(workSheetName + "(" + i.toString() + ")"))
      else ctx.workbook.worksheets.add(String(workSheetName))
      return ctx.sync()
    })
  }

  static getRangeValues(ranges: Bound[], progress?: ProgressStatus): Promise<any[][]> {
    let task: Bound[] = [];
    for (let i = 0; i < ranges.length; i++) {
      if (ranges[i].colCount * ranges[i].rowCount > this.maxCells) {
        let tmp = Bound.splitBound(ranges[i], this.maxCells, this.maxCells);
        for (let j = 0; j < tmp.length; j++) {
          task.push(tmp[j]);
        }
      } else {
        task.push(ranges[i])
      }
    }
    return Excel.run(async (ctx) => {
      let r1: Excel.Range;
      let counter = 0;
      let res: any[] = [];
      for (let i = 0; i < task.length; i++) {
        r1 = this.getRange(ctx.workbook, task[i]);
        counter += task[i].colCount * task[i].rowCount;
        r1.load("values");
        res.push(r1)
        if (progress) progress.complited++;
        if (counter > this.maxCells) {
          await ctx.sync();
          counter = 0;
        }
      }
      return ctx.sync().then(() => {
        return res.map(r => r.values);
      })
    })
  }

  static getVisibleColumns(sheet: string): Promise<any[]> {
    return Excel.run(context => {
      const worksheet = context.workbook.worksheets.getItem(sheet);
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

  static getVisibleRows(sheet: string): Promise<any[]> {
    return Excel.run(context => {
      const worksheet = context.workbook.worksheets.getItem(sheet);
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

export class Bound {
  col: number;
  row: number;
  colCount: number;
  rowCount: number;
  sheetName: string;

  constructor(col: number, row: number, colCount: number, rowCount: number, sheet?: string) {
    this.col = col;
    this.row = row;
    this.rowCount = rowCount;
    this.colCount = colCount;
    this.sheetName = sheet ? sheet : "Sheet1";
  }

  static splitBound(b: Bound, maxCells: number, maxCols: number): Bound[] {
    //exception
    let result: Bound[] = [];
    let cols = b.colCount;
    let rows = b.rowCount;

    //prime number problem
    let dividers = [];
    let maxPosible = Math.floor(maxCells / 2) + 1;
    for (let i = 1; i < maxPosible; i++) {
      if (maxCells % i == 0) dividers.push(i);
    }
    dividers.push(maxCells);

    let cur;
    dividers = dividers.map((el) => {
      el = Math.floor(maxCells / Math.min(maxCols, Math.floor(maxCells / el)));
      let divider2 = Math.floor(maxCells / el);
      let x = cols / divider2;
      let y = rows / el;
      let p = (Math.floor(x) * Math.floor(y)) * el * divider2;
      return {value: p, rows: el, cols: divider2};
    });

    cur = dividers[0];
    for (let i = 1; i < dividers.length; i++) {
      if (cur.value < dividers[i].value) {
        cur = dividers[i]
      }
    }

    for (let i = b.row; i <= b.row + b.colCount - 1; i = cur.rows + i) {
      for (let j = b.col; j <= b.col + b.colCount - 1; j = cur.cols + j) {
        result.push(new Bound(j, i, Math.min(b.col + b.colCount - j, cur.cols), Math.min(b.row + b.rowCount - i, cur.rows)));
      }
    }
    return result;
  }


}

export class TruckBounds {
  source: Bound;
  destination: Bound;

  constructor(source: Bound, destination: Bound) {
    this.source = source;
    this.destination = destination
  }
}
