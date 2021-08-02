import { range } from "rxjs";

export class OfficeEngine {
    static maxCells = 5000;
    /**
     * copies data. Each bound must be the same size
     * @param sourceInd indices of source columns
     * @param sourceSheet name of source sheet
     * @param toInd indices of target columns
     * @param toSheet indices of target sheet
     */
    static copyValues(sourceInd: Bound[], toInd: Bound[]): Promise<any> {
        if (sourceInd.length != toInd.length) {
            throw new Error("columns do not match");
        }
        let task: TruckBounds[] = [];
        for(let i = 0; i < sourceInd.length; i++) {
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
        return Excel.run((ctx) => {
            let sourceRange = ctx.workbook.worksheets.getItem(task[0].source.sheetName);
            let destRange = ctx.workbook.worksheets.getItem(task[0].destination.sheetName);
            let r1: Excel.Range, r2: Excel.Range;
            for (let i = 0; i < sourceInd.length; i++){
                r1 = this.getRange(sourceRange, sourceInd[i]);;
                r2 = this.getRange(destRange, toInd[i]);
                r2.copyFrom(r1);
            }
            console.log("before sync")
            return ctx.sync();
        })
    }
    /**
     * 
     * @param worksheet worksheet, containing range
     * @param adr Bounds of requested range
     * @returns range object
     */
    static getRange(worksheet: Excel.Worksheet, adr: Bound): Excel.Range{
        return worksheet.getRangeByIndexes(adr.row, adr.col, adr.rowCount, adr.colCount);
    }

    static fillWithSomething(rangeAdr: Bound) {
        console.log(rangeAdr)
        return Excel.run((ctx) => {
            let w = ctx.workbook.worksheets.getItem(rangeAdr.sheetName);
            let r = this.getRange(w, rangeAdr);
            let color =  "#" + ("00000" + Math.floor(Math.random() * 16581375).toString(16)).slice(-6);
                r.format.fill.color = color;
            return ctx.sync();
        }).then(() => console.log("adas33"));
    }
}
export class  Bound {
    col: number;
    row: number;
    colCount: number;
    rowCount: number;
    sheetName: string;
    constructor(col: number, row: number,colCount: number, rowCount: number, sheet?: string) {
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

        //if(cols * rows > 1000) throw new Error("too big");

        //prime number problem!!!!!
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
        
        console.log(dividers)
        for (let i = 1; i < dividers.length; i++){
            if (cur.value < dividers[i].value) {
                cur = dividers[i]
            }
        }
         
        //бить на промисы
        // не динамический массив
        for(let i = b.row; i <= b.row + b.colCount - 1; i = cur.rows + i) {
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