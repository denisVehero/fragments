import { range } from "rxjs";

export class OfficeEngine {

    /**
     * copies data. Indices arrays must be the same lenght
     * @param sourceInd indices of source columns
     * @param sourceSheet name of source sheet
     * @param toInd indices of target columns
     * @param toSheet indices of target sheet
     */
    static copyValues(sourceInd: Bound[], sourceSheet: string, toInd: Bound[], toSheet: string): Promise<any> {
        if (sourceInd.length != toInd.length) {
            throw new Error("columns count do not match");
        }
        return Excel.run((ctx) => {
            let sourceRange = ctx.workbook.worksheets.getItem(sourceSheet);
            let destRange = ctx.workbook.worksheets.getItem(toSheet);
            let r1: Excel.Range, r2: Excel.Range;
            for (let i = 0; i < sourceInd.length; i++){
                r1 = this.getRange(sourceRange, sourceInd[i]);;
                r2 = this.getRange(destRange, toInd[i]);
                r2.copyFrom(r1);
            }
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
        return Excel.run((ctx) => {
            let w = ctx.workbook.worksheets.getItem(rangeAdr.sheetName);
            let r = this.getRange(w, rangeAdr);
            let p = new Array(rangeAdr.rowCount).fill(0).map((el) => {
                let qq =new Array(rangeAdr.colCount).fill(0).map((el1: any) => Math.floor(Math.random()*100)) 
                return qq
            })
            r.values = p;
            return ctx.sync();
        })
    }
    static test() {
        Excel.run((ctx) => {
            let w = ctx.workbook.worksheets.getFirst();
            let r = this.getRange(w, new Bound(0,0,1,-1))
            r.load(["address"])
            return ctx.sync().then(() => {
                console.log(r.address)
            });
        })
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
        if(sheet) console.log(sheet)
    }
}