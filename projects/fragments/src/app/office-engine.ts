export class OfficeEngine {

  rangeArr: Array<Excel.Worksheet> = [];
  visibleColumnsArr = new Map<Number, Number>();

/*  static getVisibleColumns(): Promise<Excel.Range> {
    return Excel.run(context => {
      const sheet = context.workbook.worksheets.getItem('Sheet1');
      sheet.load(["items"]);
      const hiddenColumns: Array<OfficeExtension.ClientResult<Excel.ColumnProperties[]>> = [];
      let range: Excel.Range;
      range = sheet.getUsedRange();
      range.load(["address", "values"])
      return context.sync().then(() => {
        console.log('range', range)
        hiddenColumns.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
        hiddenColumns.forEach(el => {
          const visibleColumns: Excel.ColumnProperties[] = Object.values(el.value).filter(column => column.columnHidden === false);
          visibleColumns.forEach(column => {
            // @ts-ignore
            this.visibleColumnsArr.set(column.columnIndex, this.fromNumToChar(column.columnIndex + 1));
            /!*!// @ts-ignore
            console.log(column.columnIndex, this.fromNumToChar(column.columnIndex + 1))*!/
          })
        })
        //console.log('values', range.values)
        console.log('visibleColumnsArr', this.visibleColumnsArr)
        return range;

      })
    })
  }*/
}
