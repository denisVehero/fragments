export class OfficeEngine {
  rangeArr: Array<Excel.Worksheet> = [];

  visibleColumnsArr = new Map<Number, Number>();
  getVisibleRanges() {

    Excel.run(context => {
      const sheets = context.workbook.worksheets;
      sheets.load(["items"]);
      const hiddenColumns: Array<OfficeExtension.ClientResult<Excel.ColumnProperties[]>> = [];
      let range: Excel.Range;
      return context.sync().then(() => {
        sheets.items.forEach(sheet => {
          sheet.load(["name"])
          range = sheet.getUsedRange();
          range.load(["address", "values"])
          console.log('range', range)
          this.rangeArr.push(sheet);
          hiddenColumns.push(range.getColumnProperties({columnHidden: true, columnIndex: true}))
        })
        return context.sync().then(() => {
          hiddenColumns.forEach(el => {
            const visibleColumns: Excel.ColumnProperties[] = Object.values(el.value).filter(column => column.columnHidden === false);
            visibleColumns.forEach(column => {
              // @ts-ignore
              this.visibleColumnsArr.set(column.columnIndex, this.fromNumToChar(column.columnIndex + 1));
              /*// @ts-ignore
              console.log(column.columnIndex, this.fromNumToChar(column.columnIndex + 1))*/
            })
          })
          //console.log('values', range.values)
          console.log('visibleColumnsArr', this.visibleColumnsArr)
        })
      })
    })
  }
}
