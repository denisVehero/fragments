import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import { LinkedList } from 'linked-list-typescript';
import { Bound, OfficeEngine } from '../office-engine';
import { ProgressStatus } from "../progress-statuses";

class Column {
	index: number;
	name: string;
	checked: boolean;
	constructor(index: number, name: string, checked: boolean){
		this.index = index;
		this.name = name;
		this.checked = checked;
	}
}
@Component({
    selector: 'app-decomposer',
    templateUrl: './decomposer.component.html',
    styleUrls: ['./decomposer.component.scss']
})
export class DecomposerComponent implements OnInit {

    constructor(private _cdr: ChangeDetectorRef) { }
	progressStatuses: LinkedList<ProgressStatus> = new LinkedList<ProgressStatus>();
	columns: Column[] = [];
	mode: number = 0;
	maxRows: number = 10;
	sheetName: string = '';
	keys: {index: number, name: string, checked: boolean}[] = [];
	columnStateUpdate() {
		return OfficeEngine.getCurrentSheet().then((res) => {
			this.sheetName = res.toString();
			OfficeEngine.getVisibleColumns(this.sheetName).then((ans) => {
				this.columns = ans.map((val) => {val.checked = false; return val;})
				this.keys = JSON.parse(JSON.stringify(this.columns));
				this.columns[0].checked = true;
				// this.columns[1].checked = true;
				this.columns[10].checked = true;
				this.columns[20].checked = true;
				this._cdr.detectChanges();
			})
		})
	}
    ngOnInit(): void {
		//TODO: lazy load for columns list
		OfficeEngine.setOnSheetActivated( (arr) => {
			return this.columnStateUpdate();
		})
    }
	/**
	 * // TODO: delete all "this" refs with lets
	 */
	async splitButtonClick() {
		console.log(this.maxRows)
		let p = new ProgressStatus(10, 0, "split by " + this.maxRows);
		this.progressStatuses.append(p);
		let checkedCols = this.columns.filter((v) => v.checked)
		let hiddenRows = await OfficeEngine.getInvisibleRows(this.sheetName);
		let sourceRanges: Bound[] = [];
		let destinationRanges: Bound[] = [];
		
		hiddenRows.unshift(-1);
		
		let prev = checkedCols[0];
		let deltaX = prev.index;
		for(let i = 0; i <= checkedCols.length - 1; i++) {
			//left col: checkedCols[i].index, colcount: checkedCols[i + 1].index - checkedCols[i].index
			if (i == checkedCols.length - 1 || checkedCols[i + 1].index - checkedCols[i].index > 1){
				let counter = 1;
				let sheetCounter = 0;
				let remaining = this.maxRows;
				let j = hiddenRows[0] + 1;
				while (j < hiddenRows[hiddenRows.length - 1]){
					let dj = Math.min(remaining, hiddenRows[counter] - j)
					let tmp = new Bound(
						prev.index,
						j,
						checkedCols[i].index - prev.index + 1,
						dj,
						this.sheetName
					);
					
					let tmp2 = new Bound(
						prev.index - deltaX,
						this.maxRows - remaining,
						checkedCols[i].index - prev.index + 1,
						dj,
						sheetCounter * this.maxRows + "..." + ((sheetCounter + 1) * this.maxRows - 1)
					)
					sourceRanges.push(tmp);
					destinationRanges.push(tmp2)
					j += dj;
					remaining = remaining - dj;
					if (remaining == 0) {
						remaining = this.maxRows
						sheetCounter++;
					}
					if (j >= hiddenRows[counter]) {
						j = hiddenRows[counter] +1;
						counter++;
					}
				}	
				prev = checkedCols[i + 1];
				if (i < checkedCols.length - 1) deltaX += checkedCols[i + 1].index - checkedCols[i].index -1;
			}
		}
		let sheets = new Set<string>();
		for(let  i= 0; i < destinationRanges.length; i++) {
			sheets.add(destinationRanges[i].sheetName);
		}
		p.planed= sourceRanges.length;
		console.log(sheets)
		OfficeEngine.createWorksheet(sheets).then((e) => {
			console.log("created: ", e);
			
			OfficeEngine.copyValues(sourceRanges, destinationRanges, p).then(() => {
				if (this.progressStatuses.length > 1) this.progressStatuses.remove(p)
					else this.progressStatuses.removeHead();
			}).then(() => {
				console.log("Split complited")
			})
		})
		
	}

	delete() {
		//only for debug purposes
		Excel.run(async (ctx) => {
			let w = ctx.workbook.worksheets;
			w.load("items")
			await ctx.sync();
			w.items.forEach(el => {
				if (el.name != "Sheet1" && el.name != "Sheet2") {
					el.delete();
				}
			})
			return ctx.sync();
		}).then(() => console.log("deleted"))
	}
	cancelProgress(p: ProgressStatus) {
		if (this.progressStatuses.length==1) this.progressStatuses.removeHead()
			else this.progressStatuses.remove(p);
	}
    async fillWithRandom() {
		//only for debug purposes
        let bounds: Bound[] = new Array(10);
        for (let  i = 0; i < 100; i++){
            for (let j = 0; j < 100; j++){
				bounds.push(new Bound(i*10, j*10, 10, 10))
                
            }
        }
		await OfficeEngine.fillWithSomething(bounds);
    }
    test() {
		//only for testing purposes
		// console.log(this.n);
		// let prevSelect: boolean = true;
		// Excel.run((ctx) => {
		// 	ctx.workbook.worksheets.onSelectionChanged.add((ev:any) => {
		// 		let r = /\d/.test(ev.address)
		// 		console.log(r);
		// 		if (r && !prevSelect) {
		// 			prevSelect = false;
		// 			return this.columnStateUpdate();
		// 		}
		// 		prevSelect = r;
		// 		return Promise.resolve();
		// 	})
		// 	return ctx.sync();
		// })
		OfficeEngine.createWorksheet(new Set("1000...1009")).then((df) => console.log(df))
		// Excel.run((ctx) => {
		// 	let l = ctx.workbook.worksheets.getActiveWorksheet().tables;
		// 	l.load("items");
		// 	return ctx.sync().then(() => {
		// 		let l1 = l.items[0].columns;
		// 		l1.load("items")
		// 		return ctx.sync().then(() => {
		// 			console.log(l1.items.length)
		// 			l1.items[0].load("name");
		// 			return ctx.sync().then(() => {
		// 				console.log(l1.items[0].name)
		// 			})
		// 		})
		// 	});
		// })
    }

	
}
