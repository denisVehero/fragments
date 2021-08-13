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
	mode: number = 1;
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
				this.columns[1].checked = true;
				this.columns[2].checked = true;
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
		let arr: any = [10, 5];
		let splitters = [
			function* (n: number): Generator<number, number, number> {
				while (true) {
					yield n;
				}
			},
			function* (arr: number[]): Generator<number, number, number> {
				while (true) {
					for(let i = 0; i < arr.length; i++) {
						if(yield arr[i]) break;
					}
				}
			}
		], params: any[] = [
			this.mode,
			arr
		]
		this.split(splitters[this.mode](params[this.mode]));
		
	}

	async split(splitter: Generator<number, number, number>) {
		let p = new ProgressStatus(10, 0, "split");
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
			// test one column
			if (i == checkedCols.length - 1 || checkedCols[i + 1].index - checkedCols[i].index > 1){
				let counter = 1;
				let sheetCounter = 0;
				let remaining = splitter.next().value;
				let sheetRows = remaining;
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
						sheetRows - remaining,
						checkedCols[i].index - prev.index + 1,
						dj,
						sheetCounter + ".." + (sheetCounter + sheetRows - 1)
					)
					sourceRanges.push(tmp);
					destinationRanges.push(tmp2)
					j += dj;
					remaining = remaining - dj;
					if (remaining == 0) {
						remaining = splitter.next().value;
						sheetRows = remaining;
						sheetCounter+= sheetRows;
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
		function* n(n: number): Generator<number, string, number> {
			let i = 0;
			while(true){
				i+=n;
				let c = (yield i);
				if (c || c==0) i = c - n;
			}
		}

		this.ff(n(10))
    }

	ff(n: Generator<number, string, number>) {
		for(let i = 0; i < 10; i++) {
			let c = n.next();
			console.log(c.value);
			if (c.value==50) n.next(0)
		}
	}
}
