import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import { LinkedList } from 'linked-list-typescript';
import { Bound, OfficeEngine } from '../office-engine';
import { ProgressStatus } from "../progress-statuses";
@Component({
    selector: 'app-decomposer',
    templateUrl: './decomposer.component.html',
    styleUrls: ['./decomposer.component.scss']
})
export class DecomposerComponent implements OnInit {

    constructor(private _cdr: ChangeDetectorRef) { }
	progressStatuses: LinkedList<ProgressStatus> = new LinkedList<ProgressStatus>();
	columns: {index: number, name: string, checked: boolean}[] = [];
	mode: number = 0;
	n: number = 10;
	sheetName: string = '';

    ngOnInit(): void {
		OfficeEngine.setOnSheetActivated( (arr) => {
			return OfficeEngine.getCurrentSheet().then((res) => {
				this.sheetName = res.toString();
				OfficeEngine.getVisibleColumns(this.sheetName).then((ans) => {
					this.columns = ans.map((val) => {val.checked = false; return val;})
					this.columns[0].checked = true;
					this.columns[1].checked = true;
					this.columns[2].checked = true;
					this.columns[3].checked = true;
					this._cdr.detectChanges();
				})
			})
		})
    }
	/**
	 * // TODO: delete all "this" refs with lets
	 */
	async splitButtonClick() {
		console.log(this.n)
		let p = new ProgressStatus(10, 0, "split by " + this.n);
		this.progressStatuses.append(p);
		let checkedCols = this.columns.filter((v) => v.checked)
		let hiddenRows = await OfficeEngine.getInvisibleRows(this.sheetName);
		let sourceRanges: Bound[] = [];
		let destinationRanges: Bound[] = [];
		
		hiddenRows.unshift(-1);
		
		let prev = checkedCols[0];
		let deltaX = 0;
		for(let i = 0; i <= checkedCols.length - 1; i++) {
			//left col: checkedCols[i].index, colcount: checkedCols[i + 1].index - checkedCols[i].index
			if (i ==checkedCols.length - 1 || checkedCols[i + 1].index - checkedCols[i].index > 1){
				let counter = 1;
				let sheetCounter = 0;
				let remaining = this.n;
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
					sourceRanges.push(tmp);
					let tmp2 = new Bound(
						prev.index - deltaX,
						this.n - remaining,
						checkedCols[i].index - prev.index + 1,
						dj,
						sheetCounter * this.n + "..." + ((sheetCounter + 1) * this.n - 1)
					)
					destinationRanges.push(tmp2)
					j += dj;
					remaining = remaining - dj;
					if (remaining == 0) {
						remaining = this.n
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
		let sheets =[];
		for(let  i= 0; i < destinationRanges.length; i++) {
			sheets.push(destinationRanges[i].sheetName);
		}
		p.planed= sourceRanges.length;
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
        let bounds: Bound[] = new Array(10);
        for (let  i = 0; i < 100; i++){
            for (let j = 0; j < 100; j++){
				bounds.push(new Bound(i*10, j*10, 10, 10))
                
            }
        }
		await OfficeEngine.fillWithSomething(bounds);
    }
    copyButtonClick() {
        let b1 = [];
        let b2 = [];
        for (let i = 0; i < 1000; i++){
            b1.push(new Bound(i, 0, 1, 1000, "Sheet1"))
        }
        for (let i = 0; i < 1000; i++){
            b2.push(new Bound(i, 0, 1, 1000, "Sheet2"))
        }
		
        console.log("copy");
		let p = new ProgressStatus(b1.length, 0, "copying");
		this.progressStatuses.append(p);
		OfficeEngine.copyValues(b1, b2, p).then(() => {
			console.log("finished in ")
		});
    }
    test() {
		console.log(this.n);
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
