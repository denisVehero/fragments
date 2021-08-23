import { Component, OnInit, ChangeDetectorRef } from '@angular/core';
import { LinkedList } from 'linked-list-typescript';
import { Bound, OfficeEngine } from '../office-engine';
import { ProgressStatus } from "../progress-statuses";
import { MatSnackBar } from '@angular/material/snack-bar';
import * as XLSX from 'ts-xlsx';
class Column {
	index: number;
	name: string;
	checked: boolean;
	constructor(index: number, name: string, checked: boolean) {
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

	constructor(private _cdr: ChangeDetectorRef, private _snackBar: MatSnackBar) { }
	progressStatuses: LinkedList<ProgressStatus> = new LinkedList<ProgressStatus>();
	columns: Column[] = [];
	mode: number = 0;
	maxRows: number = 1000;
	sheetName: string = '';
	keys: Column[] = [];
	valid: boolean = true;

	showErrorAlert(message: string) {
		this._snackBar.open(message, "Ok")
	}
	columnStateUpdate() {
		return OfficeEngine.getCurrentSheet().then((res) => {
			this.sheetName = res.toString();
			OfficeEngine.getVisibleColumns(this.sheetName).then((ans) => {
				this.columns = ans.map((val) => { val.checked = false; return val; })
				console.log("got ", this.columns, " columns")
				this.keys = JSON.parse(JSON.stringify(this.columns));
				let n = Math.min(50, this.columns.length)
				for (let i = 0; i < n; i++) {
					this.columns[i].checked = true;
				}
				this.keys[0].checked = true;
				this._cdr.detectChanges();
			})
		})
	}
	ngOnInit(): void {
		//TODO: lazy load for columns list
		OfficeEngine.setOnSheetActivated((arr) => {
			return this.columnStateUpdate();
		})

		Excel.run(async (ctx) => {
			let tables = ctx.workbook.worksheets.getActiveWorksheet().tables;
			tables.load("items");
			await ctx.sync();
			let t = tables.items[0];
			t.clearFilters();
			t.columns.load("items");
			await ctx.sync();
			let c = t.columns.items[0];
			c.filter.apply({
				filterOn: Excel.FilterOn.custom,
				criterion1: "<10"
			})
			return ctx.sync();
		})


	}
	test() {
		
		Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, (res) => {
			console.log(res);
			if (res.status == Office.AsyncResultStatus.Succeeded) {
				res.value.getSliceAsync(0, (res) => {
					console.log(res)
					let workbook = XLSX.read(res.value.data, {
						bookFiles: true,
						// sheetRows: 2,
						// cellStyles: true,
						type: "array"
					});
					console.log(workbook)
					/* @ts-ignore */
					let styles = String.fromCharCode.apply(null, workbook.files["xl/styles.xml"]._data);
					let p = new DOMParser();
					let xml = p.parseFromString(styles, "text/xml");
					let pw = xml.getElementsByName("Table Style 1")[0];
					console.log(pw);
				});
			}
			res.value.closeAsync();
		});
	}
	test2() {
		Excel.run(async (ctx) => {
			let tables = ctx.workbook.worksheets.getActiveWorksheet().tables;
			tables.load("items");
			await ctx.sync();
			let t = tables.items[0];
			t.columns.load("items");
			t.autoFilter.load("criteria")
			await ctx.sync();
			console.log("autoFilter: ", t.autoFilter.criteria)
			let c = t.columns.items[0];
			c.filter.load("criteria")
			await ctx.sync();
			console.log("filter: ", c.filter.criteria)
			return ctx.sync();
		})

	}
	/**
	 *  TODO: delete all "this" refs with lets
	 */
	async splitButtonClick() {
		let t0 = performance.now();
		let splitters = [
			function* (n: any): Generator<number, number, number> {
				while (true) {
					yield n;
				}
			},
			function* (arr: any): Generator<number, number, number> {
				while (true) {
					for (let i = 0; i < arr.length; i++) {
						if (yield arr[i]) break;
					}
				}
			}
		], params: any = this.maxRows;
		if (this.mode == 1) {
			await this.getKeyIntervals().then((intervals) => {
				if (intervals.length == 0) {
					return;
				}
				params = intervals;
			})
		}
		console.log("asdas")
		return this.split(splitters[this.mode](params)).then(() => {
			console.log(performance.now() - t0)
		});
	}
	async getKeyIntervals(): Promise<number[]> {
		let task = [];
		let usedRowCount: number;
		usedRowCount = await OfficeEngine.getUsedRowCount()
		for (let i = 0; i < this.keys.length; i++) {
			if (this.keys[i].checked) task.push(new Bound(this.keys[i].index, 0, 1, usedRowCount, this.sheetName))
		}
		if (task.length < 1) {
			this.showErrorAlert("no keys set")
			return Promise.resolve([])
		}
		return OfficeEngine.getRangesValues(task).then((res: any[][][]): number[] => {
			let ans: number[] = [];
			let c = 0;

			for (let i = 1; i < res.length; i++) {
				for (let j = 0; j < res[i].length; j++) {
					res[0][j][0] += res[i][j][0].toString();
				}
			}
			for (let i = 1; i < res[0].length; i++) {
				if (res[0][i - 1][0] != res[0][c][0]) {
					ans.push(i - c);
					c = i;
				}
			}
			if (res[0][0].length - 1 - c > 0) ans.push(res[0][0].length - 1 - c)
			console.log(ans)
			return ans;
		})

	}
	async split(splitter: Generator<number, number, number>) {
		let checkedCols = this.columns.filter((v) => v.checked)
		if (checkedCols.length < 1) {
			this.showErrorAlert("Not columns checked");
			return;
		}

		let hiddenRows = await OfficeEngine.getInvisibleRows(this.sheetName);
		let sourceRanges: Bound[] = [];
		let destinationRanges: Bound[] = [];
		hiddenRows.unshift(-1);

		let prev = checkedCols[0];
		let deltaX = prev.index;
		let work = 0;
		splitter.next().value;
		for (let i = 0; i < checkedCols.length; i++) {
			//left col: checkedCols[i].index, colcount: checkedCols[i + 1].index - checkedCols[i].index
			// test one column
			if (i == checkedCols.length - 1 || checkedCols[i + 1].index - checkedCols[i].index > 1) {
				let counter = 1;
				let sheetCounter = 0;
				let remaining = splitter.next().value;
				let sheetRows = remaining;
				let j = hiddenRows[0] + 1;
				while (j < hiddenRows[hiddenRows.length - 1]) {
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
					work += tmp.colCount * tmp.rowCount;
					sourceRanges.push(tmp);
					destinationRanges.push(tmp2)
					j += dj;
					remaining = remaining - dj;
					if (remaining == 0) {
						remaining = splitter.next().value;
						sheetRows = remaining;
						sheetCounter += sheetRows;
					}
					if (j >= hiddenRows[counter]) {
						j = hiddenRows[counter] + 1;
						counter++;
					}
				}
				prev = checkedCols[i + 1];
				if (i < checkedCols.length - 1) deltaX += checkedCols[i + 1].index - checkedCols[i].index - 1;
			}
		}
		let p = new ProgressStatus(work, 0, "split");
		this.progressStatuses.append(p);
		let sheets = new Set<string>();
		for (let i = 0; i < destinationRanges.length; i++) {
			sheets.add(destinationRanges[i].sheetName);
		}

		return OfficeEngine.createWorksheet(sheets).then((e) => {
			console.log("created: ", e);

			return OfficeEngine.copyValues(sourceRanges, destinationRanges, p).then(() => {
				if (this.progressStatuses.length > 1) this.progressStatuses.remove(p)
				else this.progressStatuses.removeHead();
			}).then(() => {
				console.log("Split complited");
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
				if (el.name != "Sheet1" && el.name != "Sheet2" && el.name != "Sheet3" && el.name != "Sheet4") {
					el.delete();
				}
			})
			return ctx.sync();
		}).then(() => console.log("deleted"))
	}
	cancelProgress(p: ProgressStatus) {
		if (this.progressStatuses.length == 1) this.progressStatuses.removeHead()
		else this.progressStatuses.remove(p);
	}
	async fillWithRandom() {
		//only for debug purposes
		let bounds: Bound[] = new Array(10);
		let cols = 1;
		let rows = 5000;
		for (let i = 0; i < 50; i++) {
			for (let j = 0; j < 100; j++) {
				bounds.push(new Bound(i * cols, j * rows, cols, rows))
			}
		}
		await OfficeEngine.fillWithSomething(bounds);
	}




	ff(n: Generator<number, string, number>) {
		for (let i = 0; i < 10; i++) {
			let c = n.next();
			console.log(c.value);
			if (c.value == 50) n.next(0)
		}
	}
}
