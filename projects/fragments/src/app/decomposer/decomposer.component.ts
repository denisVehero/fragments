import { Component, OnInit } from '@angular/core';
import { LinkedList } from 'linked-list-typescript';
import { Bound, OfficeEngine } from '../office-engine';
import { ProgressStatus } from "../progress-statuses";
@Component({
    selector: 'app-decomposer',
    templateUrl: './decomposer.component.html',
    styleUrls: ['./decomposer.component.scss']
})
export class DecomposerComponent implements OnInit {

    constructor() { }
	progressStatuses: LinkedList<ProgressStatus> = new LinkedList<ProgressStatus>();
	columns: {index: number, name: string, checked: boolean}[] = [];
	mode: number = 0;
	n: number = 10;
	sheetName: string ='';
    ngOnInit(): void {
		
		OfficeEngine.getCurrentSheet().then((res) => {
			this.sheetName = res;
			OfficeEngine.getVisibleColumns(this.sheetName).then((ans) => {
				this.columns = ans.map((val) => {val.checked = false; return val;})
				this.columns[0].checked = true;
				this.columns[1].checked = true;
				this.columns[2].checked = true;
				this.columns[3].checked = true;
			})
		})
    }

	async split() {
		let checkedCols = this.columns.filter((v) => v.checked)
		let hiddenRows = await OfficeEngine.getInvisibleRows(this.sheetName);
		let sourceRanges: Bound[] = [];
		let destinationRanges: Bound[] = [];
		hiddenRows.push(100);
		hiddenRows.unshift(-1);
		console.log(checkedCols);
		

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
						"s" + sheetCounter
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
		
	}


	cancel(p: ProgressStatus) {
		if (this.progressStatuses.length==1) this.progressStatuses.removeHead()
			else this.progressStatuses.remove(p);
	}
    async fillWithRandom() {
        let bounds: Bound[] = new Array(10);
        for (let  i = 0; i < 10; i++){
            for (let j = 0; j < 10; j++){
                await OfficeEngine.fillWithSomething(new Bound(i*10, j*10, 10, 10));
            }
        }
    }
    copy() {
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
		OfficeEngine.createWorksheet("as").then((n) => {
			console.log(n);
		})
		
		// OfficeEngine.getRangeValues([new Bound(0,0,1,10), new Bound(1,0,1,10)]).then((res) => {
		// 	console.log(res);
		// })

		// OfficeEngine.getInvisibleRows("Sheet1").then((rr) => {
		// 	console.log("res",rr);
		// })
		// OfficeEngine.getVisibleSheets().then((rr) => {
		// 	console.log(rr);
		// })
		console.log(this.sheetName)
    }

	
}
