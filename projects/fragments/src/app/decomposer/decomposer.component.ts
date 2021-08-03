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
	progressStatuses: LinkedList<ProgressStatus> = new LinkedList<ProgressStatus>()
    ngOnInit(): void {
		
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
		let p = new ProgressStatus(0,0, "sometr");
		this.progressStatuses.append(p);
        OfficeEngine.copyValues(b1, b2, p).then(() => {
            console.log("finished")
        });
    }
    test() {
    }
}
