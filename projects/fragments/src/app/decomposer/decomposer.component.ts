import { Component, OnInit } from '@angular/core';
import { Bound, OfficeEngine } from '../office-engine';
@Component({
    selector: 'app-decomposer',
    templateUrl: './decomposer.component.html',
    styleUrls: ['./decomposer.component.scss']
})
export class DecomposerComponent implements OnInit {

    constructor() { }

    ngOnInit(): void {

    }

    fillWithRandom() {
        OfficeEngine.fillWithSomething(new Bound(0,0,100,100))
    }
    copy() {
        let b1 = [];
        let b2 =[];
        for (let i = 0; i < 1000; i++){
            b1.push(new Bound(i, 0, 1, 1000, "Sheet1"))
        }
        for (let i = 0; i < 1000; i++){
            b2.push(new Bound(i, 0, 1, 1000, "Sheet2"))
        }
        OfficeEngine.copyValues(b1, b2).then(() => {
            console.log("finished")
        });
    }
    test() {
        console.log(Bound.splitBound(new Bound(0, 0, 1000, 1000), 25, 5))
    }
}
