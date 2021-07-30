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
        for (let i = 0; i < 10; i++){
            b1.push(new Bound(i, 0, 1, -1))
        }
        OfficeEngine.test();
    }
}
