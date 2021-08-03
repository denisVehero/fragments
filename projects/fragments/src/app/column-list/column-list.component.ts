import { ChangeDetectionStrategy,Component, OnInit } from '@angular/core';

@Component({
	selector: 'app-column-list',
	templateUrl: './column-list.component.html',
	styleUrls: ['./column-list.component.scss'],
	changeDetection: ChangeDetectionStrategy.OnPush
})
export class ColumnListComponent implements OnInit {

	constructor() { }
	list: object[] = [];
	ngOnInit(): void {
		for (let i = 0; i < 16000; i+=1){
			this.list.push({index: i, name: "G"});
		}
	}

}
