import { Component, Input, OnInit, Output } from '@angular/core';
import { ProgressStatus } from "../progress-statuses";
import { EventEmitter } from '@angular/core';
@Component({
  selector: 'app-progress-bar',
  templateUrl: './progress-bar.component.html',
  styleUrls: ['./progress-bar.component.scss']
})
export class ProgressBarComponent implements OnInit {
	@Input() progress: ProgressStatus | undefined;

	@Output() cancel: EventEmitter<ProgressStatus> = new EventEmitter();
  	constructor() { }

  	ngOnInit(): void {
		
  	}
	cancelClick() {
		this.cancel.emit(this.progress);
	}
}
