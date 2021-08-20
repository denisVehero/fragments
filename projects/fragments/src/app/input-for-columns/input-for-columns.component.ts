import {Component, Input, OnInit} from '@angular/core';

@Component({
  selector: 'app-input-for-columns',
  templateUrl: './input-for-columns.component.html',
  styleUrls: ['./input-for-columns.component.scss']
})
export class InputForColumnsComponent implements OnInit {

  constructor() {
  }

  @Input() inputValue: {value: string} = {value: ''};

  ngOnInit(): void {
  }

}
