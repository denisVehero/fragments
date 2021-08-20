import { ComponentFixture, TestBed } from '@angular/core/testing';

import { InputForColumnsComponent } from './input-for-columns.component';

describe('InputForColumnsComponent', () => {
  let component: InputForColumnsComponent;
  let fixture: ComponentFixture<InputForColumnsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ InputForColumnsComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(InputForColumnsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
