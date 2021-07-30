import { ComponentFixture, TestBed } from '@angular/core/testing';

import { DecomposerComponent } from './decomposer.component';

describe('DecomposerComponent', () => {
  let component: DecomposerComponent;
  let fixture: ComponentFixture<DecomposerComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ DecomposerComponent ]
    })
    .compileComponents();
  });

  beforeEach(() => {
    fixture = TestBed.createComponent(DecomposerComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
