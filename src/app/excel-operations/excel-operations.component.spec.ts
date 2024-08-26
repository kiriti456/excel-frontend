import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExcelOperationsComponent } from './excel-operations.component';

describe('ExcelOperationsComponent', () => {
  let component: ExcelOperationsComponent;
  let fixture: ComponentFixture<ExcelOperationsComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ExcelOperationsComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ExcelOperationsComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
