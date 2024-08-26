import { TestBed } from '@angular/core/testing';

import { ExcelOperationsService } from './excel-operations.service';

describe('ExcelOperationsService', () => {
  let service: ExcelOperationsService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(ExcelOperationsService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
