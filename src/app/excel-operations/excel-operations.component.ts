import { Component } from '@angular/core';
import { ExcelOperationsService } from '../excel-operations.service';

@Component({
  selector: 'app-excel-operations',
  templateUrl: './excel-operations.component.html',
  styleUrl: './excel-operations.component.scss'
})
export class ExcelOperationsComponent {
  file1: File | null = null;
  file2: File | null = null;
  mergeColumnName: string = '';
  file: File | null = null;
  splitColumnName: string = '';
  zipFile: Blob | null = null;

  constructor(private excelService: ExcelOperationsService) {}

  onFile1Selected(event: any) {
    this.file1 = event.target.files[0];
  }

  onFile2Selected(event: any) {
    this.file2 = event.target.files[0];
  }

  onFileSelected(event: any) {
    this.file = event.target.files[0];
  }

  mergeFiles() {
    if (this.file1 && this.file2 && this.mergeColumnName) {
      this.excelService.mergeFiles(this.file1, this.file2, this.mergeColumnName)
        .subscribe(response => {
          this.zipFile = response;
        });
    } else {
      alert('Please select two files and provide a column name.');
    }
  }

  splitFile() {
    if (this.file && this.splitColumnName) {
      this.excelService.splitFile(this.file, this.splitColumnName)
        .subscribe(response => {
          this.zipFile = response;
        });
    } else {
      alert('Please select a file and provide a column name.');
    }
  }

  downloadFile() {
    if (this.zipFile) {
      const link = document.createElement('a');
      const url = window.URL.createObjectURL(this.zipFile);
      link.href = url;
      link.download = 'result.zip';
      link.click();
      window.URL.revokeObjectURL(url);
    }
  }
}
