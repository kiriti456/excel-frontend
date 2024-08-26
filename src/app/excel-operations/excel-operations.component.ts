import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';

@Component({
  selector: 'app-excel-operations',
  templateUrl: './excel-operations.component.html',
  styleUrls: ['./excel-operations.component.scss']
})
export class ExcelOperationsComponent {
  file1: File | null = null;
  file2: File | null = null;
  mergeColumnName: string = '';
  file: File | null = null;
  splitColumnName: string = '';

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
      const file1Reader = new FileReader();
      const file2Reader = new FileReader();

      file1Reader.onload = (e: any) => {
        const workbook1 = XLSX.read(e.target.result, { type: 'array' });
        const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
        const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });

        file2Reader.onload = (e: any) => {
          const workbook2 = XLSX.read(e.target.result, { type: 'array' });
          const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];
          const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

          // Merge data
          const { unionData, intersectionData } = this.mergeData(data1, data2, this.mergeColumnName);

          // Create and download the resulting files
          const zip = new JSZip();
          zip.file('union.xlsx', this.createExcelBlob(unionData, 'Union'));
          zip.file('intersection.xlsx', this.createExcelBlob(intersectionData, 'Intersection'));

          zip.generateAsync({ type: 'blob' }).then(content => {
            saveAs(content, 'merged_intersected_files.zip');
          });
        };

        if (this.file2) {
          file2Reader.readAsArrayBuffer(this.file2);
        }
      };

      if (this.file1) {
        file1Reader.readAsArrayBuffer(this.file1);
      }
    } else {
      alert('Please select two files and provide a column name.');
    }
  }

  splitFile() {
    if (this.file && this.splitColumnName) {
      const fileReader = new FileReader();

      fileReader.onload = (e: any) => {
        const workbook = XLSX.read(e.target.result, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Split data
        const groupedData = this.splitData(data, this.splitColumnName);

        // Create and download zip file
        const zip = new JSZip();
        Object.keys(groupedData).forEach(key => {
          zip.file(`${key}.xlsx`, this.createExcelBlob(groupedData[key], key));
        });

        zip.generateAsync({ type: 'blob' }).then(content => {
          saveAs(content, 'split_files.zip');
        });
      };

      if (this.file) {
        fileReader.readAsArrayBuffer(this.file);
      }
    } else {
      alert('Please select a file and provide a column name.');
    }
  }

  private mergeData(data1: any[], data2: any[], columnName: string) {
    const header1 = data1[0];
    const header2 = data2[0];
    const index1 = header1.indexOf(columnName);
    const index2 = header2.indexOf(columnName);
  
    if (index1 === -1 || index2 === -1) {
      alert('Column name not found in one of the files.');
      return { unionData: [], intersectionData: [] };
    }
  
    const dataMap = new Map<string, any[]>();
    data1.slice(1).forEach((row: any[]) => {
      const key = this.normalizeKey(row[index1]);
      dataMap.set(key, row);
    });
  
    const intersectionData: any[][] = [header1];
    const unionData: any[][] = [header1];
  
    data2.slice(1).forEach((row: any[]) => {
      const key = this.normalizeKey(row[index2]);
      if (dataMap.has(key)) {
        const existingRow = dataMap.get(key);
        const mergedRow = [...existingRow!];
        row.forEach((value: any, i: number) => {
          mergedRow[i] = value;
        });
        dataMap.set(key, mergedRow);
        intersectionData.push(mergedRow);
      }
      unionData.push(row);
    });
  
    dataMap.forEach(row => {
      if (!intersectionData.some(r => r[index1] === row[index1])) {
        unionData.push(row);
      }
    });
  
    return { unionData, intersectionData };
  }  

  private splitData(data: any[], columnName: string): { [key: string]: any[] } {
    const header = data[0];
    const index = header.indexOf(columnName);
    if (index === -1) {
      alert('Column name not found.');
      return {};
    }
  
    const groupedData: { [key: string]: any[] } = {};
  
    data.slice(1).forEach((row: any[]) => {
      const key = this.normalizeKey(row[index]);
      if (!groupedData[key]) {
        groupedData[key] = [header];
      }
      groupedData[key].push(row);
    });
  
    return groupedData;
  }
  

  private normalizeKey(key: any): string {
    if (key === null || key === undefined) {
      return '';
    }
    return key.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase().trim();
  }

  private createExcelBlob(data: any[], sheetName: string): Blob {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    
    // Convert workbook to array buffer
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  
    return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  
}
