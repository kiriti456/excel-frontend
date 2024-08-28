import { Component } from '@angular/core';
import * as XLSX from 'xlsx';
import JSZip from 'jszip'; // Changed to default import

@Component({
  selector: 'app-excel-operations',
  templateUrl: './excel-operations.component.html',
  styleUrls: ['./excel-operations.component.scss']
})
export class ExcelOperationsComponent {
  file1: File | null = null;
  file2: File | null = null;
  splitFileData: File | null = null; // Renamed to avoid conflict
  mergeColumnName: string = '';
  splitColumnName: string = '';

  onFile1Selected(event: any) {
    this.file1 = event.target.files[0];
  }

  onFile2Selected(event: any) {
    this.file2 = event.target.files[0];
  }

  onFileSelected(event: any) {
    this.splitFileData = event.target.files[0];
  }

  async mergeFiles() {
    if (!this.file1 || !this.file2 || !this.mergeColumnName) {
      alert('Please upload two files and specify a column name.');
      return;
    }

    const data1 = await this.readExcelFile(this.file1);
    const data2 = await this.readExcelFile(this.file2);

    const columnIdx1 = this.getColumnIndex(data1[0], this.mergeColumnName);
    const columnIdx2 = this.getColumnIndex(data2[0], this.mergeColumnName);

    if (columnIdx1 === -1 || columnIdx2 === -1) {
      alert('Column name not found in one or both files.');
      return;
    }

    const { unionData, intersectionData } = this.mergeAndIntersect(data1, data2, columnIdx1, columnIdx2);

    const unionBlob = this.createExcelBlob(unionData, 'Union');
    const intersectionBlob = this.createExcelBlob(intersectionData, 'Intersection');

    const zipBlob = await this.createZipFile(unionBlob, intersectionBlob);

    this.downloadFile(zipBlob, 'merged_intersected_files.zip');
  }

  async splitExcelFile() { // Renamed the function
    if (!this.splitFileData || !this.splitColumnName) {
      alert('Please upload a file and specify a column name.');
      return;
    }

    const data = await this.readExcelFile(this.splitFileData);
    const columnIdx = this.getColumnIndex(data[0], this.splitColumnName);

    if (columnIdx === -1) {
      alert('Column name not found in the file.');
      return;
    }

    const splitFiles = this.splitByColumn(data, columnIdx);
    const zipBlob = await this.createSplitZipFile(splitFiles);

    this.downloadFile(zipBlob, 'split_files.zip');
  }

  private async readExcelFile(file: File): Promise<any[][]> {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  }

  private getColumnIndex(headerRow: any[], columnName: string): number {
    return headerRow.findIndex(header => this.normalizeKey(header) === this.normalizeKey(columnName));
  }

  private mergeAndIntersect(data1: any[][], data2: any[][], columnIdx1: number, columnIdx2: number) {
    const headers1 = data1[0];
    const headers2 = data2[0];

    const allHeadersSet = new Set([...headers1, ...headers2]);
    const allHeaders = Array.from(allHeadersSet);

    const unionData = [allHeaders];
    const intersectionData = [allHeaders];

    const rowsMap1 = this.mapRowsByColumn(data1, columnIdx1, headers1);
    const rowsMap2 = this.mapRowsByColumn(data2, columnIdx2, headers2);

    const allKeys = new Set([...rowsMap1.keys(), ...rowsMap2.keys()]);

    allKeys.forEach(key => {
      const row1 = rowsMap1.get(key);
      const row2 = rowsMap2.get(key);
      const unionRow = this.createMergedRow(row1, row2, allHeaders, headers1, headers2);

      unionData.push(unionRow);

      if (row1 && row2) {
        const intersectionRow = this.createMergedRow(row1, row2, allHeaders, headers1, headers2);
        intersectionData.push(intersectionRow);
      }
    });

    return { unionData, intersectionData };
  }

  private mapRowsByColumn(data: any[][], columnIndex: number, headers: any[]): Map<string, any[]> {
    const rowMap = new Map<string, any[]>();
    data.slice(1).forEach(row => {
      const key = this.normalizeKey(row[columnIndex]);
      rowMap.set(key, row);
    });
    return rowMap;
  }

  private createMergedRow(
    row1: any[] | undefined,
    row2: any[] | undefined,
    allHeaders: any[],
    headers1: any[],
    headers2: any[]
  ): any[] {
    const mergedRow = new Array(allHeaders.length).fill('');

    if (row1) {
      headers1.forEach((header, i) => {
        const colIdx = allHeaders.indexOf(header);
        mergedRow[colIdx] = row1[i];
      });
    }

    if (row2) {
      headers2.forEach((header, i) => {
        const colIdx = allHeaders.indexOf(header);
        mergedRow[colIdx] = row2[i];
      });
    }

    return mergedRow;
  }

  private splitByColumn(data: any[][], columnIdx: number): Map<string, any[][]> {
    const headerRow = data[0];
    const splitFiles = new Map<string, any[][]>();

    data.slice(1).forEach(row => {
      const key = this.normalizeKey(row[columnIdx]);
      if (!splitFiles.has(key)) {
        splitFiles.set(key, [headerRow]);
      }
      splitFiles.get(key)!.push(row);
    });

    return splitFiles;
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

    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  private async createZipFile(unionBlob: Blob, intersectionBlob: Blob): Promise<Blob> {
    const zip = new JSZip(); // No error now
    zip.file('union.xlsx', unionBlob);
    zip.file('intersection.xlsx', intersectionBlob);
    return await zip.generateAsync({ type: 'blob' });
  }

  private async createSplitZipFile(splitFiles: Map<string, any[][]>): Promise<Blob> {
    const zip = new JSZip(); // No error now

    splitFiles.forEach((data, key) => {
      const blob = this.createExcelBlob(data, key);
      zip.file(`${key}.xlsx`, blob);
    });

    return await zip.generateAsync({ type: 'blob' });
  }

  private downloadFile(blob: Blob, fileName: string) {
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(link.href);
  }
}
