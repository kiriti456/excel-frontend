import { HttpClient } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';

@Injectable({
  providedIn: 'root'
})
export class ExcelOperationsService {

  private baseUrl = 'http://localhost:8080/api/excel'; // Adjust the base URL to match your backend

  constructor(private http: HttpClient) { }

  mergeFiles(file1: File, file2: File, columnName: string): Observable<Blob> {
    const formData = new FormData();
    formData.append('files', file1);
    formData.append('files', file2);
    formData.append('columnName', columnName);

    console.log("in the merge");

    return this.http.post(`${this.baseUrl}/merge`, formData, { responseType: 'blob' });
  }

  splitFile(file: File, columnName: string): Observable<Blob> {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('columnName', columnName);

    return this.http.post(`${this.baseUrl}/split`, formData, { responseType: 'blob' });
  }
}
