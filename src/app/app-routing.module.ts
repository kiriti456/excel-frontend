import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { ExcelOperationsComponent } from './excel-operations/excel-operations.component';

const routes: Routes = [
  { path: 'excel-operations', component: ExcelOperationsComponent },
  { path: '', redirectTo: '/excel-operations', pathMatch: 'full' }
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
