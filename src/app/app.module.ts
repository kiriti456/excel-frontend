import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { provideHttpClient, withFetch } from '@angular/common/http';
import { FormsModule } from '@angular/forms';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { ExcelOperationsComponent } from './excel-operations/excel-operations.component';

@NgModule({
  declarations: [
    AppComponent,
    ExcelOperationsComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    FormsModule
  ],
  providers: [
    provideHttpClient(withFetch())  // Use the new provideHttpClient with withFetch
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
