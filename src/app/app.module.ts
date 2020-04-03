import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { AppComponent } from './app.component';
import { ExcelDownloadHelper } from './excel-download/excel-download-helper';

@NgModule({
  declarations: [
    AppComponent
  ],
  imports: [
    BrowserModule
  ],
  providers: [
    ExcelDownloadHelper
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
