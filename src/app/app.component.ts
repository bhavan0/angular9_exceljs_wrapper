import { Component } from '@angular/core';
import { ExcelDownloadHelper } from './excel-download/excel-download-helper';
import { ExcelColumn } from './excel-download/models/excel-column';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  title = 'angular9-exceljs-wrapper';


  constructor(private excelDownloadHelper: ExcelDownloadHelper) {
  }

  downloadExcel() {
    // create workBook
    this.excelDownloadHelper.setUpExcelWorkBook();

    // Add Workbook Name
    this.excelDownloadHelper.setExcelWorkBookName('TestWorkBook');

    // Add two sheets to the workBook
    const productSheet = this.excelDownloadHelper.createAndGetWorkSheet('Products');
    const hiddenSheet = this.excelDownloadHelper.createAndGetWorkSheet('HiddenTest');

    // Hide second Sheet and add password to lock sheet
    this.excelDownloadHelper.hideSheet(hiddenSheet);
    this.excelDownloadHelper.protectSheet(hiddenSheet, 'password');

    // Add a column with data in hidden sheet at column 1
    this.excelDownloadHelper.addColumnData(hiddenSheet, 1, ['List1', 'List2', 'List3']);

    // Set up column headers to be shown in Products Sheet
    const columns: ExcelColumn[] = [
      { key: 'string', width: 20, header: 'String' },
      { key: 'textarea', width: 20, header: 'Textarea', style: { alignment: { wrapText: true } } },
      { key: 'number', width: 20, header: 'Number', style: { numFmt: '#,##0.00#######' } },
      { key: 'boolean', width: 20, header: 'Boolean' },
      { key: 'date', width: 20, header: 'Date', style: { numFmt: 'dd-MM-yyyy' } },
      { key: 'datetime', width: 20, header: 'Datetime', style: { numFmt: 'dd-MM-yyyy hh:mm AM/PM' } },
    ];

    // Add these columns Headers to the Products Sheet
    this.excelDownloadHelper.addAllColumns(productSheet, columns);

    // Add a specific Column Header at location 7
    this.excelDownloadHelper.addColumnHeaderAtLocation(productSheet, 7, 'TestNew');

    // Set up row data to be filled
    const rows = [
      [
        'string', 'textarea\r\ntextarea', 55.1, true, new Date(2020, 0, 1),
        new Date(Date.UTC(2020, 0, 1, 14, 55, 12))
      ],
      [
        'string', 'textarea\r\ntextarea', 55.1, true, new Date(2020, 0, 1),
        new Date(Date.UTC(2020, 0, 1, 14, 55, 12))
      ]
    ];

    // Add the rows to the Products Sheet
    this.excelDownloadHelper.addRow(productSheet, rows);

    // Freeze the first 4 columns of the Sheet
    this.excelDownloadHelper.freezeOnlyColumns(productSheet, 4);

    // Hide the first column of the Product Sheet
    this.excelDownloadHelper.hideColumn(productSheet, 1);

    // Add Style Header for the first row
    this.excelDownloadHelper.addRowHeaderStyle(productSheet, 1);

    // Add drop down formula for the column 7 getting data from the hidden sheet
    this.excelDownloadHelper.addDropDown(productSheet, 7, hiddenSheet, 'A1', 'A3');

    this.excelDownloadHelper.addColumnLengthValidation(productSheet, 8, 10);

    // Download the final Sheet
    this.excelDownloadHelper.downloadExcel();
  }
}
