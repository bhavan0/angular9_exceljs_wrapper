import 'core-js/modules/es.promise';
import 'core-js/modules/es.object.assign';
import 'core-js/modules/es.object.keys';
import 'regenerator-runtime/runtime';
import * as Excel from 'exceljs/dist/exceljs.min.js';
import * as fs from 'file-saver';
import { ExcelColumn } from './models/excel-column';

export class ExcelDownloadHelper {

    workBookName: string;
    workbook: any;

    //#region WorkSheet Methods

    setUpExcelWorkBook() {
        this.workbook = new Excel.Workbook();
    }

    setExcelWorkBookName(workBookName: string) {
        this.workBookName = workBookName;
    }

    createAndGetWorkSheet(workSheetName: string): any {
        return this.workbook.addWorksheet(workSheetName);
    }

    hideSheet(workSheet: any) {
        workSheet.state = 'veryHidden';
    }

    async protectSheet(workSheet: any, password: string) {
        await workSheet.protect(password);
    }

    async downloadExcel() {
        const buffer = await this.workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        fs.saveAs(blob, this.workBookName);
    }

    //#endregion


    //#region Columns Region

    addAllColumns(workSheet: any, columns: ExcelColumn[]) {
        workSheet.columns = columns;
    }

    addColumnHeaderAtLocation(worksheet: any, columnNo: number, headerText: string) {
        worksheet.getColumn(columnNo).values = [headerText];
    }

    addColumnData(workSheet: any, columnNo: number, data: string[]) {
        workSheet.getColumn(columnNo).values = data;
    }

    freezeOnlyColumns(workSheet: any, columnNo: number) {
        workSheet.views = [
            { state: 'frozen', xSplit: columnNo, ySplit: 0, activeCell: 'A1' }
        ];
    }

    hideColumn(workSheet: any, columnNo: number) {
        workSheet.getColumn(columnNo).hidden = true;
    }

    addColumnLengthValidation(workSheet: any, columnNo: number, maxLength: number) {
        workSheet.dataValidations.add('I1:I99999', {
            type: 'textLength',
            operator: 'lessThan',
            showErrorMessage: true,
            allowBlank: true,
            formulae: [maxLength]
        });
    }

    addDropDown(workSheet: any, columnNo: number, sourceSheet: any, startCell: string, endCell: string) {
        workSheet.getColumn(columnNo).eachCell({ includeEmpty: false }, (cell: any) => {
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: [sourceSheet.name + '!' + startCell + ':' + endCell]
            };
        });
    }

    //#endregion

    //#region Rows

    addRow(workSheet: any, rows: any) {
        workSheet.addRows(rows);
    }

    addRowHeaderStyle(workSheet: any, rowNo: number) {
        workSheet.getRow(rowNo).eachCell((cell: any) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: '4F81BD' }
            };
            cell.font = {
                bold: true,
                color: { argb: 'FFFFFF' }
            };
        });
    }

    freezeOnlyRows(workSheet: any, rowNo: number) {
        workSheet.views = [
            { state: 'frozen', xSplit: 0, ySplit: rowNo, topLeftCell: 'G10', activeCell: 'A1' }
        ];
    }

    //#endregion

}
