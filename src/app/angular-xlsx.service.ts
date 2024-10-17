import {Injectable} from "@angular/core";
import * as XLSX from 'xlsx';
import { WorkBook, WorkSheet } from 'xlsx/types';
import { saveAs } from 'file-saver';

export interface IParseResult {
  file: any
  download: () => void
}

export interface IJsonExportOptions {
  name?: string,
  colWidths?: number[]
  mergeIfMatches?: string[]
  mergeColumns?: string[]
}

@Injectable()
export class AngularXlsxService {
  /**
   * Exports an HTML table element to an Excel file.
   * @param tableElement The HTMLTableElement to export.
   * @param fileName The name of the generated Excel file.
   */
  public parseTable(tableElement: HTMLElement, fileName = 'Angular-XLSX-Service'): IParseResult {
    const worksheet: WorkSheet = XLSX.utils.table_to_sheet(tableElement);
    const workbook: WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const excelBuffer: any = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });

    return {
      file: excelBuffer,
      download: () => this.saveAsExcelFile(excelBuffer, fileName),
    }
  }

  /**
   * Exports a JSON array to an Excel file.
   * @param jsonData The JSON data to export.
   * @param params parameters like name, and merge options.
   */
  public exportJsonToExcel(jsonData: any[], params ?: IJsonExportOptions): IParseResult {
    const worksheet: WorkSheet = XLSX.utils.json_to_sheet(jsonData);

    if(params?.colWidths) {
      worksheet['!cols'] = params.colWidths.map(width => ({ wpx: width }))
    }

    const merges: XLSX.Range[] = [];

    let mergeStartRow = -1;
    let previousCallId: string[] | undefined;

    if(params?.mergeIfMatches?.length) {
      const columnsToMerge = params.mergeColumns?.map((colName: string) => Object.keys(jsonData[0]).indexOf(colName));
      console.log(columnsToMerge)
      jsonData.forEach((row, index) => {
        const currentRow = index + 1;
        const currentCallId = params.mergeIfMatches?.map(key => row[key]);

        if (currentCallId && currentCallId.every((param, index) => param === previousCallId?.[index])) {
          if (mergeStartRow === -1) {
            mergeStartRow = currentRow - 1;
          }
        } else {
          this.checkAndAddMerge(merges, mergeStartRow, currentRow, columnsToMerge ?? []);
          mergeStartRow = -1;
        }

        previousCallId = currentCallId;
      });
      this.checkAndAddMerge(merges, mergeStartRow, jsonData.length + 1, columnsToMerge ?? []);

      worksheet['!merges'] = merges;
    }

    const workbook: XLSX.WorkBook = { Sheets: { data: worksheet }, SheetNames: ['data'] };
    const excelBuffer: any = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    return {
      file: excelBuffer,
      download: () => this.saveAsExcelFile(excelBuffer, params?.name ?? 'Angular-XLSX-Service'),
    }
  }

  private saveAsExcelFile(buffer: any, fileName: string): void {
    const data: Blob = new Blob([buffer], {
      type:
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8',
    });
    saveAs(data, `${fileName}.xlsx`);
  }

  private checkAndAddMerge(merges: XLSX.Range[], mergeStartRow: number, currentRow: number, columnsToMerge: number[]): void {
    if (mergeStartRow !== -1 && mergeStartRow < currentRow - 1) {
      columnsToMerge.forEach((col) => {
        merges.push({
          s: { r: mergeStartRow, c: col },
          e: { r: currentRow - 1, c: col }
        });
      });
    }
  }
}
