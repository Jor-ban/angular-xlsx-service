import {Component, ElementRef, inject, viewChild} from '@angular/core';
import { RouterOutlet } from '@angular/router';
import {AngularXlsxService} from "./angular-xlsx.service";
import {JsonPipe} from "@angular/common";

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [
    JsonPipe
  ],
  providers: [
    AngularXlsxService,
  ],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss'
})
export class AppComponent {
  public readonly table = viewChild.required<ElementRef<HTMLTableElement>>('table')
  public readonly jsonData = [
    { id: 1, name: 'John Doe', age: 30, city: 'New York' },
    { id: 2, name: 'Jane Smith', age: 25, city: 'Los Angeles' },
    { id: 3, name: 'Sam Johnson', age: 40, city: 'Chicago' },
    { id: 4, name: 'Alice Brown', age: 35, city: 'Houston' },
    { id: 5, name: 'Bob Davis', age: 28, city: 'Phoenix' },
    { id: 5, name: 'Bob Davis', age: 30, city: 'San Francisco' }
  ]
  private readonly xlsxService = inject(AngularXlsxService)

  public tableExport(): void {
    this.xlsxService.parseTable(this.table().nativeElement, 'table-export').download()
  }

  public jsonExport(): void {
    this.xlsxService.exportJsonToExcel(
      this.jsonData,
      {
        name: 'merged excel',
        mergeIfMatches: ['id'],
        mergeColumns: ['name', 'age'],
        colWidths: [0, 200, 50, 350]
      }
    ).download()
  }
}
