# AngularXlsxService

usage:

```typescript
import {AngularXlsxService, AngularXlsxModule} from 'angular-xlsx-service';
import {Component} from "@angular/core";

@Component({
  selector: 'app-root',
  template: `
    <button (click)="export()">Export</button>
  `,
  standalone: true,
  imports: [AngularXlsxModule]
})
export class AppComponent {
  private xlsxService = inject(AngularXlsxService)

  public export(): void {
    this.xlsxService.exportJsonToExcel({
      data: [
        {name: 'Juri', surname: 'Strumpflohner'},
        {name: 'John', surname: 'Doe'}
      ],
      fileName: 'myFile',
      sheetName: 'Sheet1'
    }).download()
  }
}
```
