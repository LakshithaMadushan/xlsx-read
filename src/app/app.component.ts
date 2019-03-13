import {Component} from '@angular/core';

declare var XLSX;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  constructor() {
    console.log('load xlsx');
  }

  Upload() {
    const fileUpload = (document.getElementById('fileUpload') as HTMLInputElement);

    const regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    if (regex.test(fileUpload.value.toLowerCase())) {
      if (typeof (FileReader) !== 'undefined') {
        const reader = new FileReader();

        if (reader.readAsBinaryString) {
          reader.onload = (e) => {
            this.ProcessExcel(reader.result);
          };
          reader.readAsBinaryString(fileUpload.files[0]);
        }
      } else {
        alert('This browser does not support HTML5.');
      }
    } else {
      alert('Please upload a valid Excel file.');
    }
  }

  ProcessExcel(data) {
    const workbook = XLSX.read(data, {
      type: 'binary'
    });

    const firstSheet = workbook.SheetNames[0];

    const excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[firstSheet]);
    console.log(excelRows);
  }
}
