import {Component} from '@angular/core';

declare var XLSX;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {

  loadAPI: Promise<any>;

  constructor() {
    this.loadAPI = new Promise((resolve) => {
      this.loadScript();
      resolve(true);
    });

    this.loadAPI.then((res) => {
      console.log('Script loaded');
    });
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

  public loadScript() {
    let isFound = false;
    const scripts = document.getElementsByTagName('script');
    for (let i = 0; i < scripts.length; ++i) {
      if (scripts[i].getAttribute('src') != null && scripts[i].getAttribute('src').includes('xlsx')) {
        isFound = true;
      }
    }

    if (!isFound) {
      const node = document.createElement('script');
      node.src = 'assets/xlsx.full.min.js';
      node.type = 'text/javascript';
      node.async = false;
      node.charset = 'utf-8';
      document.getElementsByTagName('head')[0].appendChild(node);

      node.onerror = () => {
        console.log('Error loading ' + node.src);
      };
    }
  }
}
