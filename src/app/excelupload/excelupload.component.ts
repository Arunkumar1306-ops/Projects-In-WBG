import { Component } from '@angular/core';

import * as XLSX from 'XLSX'
import { read, utils, writeFile } from 'XLSX';

@Component({
  selector: 'app-excelupload',
  templateUrl: './excelupload.component.html',
  styleUrl: './excelupload.component.css'
})
export class ExceluploadComponent {
upload: any[] = [];
handleImport(event: any) {
   const files = event.target.files;
  if (files.length) {
      const file = files[0];
      const reader = new FileReader();
      reader.onload = (event: any) => {
          const wb = read(event.target.result);
          const sheets = wb.SheetNames;

          if (sheets.length) {
              const rows:any[ ]= utils.sheet_to_json(wb.Sheets[sheets[0]]);
              const validationResult=this.validateData(rows);
              if(validationResult.isValid){
              this.upload = rows;
              alert('upload successfully')
          }
        
              else{
                  alert('Error:'+ validationResult.errorMessage);
              }
      }
    }
      reader.readAsArrayBuffer(file);
  }
 }

 validateData(rows:any[]):{ isValid:boolean;errorMessage:string}{
  for(const row of rows){
    if(!this.validateEmail(row.EmailId)){
      return {isValid:false,errorMessage:'Invalid Email Format found.'}
    }
  }
  return {isValid:true,errorMessage:''};
 }

 validateEmail(email:string):boolean{
  const emailRegex=/^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
 }
}
//  handleExport() {
//   const headings = [[
//       'FirstName',
//       'LastName',
//       'DOB',
//       'EmailId'
//   ]];
//   const wb = utils.book_new();
//   const ws: any = utils.json_to_sheet([]);
//   utils.sheet_add_aoa(ws, headings);
//   utils.sheet_add_json(ws, this.upload, { origin: 'A2', skipHeader: true });
//   utils.book_append_sheet(wb, ws, 'Report');
//   writeFile(wb, 'ExcelUploaded.xlsx');
// }
// }















