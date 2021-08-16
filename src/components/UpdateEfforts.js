import React, { Component } from 'react';
import XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { format } from 'date-fns';
class UpdateEfforts extends Component {
   constructor(props) {
      super(props);
      this.dateMap = new Map();
      this.map = new Map();
      this.mapValueToArray = [];
      this.allowedFiles = ['xlsx', 'xls', 'csv'];
      this.headerString = null;
      this.valueArrayToString = null;
      this.date = new Date();
      this.targetFileDetails = [];
      this.headers = [];
      this.state = {
         disableButtonValue: true,
         excelData: null,
      };
   }

   onFileChange = (event) => {
      for (let i = 0; i < event.target.files.length; i++) {
         let fileExtension = event.target.files[i].name.split('.').pop();
         if (!this.allowedFiles.includes(fileExtension)) {
            alert(
               event.target.files[i].name + ' is not a excel file.Please check'
            );
         } else {
            this.readExcelFiles(event, i);
         }
      }
   };

   readExcelFiles = (event, i) => {
      const excelSheet = event.target.files[i];
      const filePromise = new Promise((resolve, reject) => {
         const fileReader = new FileReader();
         fileReader.readAsArrayBuffer(excelSheet);
         fileReader.onload = (e) => {
            const excelAsArray = e.target.result;
            const readExcel = XLSX.read(excelAsArray, { type: 'buffer' });
            const worksheetName = readExcel.SheetNames[0];
            const openWorksheet = readExcel.Sheets[worksheetName];
            const data = XLSX.utils.sheet_to_json(openWorksheet);

            this.headers = Object.keys(data[0]);
            this.headerString = JSON.stringify(this.headers);

            if (this.headerString.includes('PROJECT_PRODUCTIVE_FLAG')) {
               data.sort(function (a, b) {
                  return a.EMP_ID - b.EMP_ID;
               });
               for (let employee of data) {
                  if (!this.map.has(employee.EMP_ID)) {
                     this.map.set(employee.EMP_ID, '0000000');
                     const date = employee.TIMEPERIOD.toString();
                     // console.log(this.date);
                     this.dateMap.set(employee.EMP_ID, date);
                     this.createmapValueToArray(employee);
                  } else {
                     this.createmapValueToArray(employee);
                  }
                  this.valueArrayToString = this.mapValueToArray.join('');
                  this.map.set(employee.EMP_ID, this.valueArrayToString);
               }
            } else if (
               this.headerString.includes('CALCULATED_EFFORTS') ||
               this.headerString.includes('APPROVED_EFFORTS')
            ) {
               this.targetFileDetails = data
            } else {
               alert('only target and source files are allowed to upload.');
            }
            console.log(this.map);
            console.log(this.dateMap);
         };

         fileReader.onerror = (error) => {
            reject(error);
         };
      });

      filePromise.then((d) => d).catch((e) => e);
   };

   createmapValueToArray = (employee) => {
      this.mapValueToArray = this.map
      .get(employee.EMP_ID)
      .split('');

   if (
      employee.PROJECT_PRODUCTIVE_FLAG.toLowerCase() ===
         'yes' ||
      employee.PROJECT_PRODUCTIVE_FLAG.toLowerCase() === 'no'
   ) {
      if (employee.MON === 8 || employee.MON === 4) {
         this.mapValueToArray[0] = employee.MON === 8 ? 1 : 0.5;
      }
      if (employee.TUE === 8 || employee.TUE === 4) {
         this.mapValueToArray[1] = employee.TUE === 8 ? 1 : 0.5;;
      }
      if (employee.WED === 8 || employee.WED === 4) {
         this.mapValueToArray[2] = employee.WED === 8 ? 1 : 0.5;;
      }
      if (employee.THU === 8 || employee.THU === 4) {
         this.mapValueToArray[3] = employee.THU === 8 ? 1 : 0.5;;
      }
      if (employee.FRI === 8 || employee.FRI === 4) {
         this.mapValueToArray[4] = employee.FRI === 8 ? 1 : 0.5;;
      }
      if (employee.SAT === 8 || employee.SAT === 4) {
         this.mapValueToArray[5] = employee.SAT === 8 ? 1 : 0.5;;
      }
      if (employee.SUN === 8 || employee.SUN === 4) {
         this.mapValueToArray[6] = employee.SUN === 8 ? 1 : 0.5;;
      }
   }
   };

   onFileDownload = () => {
      for (let [key, value] of this.map.entries()) {
         console.log(key, value);
         this.mapValueToArray = value.split('');
         console.log(this.mapValueToArray);
         let date = new Date(this.dateMap.get(key));
         this.mapValueToArray.forEach((element, index) => {
               let datatodays = index === 0 ? date.setDate(new Date(date).getDate() + 0) : date.setDate(new Date(date).getDate() + 1) ;
               const todate = new Date(datatodays);
               const targetCell = this.targetFileDetails.find(item=> item['EFF_DATE(MM/DD/YYYY)'] === format(todate, 'MM/dd/yyyy'));
               targetCell.APPROVED_EFFORTS = element
         }) 
      }
      console.log(this.targetFileDetails);
      const wb = XLSX.utils.book_new();
      wb.Props = {
         Title: 'Merged Sheet',
      };
      wb.SheetNames.push('Merged Data');
      const ws = XLSX.utils.json_to_sheet(this.targetFileDetails);
      wb.Sheets['Merged Data'] = ws;
      const outputSheet = XLSX.write(wb, {
         bookType: 'xlsx',
         type: 'binary',
      });

      var buf = new ArrayBuffer(outputSheet.length); //convert outputSheet to arrayBuffer
      var view = new Uint8Array(buf); //create uint8array as viewer
      for (var i = 0; i < outputSheet.length; i++)
         view[i] = outputSheet.charCodeAt(i) & 0xff; //convert to octet

      saveAs(
         new Blob([buf], { type: 'application/octet-stream' }),
         'Target.xlsx'
      );
   };

   render() {
      return (
         <div>
            <input
               type='file'
               onChange={this.onFileChange}
               multiple='multiple'
               accept='.xlsx, .xls, .csv'
            />
            <button
               onClick={this.onFileDownload}
               // disabled={this.state.disableButtonValue}
            >
               Click here to download the target file
            </button>
         </div>
      );
   }
}

export default UpdateEfforts;
