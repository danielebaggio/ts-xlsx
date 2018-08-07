import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';

export const tsXLXS = () => {
    const xlsxLib = {
        buffer: '',
        exportAsExcelFile(json: any[]){
            const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
            const workbook: XLSX.WorkBook = { Sheets: { 'data': worksheet }, SheetNames: ['data'] };
            this.buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
            return this;
        },
        saveAsExcelFile(fileName: string){
            const data: Blob = new Blob([this.buffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8'});
            FileSaver.saveAs(data, fileName + '.xlsx');
            return this;
         }
    }
    return xlsxLib;
}
