const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE_PATH = path.join(__dirname, 'planilha', 'planilha.xlsx'); 

function main() {
  const workbook = XLSX.readFile(INPUT_FILE_PATH);
  const sheetName = workbook.SheetNames[0]; 
  const worksheet = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(worksheet['!ref']);

  for (let row = range.s.r; row <= range.e.r; row++) {
    const cellAddress = XLSX.utils.encode_cell({ r: row, c: 1 }); 
    const cell = worksheet[cellAddress];

    if (cell && cell.t === 's') {
      cell.v += '_Modificado'; 
    }
  }

  const OUTPUT_FILE_PATH = path.join(__dirname, 'planilha', 'nova_planilha.xlsx');
  XLSX.writeFile(workbook, OUTPUT_FILE_PATH);
  console.log('Resultados escritos no novo arquivo Excel.');
}

main();
