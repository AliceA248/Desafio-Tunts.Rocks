const XLSX = require('xlsx');
const path = require('path');

// File paths for input and output 
const INPUT_FILE_PATH = path.join(__dirname, 'planilha', 'Engenharia de Software – Desafio Alice Pereira.xlsx');

const STARTING_ROW = 3;

function calculateSituationAndNAF(matricula, nome, p1, p2, p3, faltas) {
    const media = (p1 + p2 + p3) / 3;
    const totalAulas = 60;

    // Check for excessive absences
    if (faltas > 0.25 * totalAulas) {
        return 'Reprovado por Falta';
    } else if (media < 50) {
        return 'Reprovado por Nota';
    } else if (media < 70) {
        // Proceed to "Final Exam" status and calculate NAF
        const naf = calculateNAF(media);
        return 'Exame Final';
    } else {
        return 'Aprovado';
    }
}

function calculateNAF(media) {
    // Calculate NAF based on the formula
    const naf = Math.ceil((70 - media) * 2);
    return naf >= 0 ? naf : 0;
}

function processWorksheet(worksheet) {
    // Decode the range of the worksheet
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    for (let row = range.s.r + STARTING_ROW; row <= range.e.r; row++) {
        const matricula = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })].v;
        const nome = worksheet[XLSX.utils.encode_cell({ r: row, c: 1 })].v;
        const faltas = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })].v;
        const p1 = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })].v;
        const p2 = worksheet[XLSX.utils.encode_cell({ r: row, c: 4 })].v;
        const p3 = worksheet[XLSX.utils.encode_cell({ r: row, c: 5 })].v;

        const result = calculateSituationAndNAF(matricula, nome, p1, p2, p3, faltas);

        worksheet[XLSX.utils.encode_cell({ r: row, c: 6 })] = { t: 's', v: result };

        // If the result is "Final Exam," calculate and update the NAF value
        if (result === 'Exame Final') {
            const naf = calculateNAF((p1 + p2 + p3) / 3);
            worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })] = { t: 'n', v: naf };
        } else {
            // If not, set the NAF value to 0
            worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })] = { t: 'n', v: 0 };
        }
    }
}

function main() {
    try {
        const workbook = XLSX.readFile(INPUT_FILE_PATH);

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        processWorksheet(worksheet);

        XLSX.writeFile(workbook, INPUT_FILE_PATH);

        console.log('Changes saved');
    } catch (error) {
        console.error('Error:', error.message);
    }
}

main();
