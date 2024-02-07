const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE_PATH = path.join(__dirname, 'planilha','planilha.xlsx');

function calculateSituationAndNAF(studentId, name, grade1, grade2, grade3, absences) {
    const average = (grade1 + grade2 + grade3) / 3;
    const totalClasses = 60;

    if (absences > 0.25 * totalClasses) {
        console.log(`Student with ID ${studentId} and name ${name} failed due to excessive absences.`);
        return 'Failed due to Absences';
    } else if (average < 50) {
        console.log(`Student with ID ${studentId} and name ${name} failed due to low grades.`);
        return 'Failed due to Grades';
    } else if (average < 70) {
        // Check status "Final Exam"
        const naf = calculateNAF(average);
        console.log(`Student with ID ${studentId} and name ${name} in Final Exam. NAF: ${naf}`);
        return 'Final Exam';
    } else {
        console.log(`Student with ID ${studentId} and name ${name} passed.`);
        return 'Passed';
    }
}

function calculateNAF(average) {
    console.log(average);
    const naf = Math.ceil((50 - average) * 2);

    const finalNAF = naf >= 0 ? naf : 0;
    console.log(`Test ${finalNAF}`);
    console.log(`NAF ${naf}`);
    return finalNAF;
}

function main() {
    const workbook = XLSX.readFile(INPUT_FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    for (let row = range.s.r + 3; row <= range.e.r; row++) {
        const studentId = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })].v;
        const name = worksheet[XLSX.utils.encode_cell({ r: row, c: 1 })].v;
        const absences = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })].v;
        const grade1 = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })].v;
        const grade2 = worksheet[XLSX.utils.encode_cell({ r: row, c: 4 })].v;
        const grade3 = worksheet[XLSX.utils.encode_cell({ r: row, c: 5 })].v;

        const result = calculateSituationAndNAF(studentId, name, grade1, grade2, grade3, absences);

        worksheet[XLSX.utils.encode_cell({ r: row, c: 6 })] = { t: 's', v: result };

        if (result === 'Final Exam') {
            const naf = calculateNAF((grade1 + grade2 + grade3) / 3);
            worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })] = { t: 'n', v: naf };
        } else {
            worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })] = { t: 'n', v: 0 };
        }
    }

    const OUTPUT_FILE_PATH = path.join(__dirname, 'planilha', 'planilha.xlsx');
    XLSX.writeFile(workbook, OUTPUT_FILE_PATH);
    console.log('Results written to the new Excel file.');
}

main();
