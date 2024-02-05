const XLSX = require('xlsx');
const path = require('path');

const INPUT_FILE_PATH = path.join(__dirname, 'planilha', 'planilha.xlsx'); 

function calcularSituationAndNAF(matricula, nome, p1, p2, p3, faltas) {
    const media = (p1 + p2 + p3) / 3;
    const totalAulas = 60; 

    if (faltas > 0.25 * totalAulas) {
        console.log(`Aluno com matrícula ${matricula} e nome ${nome} reprovado por falta.`);
        return 'Reprovado por Falta';
    } else if (media < 50) {
        console.log(`Aluno com matrícula ${matricula} e nome ${nome} reprovado por nota.`);
        return 'Reprovado por Nota';
    } else if (media < 7) {
        // Verificar situação "Exame Final"
        const naf = calcularNAF(media);
        console.log(`Aluno com matrícula ${matricula} e nome ${nome} em Exame Final. NAF: ${naf}`);
        return { situation: 'Exame Final', naf: naf };
    } else {
        console.log(`Aluno com matrícula ${matricula} e nome ${nome} aprovado.`);
        return 'Aprovado';
    }
}

function calcularNAF(media) {
    const naf = Math.ceil((5 - media) * 2); 
    return naf >= 0 ? naf : 0; 
}

function main() {
    const workbook = XLSX.readFile(INPUT_FILE_PATH);
    const sheetName = workbook.SheetNames[0]; 
    const worksheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(worksheet['!ref']);

    for (let row = range.s.r + 3; row <= range.e.r; row++) {
        const matricula = worksheet[XLSX.utils.encode_cell({ r: row, c: 0 })].v;
        const nome = worksheet[XLSX.utils.encode_cell({ r: row, c: 1 })].v;
        const faltas = worksheet[XLSX.utils.encode_cell({ r: row, c: 2 })].v;
        const p1 = worksheet[XLSX.utils.encode_cell({ r: row, c: 3 })].v;
        const p2 = worksheet[XLSX.utils.encode_cell({ r: row, c: 4 })].v;
        const p3 = worksheet[XLSX.utils.encode_cell({ r: row, c: 5 })].v;

        const result = calcularSituationAndNAF(matricula, nome, p1, p2, p3, faltas);

        worksheet[XLSX.utils.encode_cell({ r: row, c: 6 })] = { t: 's', v: result };

        if (result === 'Exame Final') {
            const naf = calcularNAF((p1 + p2 + p3) / 3);
            worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })] = { t: 'n', v: naf };
        } else {
            worksheet[XLSX.utils.encode_cell({ r: row, c: 7 })] = { t: 'n', v: 0 };
        }
    }

    const OUTPUT_FILE_PATH = path.join(__dirname, 'planilha', 'nova_planilha.xlsx'); 
    XLSX.writeFile(workbook, OUTPUT_FILE_PATH);
    console.log('Resultados escritos no novo arquivo Excel.');
}

main();
