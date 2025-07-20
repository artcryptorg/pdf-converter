import * as pdfjsLib from 'pdfjs-dist';
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

const pdfFolder = './pdf_files'; // Папка с PDF файлами
const outputExcel = './output.xlsx'; // Путь к итоговому Excel

async function extractDataFromPDF(pdfPath) {
    const pdfDocument = await pdfjsLib.getDocument(pdfPath).promise;
    const page = await pdfDocument.getPage(1);
    const textContent = await page.getTextContent();

    const groupedByY = {};
    textContent.items.forEach(item => {
        const y = Math.round(item.transform[5]);
        if (!groupedByY[y]) groupedByY[y] = [];
        groupedByY[y].push(item.str);
    });

    const groupedRows = Object.keys(groupedByY)
        .sort((a, b) => b - a)
        .map(y => groupedByY[y].join(' '));

    console.log('Сгруппированные строки:', groupedRows);

    let numeroFattura = '';
    let data = '';
    let prodotto = '';
    let riferimento = '';
    let iva = '';
    let prezzoBase = '';

    groupedRows.forEach((row, index) => {
        if (row.includes('Numero Fattura') && row.includes('Data di Emissione')) {
            const parts = groupedRows[index + 1]?.split(/\s{2,}/);
            if (parts) {
                numeroFattura = parts[0];
                data = parts[1];
            }
        }

        if (/IT\d{4}PRO/.test(row)) {
            const parts = row.split(/\s{2,}/);
            riferimento = parts[0]?.trim();
            if (parts.length > 1) {
                prodotto = parts[1]?.trim();
            }
        }

        if (row.includes('Totale (imp escl.)')) {
            const parts = row.split(/\s+/);
            prezzoBase = parts[parts.length - 2];
        }
    });

    return {
        'Numero Fattura': numeroFattura,
        Data: data,
        Riferimento: riferimento,
        Prodotto: prodotto,
        IVA: iva,
        'Prezzo Base': prezzoBase,
    };
}

async function processAllPDFs() {
    if (!fs.existsSync(pdfFolder)) {
        console.error('Папка с PDF файлами не найдена:', pdfFolder);
        return;
    }

    const files = fs.readdirSync(pdfFolder).filter(file => file.endsWith('.pdf'));
    if (files.length === 0) {
        console.error('Нет PDF файлов для обработки.');
        return;
    }

    const excelData = [];
    for (const file of files) {
        const filePath = path.join(pdfFolder, file);
        console.log(`Обрабатывается файл: ${file}`);
        try {
            const extractedData = await extractDataFromPDF(filePath);
            console.log('Извлечённые данные:', extractedData);
            excelData.push(extractedData);
        } catch (err) {
            console.error(`Ошибка обработки ${file}:`, err);
        }
    }

    const worksheet = XLSX.utils.json_to_sheet(excelData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');

    try {
        XLSX.writeFile(workbook, outputExcel);
        console.log(`Данные успешно сохранены в файл: ${outputExcel}`);
    } catch (err) {
        if (err.code === 'EBUSY') {
            console.error('Файл output.xlsx заблокирован. Закройте его и попробуйте снова.');
        } else {
            console.error('Ошибка при сохранении файла:', err);
        }
    }
}

processAllPDFs().catch(err => console.error('Ошибка:', err));
