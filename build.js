const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const dataDir = path.join(__dirname, 'data');
const outputFile = path.join(__dirname, 'file-list.json');

try {
    if (fs.existsSync(dataDir)) {
        const rawFiles = fs.readdirSync(dataDir);
        const finalCsvFiles = [];

        rawFiles.forEach(file => {
            const ext = path.extname(file).toLowerCase();
            const filePath = path.join(dataDir, file);

            if (ext === '.xlsx' || ext === '.xls') {
                const csvFileName = file.replace(ext, '.csv');
                const csvPath = path.join(dataDir, csvFileName);
                
                console.log(`Converting ${file} to CSV...`);
                const workbook = xlsx.readFile(filePath);
                const firstSheet = workbook.SheetNames[0];
                const csvData = xlsx.utils.sheet_to_csv(workbook.Sheets[firstSheet]);
                
                // Write with UTF-8 BOM so our main.js detects it correctly
                fs.writeFileSync(csvPath, '\ufeff' + csvData, 'utf-8');
                finalCsvFiles.push(csvFileName);
            } else if (ext === '.csv') {
                finalCsvFiles.push(file);
            }
        });

        fs.writeFileSync(outputFile, JSON.stringify(finalCsvFiles));
        console.log(`Successfully generated file-list.json with ${finalCsvFiles.length} CSV files.`);
    } else {
        console.log('No data directory found. Generating empty file-list.json.');
        fs.writeFileSync(outputFile, JSON.stringify([]));
    }
} catch (error) {
    console.error('Error generating file-list.json:', error);
    process.exit(1);
}
