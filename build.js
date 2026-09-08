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
                
                // Only convert if CSV does not exist or XLSX is newer
                const needConvert = !fs.existsSync(csvPath) || fs.statSync(filePath).mtimeMs > fs.statSync(csvPath).mtimeMs;
                if (needConvert) {
                    console.log(`Converting ${file} to CSV...`);
                    try {
                        const workbook = xlsx.readFile(filePath);
                        const firstSheet = workbook.SheetNames[0];
                        const csvData = xlsx.utils.sheet_to_csv(workbook.Sheets[firstSheet]);
                        
                        // Write with UTF-8 BOM so our main.js detects it correctly
                        fs.writeFileSync(csvPath, '\ufeff' + csvData, 'utf-8');
                    } catch (err) {
                        console.warn(`Could not write ${csvFileName} (may be locked or open):`, err.message);
                    }
                }
                if (fs.existsSync(csvPath) && !finalCsvFiles.includes(csvFileName)) {
                    finalCsvFiles.push(csvFileName);
                }
            } else if (ext === '.csv') {
                if (!finalCsvFiles.includes(file)) {
                    finalCsvFiles.push(file);
                }
            }
        });

        fs.writeFileSync(outputFile, JSON.stringify(finalCsvFiles));
        console.log(`Successfully generated file-list.json with ${finalCsvFiles.length} CSV files:`, finalCsvFiles);
    } else {
        console.log('No data directory found. Generating empty file-list.json.');
        fs.writeFileSync(outputFile, JSON.stringify([]));
    }
} catch (error) {
    console.error('Error generating file-list.json:', error);
    process.exit(1);
}
