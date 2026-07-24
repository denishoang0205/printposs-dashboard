const fs = require('fs');
const path = require('path');

const dataDir = path.join(__dirname, 'data');
const outputFile = path.join(__dirname, 'file-list.json');

try {
    if (fs.existsSync(dataDir)) {
        const files = fs.readdirSync(dataDir).filter(file => file.endsWith('.csv'));
        fs.writeFileSync(outputFile, JSON.stringify(files));
        console.log(`Successfully generated file-list.json with ${files.length} CSV files.`);
    } else {
        console.log('No data directory found. Generating empty file-list.json.');
        fs.writeFileSync(outputFile, JSON.stringify([]));
    }
} catch (error) {
    console.error('Error generating file-list.json:', error);
    process.exit(1);
}
