const excel = require('excel4node');

const workbook = new excel.Workbook();
const filePath = 'template/SquareTemplate.xlsx'; // Replace with the actual path to your file
workbook.xlsx.readFile(filePath)
    .then(() => {
        // File loaded successfully
        const worksheet = workbook.getWorksheet('Sheet1'); // Replace 'Sheet1' with the actual sheet name

        // Writing to cell D3
        const cellD8 = worksheet.getCell('D8');
        cellD8.value = 115367;

        // Reading from cell D9
        const cellE75 = worksheet.getCell('E75');
        const valueE75 = cellE75.value;

        console.log('Value in cell D9:', valueE75);

        // Save the changes to the file
        return workbook.xlsx.writeFile(filePath);
    })
    .then(() => {
        console.log('File saved successfully.');
    })
    .catch((error) => {
        console.error('Error:', error);
    });