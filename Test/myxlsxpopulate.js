const XlsxPopulate = require('xlsx-populate');

const filePath = 'template/def.xlsx'; // Replace with the actual path to your file

XlsxPopulate.fromFileAsync(filePath)
    .then(workbook => {
        // File loaded successfully
        const worksheet = workbook.sheet(0); // Replace 'Sheet1' with the actual sheet name

        // Writing to cell D3
        worksheet.cell('D8').value(115367);
        worksheet.cell('E8').value(359452032);

        var result = workbook.toFileAsync(filePath);

        // Reading from cell D9
        const valueE75 = worksheet.cell('D75').value();

        console.log('Value in cell E75:', valueE75);

        // Save the changes to the file
        return result;
    })
    .then(() => {
        console.log('File saved successfully.');
    })
    .catch((error) => {
        console.error('Error:', error);
    });
