const unzipper = require('unzipper');
const fs = require("fs");
function formatDate(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();
    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;
    return [year, month, day].join('-');
}
function formatDateSlash(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();
    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;
    return [year, month, day].join('/');
}

function formatDateNull(date) {
    var d = new Date(date),
        month = '' + (d.getMonth() + 1),
        day = '' + d.getDate(),
        year = d.getFullYear();
    if (month.length < 2)
        month = '0' + month;
    if (day.length < 2)
        day = '0' + day;
    return [year, month, day].join('');
}

function unzipFile(zipFilePath, destinationFolder) {s
    fs.createReadStream(zipFilePath)
        .pipe(unzipper.Extract({ path: destinationFolder }))
        .on('close', () => {
            console.log('Unzip complete!');
        })
        .on('error', (err) => {
            console.error('Error occurred while unzipping:', err);
        });
}

module.exports ={formatDate,formatDateSlash,formatDateNull, unzipFile};