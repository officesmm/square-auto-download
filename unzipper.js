const decompress = require("decompress");

decompress("C:\\Users\\SoeMyatMin\\Downloads\\【支払明細票 兼 売上票】日本恒生ソフトウェア株式会社-Square株式会社様(20230601_20230615)_20230618025830.zip", "unzippedtried")
    .then((files) => {
        console.log(files);
    })
    .catch((error) => {
        console.log(error);
    });

// const fs = require('fs');
//
// var UnzippedFiles = fs.readdirSync('./unzippedtried/');
// var fileList =fs.readdirSync('./unzippedtried/' + UnzippedFiles[0]);
// var filename;
//
// for (let i =0; i<fileList.length; i++){
//     if(files2[i].includes("PayPay(1)")){
//         filename = fileList[i];
//         break;
//     }
// }
// console.log(filename);