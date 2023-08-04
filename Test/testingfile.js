var path = require('path');

var url = 'D:\\2023-07-21\\Download30\\abcdef.zip';

// Use the path.basename() method to extract the last part of the path (filename)
var filename = path.basename(url);

console.log(filename); // Output: abcdef.zip