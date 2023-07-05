
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());
const xlsx = require('xlsx')
const Tesseract = require('tesseract.js')
const fs = require("fs");

const userid = "square_seisan";
const Password = "Square123##";

const log_in = async () => {
    const browser = await puppeteer.launch({
        headless: false,
        args: ['--no-sandbox', '--disable-gpu', '--enable-webgl', '--window-size=1200,800','--single-process', '--no-zygote',]
    });

    const ua = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36';
    const page = await browser.newPage();

    await page.setDefaultNavigationTimeout(0);
    await page.setViewport({width: 1200, height: 720});
    await page.goto('https://console.onepay.finance/login', {waitUntil: 'networkidle0'}); // wait until page load

    //download image to local and get text from image
    const svgImage = await page.$('#captchaImage');
    await svgImage.screenshot({
        path: 'captcha-screenshot.png', omitBackground: true,
    });

    function getUser() {
        return Tesseract.recognize('./captcha-screenshot.png', 'eng',).then(({data: {text}}) => {
            return text;
        })
    }

    const authCode = await getUser();
    await page.type('#userAccount', userid);
    await page.type('#password', Password);
    await page.type('#captcha', authCode);
    await page.waitForTimeout(6000);
    await page.keyboard.press('Enter');

    await page.goto('https://console.onepay.finance/diffWithTradeAndActuaryManage/clearRecord', {waitUntil: 'domcontentloaded'}); // wait until page load

    await page.waitForTimeout(3000);
    await page.type('#queryNum', "100");

    await page.evaluate(() => {
        search();
    });
    if (await page.$(".dataTables_empty") !== null) {
        console.log('データがありません');
    } else {
        console.log('データがある');
    }
    const found = await page.evaluate(() => window.find("001sqr"));
    if (found) {
        console.log("001sqr is found");
    }
    const result = await page.$$eval('#tbInfo tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            Array.from(columns, column => {
                if (column.innerText[2] == "001sqr") {
                }
            });
            return Array.from(columns, column => column.innerHTML);
        });
    });
    let index = -1;
    for (let i = 0; i < result.length; i++) {
        if (result[i][2] === "001sqr") {
            index = i;
            break;
        }
    }

    console.log(result[index][11]);
    var mailstr = result[index][11];
    var mySubString = mailstr.substring(
        mailstr.indexOf("excelDownLoad(&quot;") + 20,
        mailstr.lastIndexOf("&quot;);return false")
    );
    console.log(mySubString);

    await page.evaluate((mySubString) => { // Downloading Excels
        excelDownLoad(mySubString);
    }, mySubString);


    await page.goto('https://console.onepay.finance/statisticsAnalysisManage/tradeStatistics', {waitUntil: 'domcontentloaded'}); // wait until page load

    await page.waitForTimeout(2000);
    await page.$eval('#agentCodeControl > option', e => e.setAttribute("value", "OA019265"));
    await page.$eval('#payType > option', e => e.setAttribute("value", "10"));
    await page.$eval('#basicDate', el => el.value = '2023-06-15');
    await page.evaluate(() => {
        const checkbox = document.getElementById('isOnlyFlgControl');
        checkbox.checked = false;
    });
    await page.evaluate(() => {
        search();
    });

    await page.waitForTimeout(4000);
    const result2 = await page.$$eval('#tbInfo tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            return Array.from(columns, column => column.innerHTML);
        });
    });

    console.log(result2)
// Load the Excel file
    const workbook = xlsx.readFile('./template/SquareTemplate.xlsx');

// Select the first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const numberFormat = '#,##0';

    worksheet['B8'].v = result2[0][1];
    worksheet['C8'].v = result2[0][2];
    worksheet['D8'].v = result2[0][3];
    worksheet['E8'].v = result2[0][4];
    worksheet['F8'].v = result2[0][5];
    worksheet['G8'].v = result2[0][6];

    // get the information from unzipped
    var UnzippedFiles = fs.readdirSync('./unzippedtried/');
    var fileList =fs.readdirSync('./unzippedtried/' + UnzippedFiles[0]);
    var filename;

    for (let i =0; i<fileList.length; i++){
        if(fileList[i].includes("PayPay(1)")){
            filename = fileList[i];
            break;
        }
    }

    console.log(filename);
    const PayPayWorkbook = xlsx.readFile('./unzippedtried/' + UnzippedFiles[0] + "/" + filename);
    const PayPaySheetName = PayPayWorkbook.SheetNames[0];
    const PayPayWorkSheet = PayPayWorkbook.Sheets[PayPaySheetName];

    worksheet['B56'].v = PayPayWorkSheet['B8'].v;
    console.log('PayPayWorkSheet[\'B8\'].v :',  PayPayWorkSheet['B8'].v);
    worksheet['B57'].v = PayPayWorkSheet['B9'].v;
    console.log('PayPayWorkSheet[\'B9\'].v :', PayPayWorkSheet['B9'].v);
    worksheet['B58'].v = PayPayWorkSheet['B10'].v;
    console.log('PayPayWorkSheet[\'B10\'].v :', PayPayWorkSheet['B10'].v);

    await page.waitForTimeout(2000);

    const calculatedValuetemp = worksheet['E75'].v;
    console.log('E75 :', calculatedValuetemp);

    xlsx.writeFile(workbook, `./template/SquareTemplate.xlsx`);
}

log_in()