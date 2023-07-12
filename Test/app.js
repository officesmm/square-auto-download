// const { Keyboard } = require('puppeteer');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());
// const fs = require('fs');
// const path = require('path');e
const xlsx = require('xlsx')
// import xlsx from 'xlsx/xlsx';
// const stringify = require('csv-stringify');
// const ChatBot = require('dingtalk-robot-sender');
const Tesseract = require('tesseract.js')
const fs = require("fs");

const userid = "square_seisan";
const Password = "Square123##";
// const NumStatus = ["99", "01", "02"]
// const NumStatusName = ["ファイル不正", "ダウンロード中", "ダウンロード済み"]
let FinalStatus = {};
let content = "";
// const NumStatus = ["99", "01", "02", "03", "04", "05", "06", "07", "08", "09"]
// const NumStatusName = ["ファイル不正", "ダウンロード中", "ダウンロード済み", "ダウンロード失敗", "インポート中", "インポート済み", "インポート失敗", "取込未了", "取込完了", "取込中止"]
const NumStatus = ["99", "01", "02", "03", "04", "06", "07", "08", "09"]
const NumStatusName = ["ファイル不正", "ダウンロード中", "ダウンロード済み", "ダウンロード失敗", "インポート中", "インポート失敗", "取込未了", "取込完了", "取込中止"]


const log_in = async () => {
    const browser = await puppeteer.launch({
        headless: false,
        args: ['--no-sandbox', '--disable-gpu', '--enable-webgl', '--window-size=1200,800', // '--disable-background-timer-throttling',
            // '--disable-backgrounding-occluded-windows',
            // '--disable-renderer-backgrounding'
            '--single-process', '--no-zygote',]
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
    // await page.goto('https://test.onepay.finance/console/diffWithTradeAndActuaryManage/clearRecord', {waitUntil: 'domcontentloaded'}); // wait until page load

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
    // await page.goto('https://test.onepay.finance/console/statisticsAnalysisManage/tradeStatistics', {waitUntil: 'domcontentloaded'}); // wait until page load

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

    await page.waitForTimeout(10000);
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
    // worksheet['C2'].z = numberFormat;

    // worksheet['B8'].v = "1";
    // worksheet['C8'].v = "2";
    // worksheet['D8'].v = "3";
    // worksheet['E8'].v = "4";
    // worksheet['F8'].v = "5";
    // worksheet['G8'].v = "6";

    worksheet['B8'].v = result2[0][1].toString();
    worksheet.numberFormat('0.00', 'B8');
    //
    worksheet['C8'].v = result2[0][2].toString();
    worksheet['C8'].z = numberFormat;
    worksheet['D8'].v = result2[0][3].toString();
    worksheet['D8'].z = numberFormat;
    worksheet['E8'].v = result2[0][4].toString();
    worksheet['E8'].z = numberFormat;
    worksheet['F8'].v = result2[0][5].toString();
    worksheet['F8'].z = numberFormat;
    worksheet['G8'].v = result2[0][6].toString();
    worksheet['G8'].z = numberFormat;

    // worksheet['B9'].v = result2[1][1].toString();
    // worksheet['C9'].v = result2[1][2].toString();
    // worksheet['D9'].v = result2[1][3].toString();
    // worksheet['E9'].v = result2[1][4].toString();
    // worksheet['F9'].v = result2[1][5].toString();
    // worksheet['G9'].v = result2[1][6].toString();
    //
    // worksheet['B10'].v = result2[2][1].toString();
    // worksheet['C10'].v = result2[2][2].toString();
    // worksheet['D10'].v = result2[2][3].toString();
    // worksheet['E10'].v = result2[2][4].toString();
    // worksheet['F10'].v = result2[2][5].toString();
    // worksheet['G10'].v = result2[2][6].toString();
    //
    // worksheet['B11'].v = result2[3][1].toString();
    // worksheet['C11'].v = result2[3][2].toString();
    // worksheet['D11'].v = result2[3][3].toString();
    // worksheet['E11'].v = result2[3][4].toString();
    // worksheet['F11'].v = result2[3][5].toString();
    // worksheet['G11'].v = result2[3][6].toString();
    //
    // worksheet['B12'].v = result2[4][1].toString();
    // worksheet['C12'].v = result2[4][2].toString();
    // worksheet['D12'].v = result2[4][3].toString();
    // worksheet['E12'].v = result2[4][4].toString();
    // worksheet['F12'].v = result2[4][5].toString();
    // worksheet['G12'].v = result2[4][6].toString();
    //
    // worksheet['B13'].v = result2[5][1].toString();
    // worksheet['C13'].v = result2[5][2].toString();
    // worksheet['D13'].v = result2[5][3].toString();
    // worksheet['E13'].v = result2[5][4].toString();
    // worksheet['F13'].v = result2[5][5].toString();
    // worksheet['G13'].v = result2[5][6].toString();



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
    worksheet['B58'].z = numberFormat;
    console.log('PayPayWorkSheet[\'B10\'].v :', PayPayWorkSheet['B10'].v);

    worksheet['B63'].v = PayPayWorkSheet['B15'].v;
    worksheet.numberFormat('0.00', 'B63');

    await page.waitForTimeout(2000);

    const calculatedValue = worksheet['B31'].v;
    const calculatedValue1 = worksheet['E75'].v;
    const calculatedValue2 = worksheet['E76'].v;
    const calculatedValue3 = worksheet['E81'].v;
    console.log('Calculated Value1:', calculatedValue);
    console.log('Calculated Value2:', calculatedValue1);
    console.log('Calculated Value3:', calculatedValue2);
    console.log('Calculated Value4:', calculatedValue3);

    xlsx.writeFile(workbook, `./template/SquareTemplate.xlsx`);

    // const robot = new ChatBot({
    //     webhook: 'https://oapi.dingtalk.com/robot/send?access_token=79d8377bbd2b6104fbbf34e422f0769cb119f3bfbbd5b5c9828105bedbf824a9'
    // });
    // for (var i = 0; i < NumStatus.length; i++) {
    //     await page.select('select[name="status"]', NumStatus[i])
    //
    //     await page.waitForTimeout(2000);
    //     await page.waitForSelector('[type="submit"]');
    //     await page.click('[type="submit"]')
    //
    //     await page.waitForTimeout(2000);
    //     if (await page.$(".dataTables_empty") !== null) {
    //         // ResultArr.push({[NumStatusName[i]] : NumStatus[i] +'No Data'})
    //         console.log(NumStatusName[i] + ' データがありません');
    //         FinalStatus[NumStatusName[i]] = NumStatus[i] + ' データがありません'
    //         content = NumStatusName[i] + ' データがありません';
    //     } else {
    //         // ResultArr.push({[NumStatusName[i]] : NumStatus[i] +'Exist Data'})
    //         console.log(NumStatusName[i] + 'データがある');
    //         FinalStatus[NumStatusName[i]] = NumStatus[i] + ' データがある'
    //         content = NumStatusName[i] + ' データがある';
    //
    //     }
    //     await page.waitForTimeout(3000);
    //     robot.text(content);
    // }
    //
    // const workbook = xlsx.utils.book_new();
    // const worksheet = xlsx.utils.json_to_sheet([FinalStatus])
    // xlsx.utils.book_append_sheet(workbook, worksheet, "Result");
    // xlsx.writeFile(workbook, `申請システム情報.xlsx`);
    //
    // FinalStatus = {}
    //
    // await page.waitForTimeout(2000);
    //
    // await page.goto('https://test.onepay.finance/console/systemManage/ruleImportRecord', {waitUntil: 'domcontentloaded'}); // wait until page load
    // await page.select('select[name="status"]', "1")
    //
    // await page.waitForTimeout(2000);
    // await page.waitForSelector('[type="submit"]');
    // await page.click('[type="submit"]')
    //
    // if (await page.$(".dataTables_empty") !== null) {
    //
    //     // ResultArr.push({[NumStatusName[i]] : NumStatus[i] +'No Data'})
    //     console.log('失敗 データがありません');
    //     FinalStatus["失敗"] = '失敗 データがありません';
    //     content = '失敗 データがありません';
    //
    // } else {
    //     // ResultArr.push({[NumStatusName[i]] : NumStatus[i] +'Exist Data'})
    //     console.log('失敗 データがある');
    //     FinalStatus["失敗"] = '失敗 データがある';
    //     content = '失敗 データがある';
    //
    // }
    //
    // const workbook1 = xlsx.utils.book_new();
    // const worksheet1 = xlsx.utils.json_to_sheet([FinalStatus])
    // xlsx.utils.book_append_sheet(workbook1, worksheet1, "Result");
    // xlsx.writeFile(workbook1, `インポート履歴.xlsx`);
    // robot.text(content);
    // await browser.close();
}

log_in()