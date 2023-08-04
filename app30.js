const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());

const decompress = require("decompress");
const xlsx = require('xlsx')
const Tesseract = require('tesseract.js')
const fs = require("fs");
const {spawn} = require('child_process');
const f = require('./function.js')
var path = require('path');
const ChatBot = require('dingtalk-robot-sender');

const SFTPClient = require("ssh2-sftp-client");
const robot = new ChatBot({
    webhook: 'https://oapi.dingtalk.com/robot/send?access_token=290fc32bb41a5326ad0e44fa59ee2f7c60b58f40820a904268665656ab025f73'
});

const userid = "square_seisan";
const Password = "Square123##";
var base_dir = "D:/";
// var base_dir = "/home/onepay/Zip/" // use this on server site
var workBookTemplateLink = './template/SquareTemplate30.xlsx';
// var workBookTemplateLink = '/home/onepay/SquareCalcu/template/SquareTemplate30.xlsx'; // use this on server site

const SFTPhost= "192.168.0.77";
const SFTPport ='2836';
const SFTPUserName= "onepay";
const SFTPpassword= "onepay001";

var CalculationDate;
var CalculationDate30;
const log_in = async () => {
    // date calculation
    const today = new Date();
    // Check if today is after the 18th of the month -> for day 15


    // Check if today is after day 3 of the month -> for day 30
    if (today.getDate() >= 3) {
        var previousMonth = new Date(today.getFullYear(), today.getMonth(), 0);
        CalculationDate30 = previousMonth.getFullYear() + '-' + (previousMonth.getMonth() + 1) + '-' + previousMonth.getDate();
    } else {
        var previousMonth = new Date(today.getFullYear(), today.getMonth() - 1, 0);
        CalculationDate30 = previousMonth.getFullYear() + '-' + (previousMonth.getMonth() + 1) + '-' + previousMonth.getDate();
    }
    let currentCalculationDate30 = f.formatDateSlash(CalculationDate30);

    let dummyDate = new Date(CalculationDate30);
    if (dummyDate.getDate() >= 18) {
        CalculationDate = new Date(dummyDate.getFullYear(), dummyDate.getMonth(), 15);
    } else {
        const previousMonth = dummyDate.getMonth() - 1;
        CalculationDate = new Date(dummyDate.getFullYear(), previousMonth, 15);
    }
    let currentCalculationDate = f.formatDateSlash(CalculationDate);

    console.log("30 Calculation Date: " +currentCalculationDate30);
    console.log("15 Calculation Date: " +currentCalculationDate);

    const browser = await puppeteer.launch({
        headless: false,
        // headless: true, // use this on server site
        args: ['--no-sandbox', '--disable-gpu', '--enable-webgl', '--window-size=1200,800', '--single-process', '--no-zygote',]
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

    await page.waitForTimeout(3000);

    const authCode = await getUser();
    await page.type('#userAccount', userid);
    await page.waitForTimeout(2000);
    await page.type('#password', Password);
    await page.waitForTimeout(2000);
    await page.type('#captcha', authCode);
    await page.waitForTimeout(2000);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);

    await page.goto('https://console.onepay.finance/diffWithTradeAndActuaryManage/clearRecord', {waitUntil: 'domcontentloaded'}); // wait until page load

    await page.waitForTimeout(3000);
    await page.type('#queryNum', "100");

    await page.$eval('#clearDate', (el, date) => {
        el.value = date;
    }, currentCalculationDate30);

    await page.waitForTimeout(1000);
    // await page.$eval('#clearDate', el => el.setAttribute("value", "2023/06/15"));

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
    const downloadPath = path.resolve(base_dir + f.formatDate(new Date()) + "/Download30");
    const client = await page.target().createCDPSession()
    await client.send('Page.setDownloadBehavior', {
        behavior: 'allow',
        downloadPath: downloadPath,
    })
    // Ready to download
    await page.waitForTimeout(5000);
    await page.evaluate((mySubString) => { // Downloading Excels
        excelDownLoad(mySubString);
        console.log("Downloading Excels . . . ");
    }, mySubString);
    await page.waitForTimeout(10000);
    // End of download data
    // Decompress the downloaded file
    var DownloadedFilePath = fs.readdirSync(downloadPath);
    var downloadedZipPath = downloadPath + '/' + DownloadedFilePath;
    let fileToUpload = downloadedZipPath;

    var UnZippedPath = downloadPath + '/unzipped';
    decompress(downloadedZipPath, UnZippedPath)
        .then((files) => {
            console.log(files);
        })
        .catch((error) => {
            console.log(error);
        });

    // Load the Excel file
    const workbook = xlsx.readFile(workBookTemplateLink);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Day 15
    worksheet['D1'].v = currentCalculationDate.toString();
    worksheet['D11'].v = currentCalculationDate.toString();

    // Day30
    worksheet['H1'].v = currentCalculationDate30.toString();
    worksheet['H11'].v = currentCalculationDate30.toString();

    // Day 15 will start here
    // start of adding Sqr001 infos Day 15
    await page.goto('https://console.onepay.finance/statisticsAnalysisManage/tradeStatistics', {waitUntil: 'domcontentloaded'}); // wait until page load
    await page.waitForTimeout(2000);
    await page.$eval('#agentCodeControl > option', e => e.setAttribute("value", "OA019265"));
    await page.$eval('#payType > option', e => e.setAttribute("value", "10"));

    await page.$eval('#basicDate', (el, date) => {
        el.value = date;
    }, currentCalculationDate);

    await page.evaluate(() => {
        const checkbox = document.getElementById('isOnlyFlgControl');
        checkbox.checked = false;
    });
    await page.evaluate(() => {
        search();
    });

    await page.waitForTimeout(4000);
    const result001sqr = await page.$$eval('#tbInfo tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            return Array.from(columns, column => column.innerHTML);
        });
    });

    console.log(result001sqr)
    worksheet['B1'].v = "後半"
    worksheet['D1'].v = currentCalculationDate.toString();
    worksheet['B11'].v = "後半"
    worksheet['D11'].v = currentCalculationDate.toString();

    // Sqr001 Count and Yen
    let Sqr001UriageCount = parseInt(result001sqr[0][3].toString().replace(/,/g, ''));
    worksheet['B5'].v = Sqr001UriageCount;
    let Sqr001UriageYen = parseInt(result001sqr[0][4].toString().replace(/,/g, ''));
    worksheet['C5'].v = Sqr001UriageYen;

    let Sqr001HenKinCount = parseInt(result001sqr[1][3].toString().replace(/,/g, ''));
    worksheet['B6'].v = Sqr001HenKinCount;
    let Sqr001HenKinYen = parseInt(result001sqr[1][4].toString().replace(/,/g, ''));
    worksheet['C6'].v = Sqr001HenKinYen;

    let Sqr001TorikeshiCount = parseInt(result001sqr[2][3].toString().replace(/,/g, ''));
    worksheet['B7'].v = Sqr001TorikeshiCount;
    let Sqr001TorikeshiYen = parseInt(result001sqr[2][4].toString().replace(/,/g, ''));
    worksheet['C7'].v = Sqr001TorikeshiYen;

    let Sqr001UketoriCount = parseInt(result001sqr[4][3].toString().replace(/,/g, ''));
    worksheet['B8'].v = Sqr001UketoriCount;
    let Sqr001UketoriYen = parseInt(result001sqr[4][4].toString().replace(/,/g, ''));
    worksheet['C8'].v = Sqr001UketoriYen;

    let Sqr001ShiharaikingakuCount = parseInt(result001sqr[5][3].toString().replace(/,/g, ''));
    worksheet['B9'].v = Sqr001ShiharaikingakuCount;
    let Sqr001ShiharaikingakuYen = parseInt(result001sqr[5][4].toString().replace(/,/g, ''));
    worksheet['C9'].v = Sqr001ShiharaikingakuYen;
    // end of adding Sqr001 infos Day 15

    // start of adding Sqr001 infos Day 30
    await page.goto('https://console.onepay.finance/statisticsAnalysisManage/tradeStatistics', {waitUntil: 'domcontentloaded'}); // wait until page load
    await page.waitForTimeout(2000);
    await page.$eval('#agentCodeControl > option', e => e.setAttribute("value", "OA019265"));
    await page.$eval('#payType > option', e => e.setAttribute("value", "10"));

    await page.$eval('#basicDate', (el, date) => {
        el.value = date;
    }, currentCalculationDate30);

    await page.evaluate(() => {
        const checkbox = document.getElementById('isOnlyFlgControl');
        checkbox.checked = false;
    });
    await page.evaluate(() => {
        search();
    });

    await page.waitForTimeout(4000);
    const result001sqr30 = await page.$$eval('#tbInfo tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            return Array.from(columns, column => column.innerHTML);
        });
    });

    console.log(result001sqr30)

    // Sqr001 Count and Yen
    let Sqr001UriageCount30 = parseInt(result001sqr30[0][3].toString().replace(/,/g, ''));
    worksheet['F5'].v = Sqr001UriageCount30;
    let Sqr001UriageYen30 = parseInt(result001sqr30[0][4].toString().replace(/,/g, ''));
    worksheet['G5'].v = Sqr001UriageYen30;

    let Sqr001HenKinCount30 = parseInt(result001sqr30[1][3].toString().replace(/,/g, ''));
    worksheet['F6'].v = Sqr001HenKinCount30;
    let Sqr001HenKinYen30 = parseInt(result001sqr30[1][4].toString().replace(/,/g, ''));
    worksheet['G6'].v = Sqr001HenKinYen30;

    let Sqr001TorikeshiCount30 = parseInt(result001sqr30[2][3].toString().replace(/,/g, ''));
    worksheet['F7'].v = Sqr001TorikeshiCount30;
    let Sqr001TorikeshiYen30 = parseInt(result001sqr30[2][4].toString().replace(/,/g, ''));
    worksheet['G7'].v = Sqr001TorikeshiYen30;

    let Sqr001UketoriCount30 = parseInt(result001sqr30[4][3].toString().replace(/,/g, ''));
    worksheet['F8'].v = Sqr001UketoriCount30;
    let Sqr001UketoriYen30 = parseInt(result001sqr30[4][4].toString().replace(/,/g, ''));
    worksheet['G8'].v = Sqr001UketoriYen30;

    let Sqr001ShiharaikingakuCount30 = parseInt(result001sqr30[5][3].toString().replace(/,/g, ''));
    worksheet['F9'].v = Sqr001ShiharaikingakuCount30;
    let Sqr001ShiharaikingakuYen30 = parseInt(result001sqr30[5][4].toString().replace(/,/g, ''));
    worksheet['G9'].v = Sqr001ShiharaikingakuYen30;
    // end of adding Sqr001 infos Day 30

    // start of Sqr001 Test infos Day 15
    await page.goto('https://console.onepay.finance/statisticsAnalysisManage/tradeStatistics', {waitUntil: 'domcontentloaded'}); // wait until page load

    await page.waitForTimeout(2000);
    await page.$eval('#agentCodeControl > option', e => e.setAttribute("value", "OA011675"));
    await page.$eval('#payType > option', e => e.setAttribute("value", "10"));
    await page.$eval('#basicDate', (el, date) => {
        el.value = date;
    }, currentCalculationDate);

    await page.$eval('#branchCodeControl > option', e => e.setAttribute("value", "SQ0999999999"));
    await page.evaluate(() => {
        const checkbox = document.getElementById('isOnlyFlgControl');
        checkbox.checked = false;
    });
    await page.evaluate(() => {
        search();
    });
    await page.waitForTimeout(4000);
    const result001sqrtest = await page.$$eval('#tbInfo tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            return Array.from(columns, column => column.innerHTML);
        });
    });
    console.log(result001sqrtest)

    // Sqr001 Test Count and Yen
    let Sqr001TestUriageCount = parseInt(result001sqrtest[0][3].toString().replace(/,/g, ''));
    worksheet['B15'].v = Sqr001TestUriageCount;
    let Sqr001TestUriageYen = parseInt(result001sqrtest[0][4].toString().replace(/,/g, ''));
    worksheet['C15'].v = Sqr001TestUriageYen;

    let Sqr001TestHenKinCount = parseInt(result001sqrtest[1][3].toString().replace(/,/g, ''));
    worksheet['B16'].v = Sqr001TestHenKinCount;
    let Sqr001TestHenKinYen = parseInt(result001sqrtest[1][4].toString().replace(/,/g, ''));
    worksheet['C16'].v = Sqr001TestHenKinYen;

    let Sqr001TestTorikeshiCount = parseInt(result001sqrtest[2][3].toString().replace(/,/g, ''));
    worksheet['B17'].v = Sqr001TestTorikeshiCount;
    let Sqr001TestTorikeshiYen = parseInt(result001sqrtest[2][4].toString().replace(/,/g, ''));
    worksheet['C17'].v = Sqr001TestTorikeshiYen;

    let Sqr001TestUketoriCount = parseInt(result001sqrtest[4][3].toString().replace(/,/g, ''));
    worksheet['B18'].v = Sqr001TestUketoriCount;
    let Sqr001TestUketoriYen = parseInt(result001sqrtest[4][4].toString().replace(/,/g, ''));
    worksheet['C18'].v = Sqr001TestUketoriYen;

    let Sqr001TestShiharaikingakuCount = parseInt(result001sqrtest[5][3].toString().replace(/,/g, ''));
    worksheet['B19'].v = Sqr001TestShiharaikingakuCount;
    let Sqr001TestShiharaikingakuYen = parseInt(result001sqrtest[5][4].toString().replace(/,/g, ''));
    worksheet['C19'].v = Sqr001TestShiharaikingakuYen;
    // end of adding Sqr001 Test infos Day 15

    // start of adding Sqr001 Test infos Day 30
    await page.goto('https://console.onepay.finance/statisticsAnalysisManage/tradeStatistics', {waitUntil: 'domcontentloaded'}); // wait until page load

    await page.waitForTimeout(2000);
    await page.$eval('#agentCodeControl > option', e => e.setAttribute("value", "OA011675"));
    await page.$eval('#payType > option', e => e.setAttribute("value", "10"));
    await page.$eval('#basicDate', (el, date) => {
        el.value = date;
    }, currentCalculationDate30);

    await page.$eval('#branchCodeControl > option', e => e.setAttribute("value", "SQ0999999999"));
    await page.evaluate(() => {
        const checkbox = document.getElementById('isOnlyFlgControl');
        checkbox.checked = false;
    });
    await page.evaluate(() => {
        search();
    });
    await page.waitForTimeout(4000);
    const result001sqrtest30 = await page.$$eval('#tbInfo tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            return Array.from(columns, column => column.innerHTML);
        });
    });
    console.log(result001sqrtest30)

    // Sqr001 Test Count and Yen
    let Sqr001TestUriageCount30 = parseInt(result001sqrtest30[0][3].toString().replace(/,/g, ''));
    worksheet['F15'].v = Sqr001TestUriageCount30;
    let Sqr001TestUriageYen30 = parseInt(result001sqrtest30[0][4].toString().replace(/,/g, ''));
    worksheet['G15'].v = Sqr001TestUriageYen30;

    let Sqr001TestHenKinCount30 = parseInt(result001sqrtest30[1][3].toString().replace(/,/g, ''));
    worksheet['F16'].v = Sqr001TestHenKinCount30;
    let Sqr001TestHenKinYen30 = parseInt(result001sqrtest30[1][4].toString().replace(/,/g, ''));
    worksheet['G16'].v = Sqr001TestHenKinYen30;

    let Sqr001TestTorikeshiCount30 = parseInt(result001sqrtest30[2][3].toString().replace(/,/g, ''));
    worksheet['F17'].v = Sqr001TestTorikeshiCount30;
    let Sqr001TestTorikeshiYen30 = parseInt(result001sqrtest30[2][4].toString().replace(/,/g, ''));
    worksheet['G17'].v = Sqr001TestTorikeshiYen30;

    let Sqr001TestUketoriCount30 = parseInt(result001sqrtest30[4][3].toString().replace(/,/g, ''));
    worksheet['F18'].v = Sqr001TestUketoriCount30;
    let Sqr001TestUketoriYen30 = parseInt(result001sqrtest30[4][4].toString().replace(/,/g, ''));
    worksheet['G18'].v = Sqr001TestUketoriYen30;

    let Sqr001TestShiharaikingakuCount30 = parseInt(result001sqrtest30[5][3].toString().replace(/,/g, ''));
    worksheet['F19'].v = Sqr001TestShiharaikingakuCount30;
    let Sqr001TestShiharaikingakuYen30 = parseInt(result001sqrtest30[5][4].toString().replace(/,/g, ''));
    worksheet['G19'].v = Sqr001TestShiharaikingakuYen30;
    // end of adding Sqr001 Test infos Day 30


    // start here Day 15 -> Calculation Real Data - Test Data
    let Sqr001Day15CombineUriageCount = Sqr001UriageCount - Sqr001TestUriageCount;
    worksheet['B23'].v = Sqr001Day15CombineUriageCount;
    let Sqr001Day15CombineUriageYen = Sqr001UriageYen - Sqr001TestUriageYen;
    worksheet['C23'].v = Sqr001Day15CombineUriageYen;

    let Sqr001Day15CombineHenKinCount = Sqr001HenKinCount - Sqr001TestHenKinCount;
    worksheet['B24'].v = Sqr001Day15CombineHenKinCount;
    let Sqr001Day15CombineHenKinYen = (Sqr001HenKinYen - Sqr001TestHenKinYen) * -1;
    worksheet['C24'].v = Sqr001Day15CombineHenKinYen;

    let Sqr001Day15CombineTorikeshiCount = Sqr001TorikeshiCount - Sqr001TestTorikeshiCount;
    worksheet['B25'].v = Sqr001Day15CombineTorikeshiCount;
    let Sqr001Day15CombineTorikeshiYen = Sqr001TorikeshiYen - Sqr001TestTorikeshiYen;
    worksheet['C25'].v = Sqr001Day15CombineTorikeshiYen;

    let Sqr001Day15CombineUketoriCount = Sqr001UketoriCount - Sqr001TestUketoriCount;
    worksheet['B26'].v = Sqr001Day15CombineUketoriCount;
    let Sqr001Day15CombineUketoriYen = Sqr001UketoriYen - Sqr001TestUketoriYen;
    worksheet['C26'].v = Sqr001Day15CombineUketoriYen;

    let Sqr001Day15CombineShiharaikingakuCount = Sqr001ShiharaikingakuCount - Sqr001TestShiharaikingakuCount;
    worksheet['B27'].v = Sqr001Day15CombineShiharaikingakuCount;
    let Sqr001Day15CombineShiharaikingakuYen = Sqr001ShiharaikingakuYen - Sqr001TestShiharaikingakuYen;
    worksheet['C27'].v = Sqr001Day15CombineShiharaikingakuYen;
    // end here Day 15 -> Calculation Real Data - Test Data

    // start here Day 30 -> Calculation Real Data - Test Data
    let Sqr001Day15CombineUriageCount30 = Sqr001UriageCount30 - Sqr001TestUriageCount30;
    worksheet['F23'].v = Sqr001Day15CombineUriageCount30;
    let Sqr001Day15CombineUriageYen30 = Sqr001UriageYen30 - Sqr001TestUriageYen30;
    worksheet['G23'].v = Sqr001Day15CombineUriageYen30;

    let Sqr001Day15CombineHenKinCount30 = Sqr001HenKinCount30 - Sqr001TestHenKinCount30;
    worksheet['F24'].v = Sqr001Day15CombineHenKinCount30;
    let Sqr001Day15CombineHenKinYen30 = (Sqr001HenKinYen30 - Sqr001TestHenKinYen30) * -1;
    worksheet['G24'].v = Sqr001Day15CombineHenKinYen30;

    let Sqr001Day15CombineTorikeshiCount30 = Sqr001TorikeshiCount30 - Sqr001TestTorikeshiCount30;
    worksheet['F25'].v = Sqr001Day15CombineTorikeshiCount30;
    let Sqr001Day15CombineTorikeshiYen30 = Sqr001TorikeshiYen30 - Sqr001TestTorikeshiYen30;
    worksheet['G25'].v = Sqr001Day15CombineTorikeshiYen30;

    let Sqr001Day15CombineUketoriCount30 = Sqr001UketoriCount30 - Sqr001TestUketoriCount30;
    worksheet['F26'].v = Sqr001Day15CombineUketoriCount30;
    let Sqr001Day15CombineUketoriYen30 = Sqr001UketoriYen30 - Sqr001TestUketoriYen30;
    worksheet['G26'].v = Sqr001Day15CombineUketoriYen30;

    let Sqr001Day15CombineShiharaikingakuCount30 = Sqr001ShiharaikingakuCount30 - Sqr001TestShiharaikingakuCount30;
    worksheet['F27'].v = Sqr001Day15CombineShiharaikingakuCount30;
    let Sqr001Day15CombineShiharaikingakuYen30 = Sqr001ShiharaikingakuYen30 - Sqr001TestShiharaikingakuYen30;
    worksheet['G27'].v = Sqr001Day15CombineShiharaikingakuYen30;
    // end here Day 30 -> Calculation Real Data - Test Data

    //Start Diff Calculation
    let DiffSqr001Day15CombineUriageCount30 = Sqr001Day15CombineUriageCount30 - Sqr001Day15CombineUriageCount;
    worksheet['J23'].v = DiffSqr001Day15CombineUriageCount30;
    let DiffSqr001Day15CombineUriageYen30 = Sqr001Day15CombineUriageYen30 - Sqr001Day15CombineUriageYen;
    worksheet['K23'].v = DiffSqr001Day15CombineUriageYen30;

    let DiffSqr001Day15CombineHenKinCount30 = Sqr001Day15CombineHenKinCount30 - Sqr001Day15CombineHenKinCount;
    worksheet['J24'].v = DiffSqr001Day15CombineHenKinCount30;
    let DiffSqr001Day15CombineHenKinYen30 = Sqr001Day15CombineHenKinYen30 - Sqr001Day15CombineHenKinYen;
    worksheet['K24'].v = DiffSqr001Day15CombineHenKinYen30;

    let DiffSqr001Day15CombineTorikeshiCount30 = Sqr001Day15CombineTorikeshiCount30 - Sqr001Day15CombineTorikeshiCount;
    worksheet['J25'].v = DiffSqr001Day15CombineTorikeshiCount30;
    let DiffSqr001Day15CombineTorikeshiYen30 = Sqr001Day15CombineTorikeshiYen30 - Sqr001Day15CombineTorikeshiYen;
    worksheet['K25'].v = DiffSqr001Day15CombineTorikeshiYen30;

    let DiffSqr001Day15CombineUketoriCount30 = Sqr001Day15CombineUketoriCount30 - Sqr001Day15CombineUketoriCount;
    worksheet['J26'].v = DiffSqr001Day15CombineUketoriCount30;
    let DiffSqr001Day15CombineUketoriYen30 = Sqr001Day15CombineUketoriYen30 - Sqr001Day15CombineUketoriYen;
    worksheet['K26'].v = DiffSqr001Day15CombineUketoriYen30;

    let DiffSqr001Day15CombineShiharaikingakuCount30 = Sqr001Day15CombineShiharaikingakuCount30 - Sqr001Day15CombineShiharaikingakuCount;
    worksheet['J27'].v = DiffSqr001Day15CombineShiharaikingakuCount30;
    let DiffSqr001Day15CombineShiharaikingakuYen30 = Sqr001Day15CombineShiharaikingakuYen30 - Sqr001Day15CombineShiharaikingakuYen;
    worksheet['K27'].v = DiffSqr001Day15CombineShiharaikingakuYen30;
    //End Diff Calculation

    // Compare
    let CompareUriagekensuTotalling = DiffSqr001Day15CombineUriageCount30;
    worksheet['B48'].v = CompareUriagekensuTotalling;
    let CompareUriageTotalling = DiffSqr001Day15CombineUriageYen30;
    worksheet['B49'].v = CompareUriageTotalling;
    let CompareHenkinkensuTotalling = DiffSqr001Day15CombineHenKinCount30;
    worksheet['B50'].v = CompareHenkinkensuTotalling;
    let CompareHenkinTotalling = DiffSqr001Day15CombineHenKinYen30;
    worksheet['B51'].v = CompareHenkinTotalling;
    let CompareUketorikensuTotalling = DiffSqr001Day15CombineUketoriYen30;
    worksheet['B52'].v = CompareUketorikensuTotalling;
    let CompareToriatsukaikoTotalling = DiffSqr001Day15CombineUriageYen30 + DiffSqr001Day15CombineHenKinYen30;
    worksheet['B53'].v = CompareToriatsukaikoTotalling;
    let CompareShiharaikingakuTotalling = DiffSqr001Day15CombineShiharaikingakuYen30;
    worksheet['B54'].v = CompareShiharaikingakuTotalling;

    // get the information from unzipped
    var UnzippedFiles = fs.readdirSync(UnZippedPath);
    var fileList = fs.readdirSync(UnZippedPath + '/' + UnzippedFiles[0]);
    var filename;
    for (let i = 0; i < fileList.length; i++) {
        if (fileList[i].includes("PayPay(1)")) {
            filename = fileList[i];
            break;
        }
    }

    // Open Pay Pay Work Book
    const PayPayWorkbook = xlsx.readFile(UnZippedPath + '/' + UnzippedFiles[0] + "/" + filename);
    const PayPaySheetName = PayPayWorkbook.SheetNames[0];
    const PayPayWorkSheet = PayPayWorkbook.Sheets[PayPaySheetName];

    // 締め日    支払日    手数料率(%)
    worksheet['B34'].v = PayPayWorkSheet['B8'].v.toString();
    worksheet['B35'].v = PayPayWorkSheet['B9'].v.toString();
    worksheet['B36'].v = PayPayWorkSheet['B10'].v;

    // 精算金額(円）手数料総額（円）取扱高（円）消費税（10%）（円）
    let PayPaySeisanKingaku = PayPayWorkSheet['B14'].v;
    worksheet['B39'].v = PayPaySeisanKingaku;
    let PayPayTesuryoSogaku = PayPayWorkSheet['B15'].v;
    worksheet['B40'].v = PayPayTesuryoSogaku;
    let PayPayToriatsukaiko = PayPayWorkSheet['B16'].v;
    worksheet['B41'].v = PayPayToriatsukaiko;
    let PayPayShohizei = PayPayWorkSheet['D15'].v;
    worksheet['D40'].v = PayPayShohizei;

    let PayPayShiharaiKingaku = PayPayWorkSheet['B19'].v;
    worksheet['B43'].v = PayPayShiharaiKingaku;
    let PayPayHankinKingaku = PayPayWorkSheet['B20'].v;
    worksheet['B44'].v = PayPayHankinKingaku;
    let PayPayShiharaiKensu = PayPayWorkSheet['D19'].v;
    worksheet['D43'].v = PayPayShiharaiKensu;
    let PayPayHankinKensu = PayPayWorkSheet['D20'].v;
    worksheet['D44'].v = PayPayHankinKensu;
    let PayPayShiharaiTesuryo = PayPayWorkSheet['F19'].v;
    worksheet['F43'].v = PayPayShiharaiTesuryo;
    let PayPayHankinTesuryo = PayPayWorkSheet['F20'].v;
    worksheet['F44'].v = PayPayHankinTesuryo;

    let CompareUriagekensuSalesSlip = PayPayShiharaiKensu;
    worksheet['C48'].v = CompareUriagekensuSalesSlip;
    let CompareUriageSalesSlip = PayPayShiharaiKingaku;
    worksheet['C49'].v = CompareUriageSalesSlip;
    let CompareHenkinkensuSalesSlip = PayPayHankinKensu;
    worksheet['C50'].v = CompareHenkinkensuSalesSlip;
    let CompareHenkinSalesSlip = PayPayHankinKingaku;
    worksheet['C51'].v = CompareHenkinSalesSlip;
    let CompareUketorikensuSalesSlip = PayPayShiharaiTesuryo + PayPayHankinTesuryo;
    worksheet['C52'].v = CompareUketorikensuSalesSlip;
    let CompareToriatsukaikoSalesSlip = PayPayShiharaiKingaku + PayPayHankinKingaku;
    worksheet['C53'].v = CompareToriatsukaikoSalesSlip;
    let CompareShiharaikingakuSalesSlip = PayPayToriatsukaiko - PayPayTesuryoSogaku;
    worksheet['C54'].v = CompareShiharaikingakuSalesSlip;

    let TFResultUriagekensuSalesSlip = (CompareUriagekensuTotalling == CompareUriagekensuSalesSlip);
    worksheet['E48'].v = TFResultUriagekensuSalesSlip.toString();
    let TFResultUriageSalesSlip = (CompareUriageTotalling == CompareUriageSalesSlip);
    worksheet['E49'].v = TFResultUriageSalesSlip.toString();
    let TFResultHenkinkensuSalesSlip = (CompareHenkinkensuTotalling == CompareHenkinkensuSalesSlip);
    worksheet['E50'].v = TFResultHenkinkensuSalesSlip.toString();
    let TFResultHenkinSalesSlip = (CompareHenkinTotalling == CompareHenkinSalesSlip);
    worksheet['E51'].v = TFResultHenkinSalesSlip.toString();
    let TFResultUketorikensuSalesSlip = (CompareUketorikensuTotalling == CompareUketorikensuSalesSlip);
    worksheet['E52'].v = TFResultUketorikensuSalesSlip.toString();
    let TFResultToriatsukaikoSalesSlip = (CompareToriatsukaikoTotalling == CompareToriatsukaikoSalesSlip);
    worksheet['E53'].v = TFResultToriatsukaikoSalesSlip.toString();
    let TFResultShiharaikingakuSalesSlip = (CompareShiharaikingakuTotalling == CompareShiharaikingakuSalesSlip);
    worksheet['E54'].v = TFResultShiharaikingakuSalesSlip.toString();

    let DifResultUriagekensuSalesSlip = (CompareUriagekensuTotalling - CompareUriagekensuSalesSlip);
    worksheet['D48'].v = DifResultUriagekensuSalesSlip;
    let DifResultUriageSalesSlip = (CompareUriageTotalling - CompareUriageSalesSlip);
    worksheet['D49'].v = DifResultUriageSalesSlip;
    let DifResultHenkinkensuSalesSlip = (CompareHenkinkensuTotalling - CompareHenkinkensuSalesSlip);
    worksheet['D50'].v = DifResultHenkinkensuSalesSlip;
    let DifResultHenkinSalesSlip = (CompareHenkinTotalling - CompareHenkinSalesSlip);
    worksheet['D51'].v = DifResultHenkinSalesSlip;
    let DifResultUketorikensuSalesSlip = (CompareUketorikensuTotalling - CompareUketorikensuSalesSlip);
    worksheet['D52'].v = DifResultUketorikensuSalesSlip;
    let DifResultToriatsukaikoSalesSlip = (CompareToriatsukaikoTotalling - CompareToriatsukaikoSalesSlip);
    worksheet['D53'].v = DifResultToriatsukaikoSalesSlip;
    let DifResultShiharaikingakuSalesSlip = (CompareShiharaikingakuTotalling - CompareShiharaikingakuSalesSlip);
    worksheet['D54'].v = DifResultShiharaikingakuSalesSlip;

    let reportText = (CalculationDate.getMonth() + 1) + "月後半分の";
    
        if (TFResultUriagekensuSalesSlip && TFResultUriageSalesSlip && TFResultHenkinkensuSalesSlip && TFResultHenkinSalesSlip
        && TFResultUketorikensuSalesSlip && TFResultToriatsukaikoSalesSlip && TFResultShiharaikingakuSalesSlip) {
            reportText += "計算に問題はありません。結果は以下の通りです。\n\n";
        }
        else{
            reportText += "計算に問題はあります。結果は以下の通りです。\n\n";
        }

        reportText += "集計 - 売上件数: "+ CompareUriagekensuTotalling +"、";
        reportText += "売上票 - 売上件数: "+ CompareUriagekensuSalesSlip +"、";
        reportText += "差異 - 売上件数: "+ DifResultUriagekensuSalesSlip +"\n";

        reportText += "集計 - 売上: "+ CompareUriageTotalling +"、";
        reportText += "売上票 - 売上: "+ CompareUriageSalesSlip +"、";
        reportText += "差異 - 売上: "+ DifResultUriageSalesSlip +"\n\n";

        reportText += "集計 - 返金件数: "+ CompareHenkinkensuTotalling +"、";
        reportText += "売上票 - 返金件数: "+ CompareHenkinkensuSalesSlip +"、";
        reportText += "差異 - 返金件数: "+ DifResultHenkinkensuSalesSlip +"\n";

        reportText += "集計 - 返金: "+ CompareHenkinTotalling +"、";
        reportText += "売上票 - 返金: "+ CompareHenkinSalesSlip +"、";
        reportText += "差異 - 返金: "+ DifResultHenkinSalesSlip +"\n\n";

        reportText += "集計 - 手数料総額: "+ CompareUketorikensuTotalling +"、";
        reportText += "売上票 - 手数料総額: "+ CompareUketorikensuSalesSlip +"、";
        reportText += "差異 - 手数料総額: "+ DifResultUketorikensuSalesSlip +"\n\n";

        reportText += "集計 - 取扱高: "+ CompareToriatsukaikoTotalling +"、";
        reportText += "売上票 - 取扱高: "+ CompareToriatsukaikoSalesSlip +"、";
        reportText += "差異 - 取扱高: "+ DifResultToriatsukaikoSalesSlip +"\n\n";

        reportText += "集計 - 支払金額: "+ CompareShiharaikingakuTotalling +"、";
        reportText += "売上票 - 支払金額: "+ CompareShiharaikingakuSalesSlip +"、";
        reportText += "差異 - 支払金額: "+ DifResultShiharaikingakuSalesSlip +"\n\n";
        

    reportText+="ファイルは以下のSFTPパスにアップロードしました。\nSFTP 192.168.0.183\n";
    reportText+="/home/onepay/Zip/" + f.formatDate(new Date()) +"/Download30";

    robot.text("Square 計算からのレポート\n\n" + reportText);
    console.log("Square 計算からのレポート\n\n" + reportText);

    await page.waitForTimeout(2000);
    let outputFileName = "Square"+(CalculationDate.getMonth() + 1)+"月夜半精算結果確認.xlsx"
    xlsx.writeFile(workbook, downloadPath + "/"+outputFileName);
    // Square7月後半精算結果確認

    // Upload to SFTP
    console.log(downloadedZipPath + " :downloadedZipPath 2 here ");
    console.log(fileToUpload + " :fileToUpload 2 here");
    var theZipFileName = path.basename(fileToUpload);

    // let makeDirectory1 = "/OA011675/日本恒生ソフトウェア株式会社/" + f.formatDateNull(today);
    // let makeDirectory2 = "/OA011675/日本恒生ソフトウェア株式会社/" + f.formatDateNull(today)+'/settlement/';
    let makeDirectory1 = "/Public/SquareDownload/" + f.formatDateNull(today);
    let makeDirectory2 = "/Public/SquareDownload/" + f.formatDateNull(today)+'/settlement/';

    let fileToLocated = makeDirectory2 + theZipFileName;

    const sftp = new SFTPClient()
    await sftp.connect({
        host: SFTPhost,
        port: SFTPport,
        username: SFTPUserName,
        password: SFTPpassword
    }).then( async () => {
        await sftp.mkdir(makeDirectory1,false);
        await sftp.mkdir(makeDirectory2, false);
        console.log("uploading");
        await sftp.put(fileToUpload, fileToLocated, false);
        console.log("upload complete");
    }).then((data) => {
        console.log(data, 'data upload complete');
    }).catch((err) => {
        console.log(err, 'catch error');
    });
    // Square7月前半精算結果確認
}

log_in()