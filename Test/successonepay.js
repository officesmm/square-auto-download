const {Keyboard} = require('puppeteer');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx')
// import xlsx from 'xlsx/xlsx';
// const stringify = require('csv-stringify');
const ChatBot = require('dingtalk-robot-sender');
const Tesseract = require('tesseract.js')
const dotenv = require('dotenv') // .env file
dotenv.config() // using .env

// const userid = "ha_shine";
// const Password = "@Hashine21";
let content = "";


let OnePayConfig = JSON.parse(process.env.BILLMANAGEDEVELOPMENT);

console.log(OnePayConfig.USERID)
// const ConditionOne = [
// '2188120001696976'
// ,'2188120000744447'
// ,'2188120000582348'
// ,'2188120000151270'
// ,'2188120000053556'
// ,'211011000010000163591'
// ,'2188120000611503'
// ,'55vid0zi0irmv4sukb60c2i4vq3u0fu8_BM00003033'
// ,'55vid0zi0irmv4sukb60c2i4vq3u0fu8_BM00003035'
// ,'55vid0zi0irmv4sukb60c2i4vq3u0fu8_BM00003032'
// ]


// const ConditionTwo = [

// '2088141224843610'
// ,'2088331463920238'
// ,'131544548'
// ,'1511090511'
// ,'132886505'
// ,'105342130'
// ,'114713612'
// ,'2188120000608884'
// ,'2188120000215314'
// ,'2188120000615103'
// ,'A111245600000000 (A111245600000200)'
// ,'A111245600000000 (A111245600000300)'
// ,'2088331460085745'
// ,'115050493'
// ,'A111245500000000 (A111245500000200)'
// ,'JHS_gwWechaPay_invalid'
// ]

// const ConditionThree = [
//   '2188120000005294'
// ]

const log_in = async () => {
    const browser = await puppeteer.launch({
        headless: false,
        args: [
            '--no-sandbox',
            '--disable-gpu',
            '--enable-webgl',
            '--window-size=1200,800',
            '--single-process', '--no-zygote',
        ]
    });


    const ua = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36';
    const page = await browser.newPage();

    await page.setDefaultNavigationTimeout(0);
    await page.setViewport({width: 1200, height: 720});
    await page.goto(OnePayConfig.LOGIN_URL, {waitUntil: 'networkidle0'}); // wait until page load


//download image to local and get text from image
    const svgImage = await page.$('#captchaImage');
    await svgImage.screenshot({
        path: 'captcha-screenshot.png',
        omitBackground: true,
    });


    function getUser() {
        return Tesseract.recognize(
            './captcha-screenshot.png',
            'eng',
        ).then(({data: {text}}) => {
            return text;
        })

    }

    const authCode = await getUser();
//download image to local and get text from image

    await page.type('#userAccount', OnePayConfig.USERID);
    await page.type('#password', OnePayConfig.PASSWORD);
    await page.type('#captcha', authCode);
    await page.waitForTimeout(6000);
    await page.keyboard.press('Enter');

    // await page.waitForSelector('[id="loginSubmit"]');
    //await page.click('[id="loginSubmit"]')

    await page.goto(OnePayConfig.BATCHJOBMANAGE_URL, {waitUntil: 'domcontentloaded'}); // wait until page load

    const robot = new ChatBot({
        webhook: 'https://oapi.dingtalk.com/robot/send?access_token=79d8377bbd2b6104fbbf34e422f0769cb119f3bfbbd5b5c9828105bedbf824a9'
    });
    await page.waitForTimeout(6000);

    const result = await page.$$eval('#tbInfo tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            return Array.from(columns, column => column.innerText);
        });
    });

    // console.log(result);
    const success = await checkSuccess(result);

    function checkSuccess(result) {
        for (var i = 0; i < result.length; i++) {
            if (result[i][1] == "実績ファイルDL" && result[i][3] == "成功") {
                return true;
            }
        }
    }

    // console.log(success)

    if (success) {
        await page.goto(OnePayConfig.CHECKBILLMANAGE_URL, {waitUntil: 'domcontentloaded'}); // wait until page load
        await page.select('select[name="status"]', "02")
        await page.waitForTimeout(2000);
        await page.waitForSelector('[type="submit"]');
        await page.click('[type="submit"]')
        await page.waitForTimeout(2000);

        await page.waitForSelector('#tbInfo_info')
        let element = await page.$('#tbInfo_info')
        let value = await page.evaluate(el => el.textContent, element)
        const before_ = value.substring(0, value.indexOf('ページ'));
        const pageNo = before_.split('、')[1];

        const FinalArray = []
        for (var i = 1; i <= pageNo; i++) {
            await page.waitForSelector('#tbInfo')
            await page.waitForTimeout(2000);

            const result = await page.$$eval('#tbInfo tbody tr', rows => {
                return Array.from(rows, row => {
                    const columns = row.querySelectorAll('td');
                    return Array.from(columns, column => column.innerText);
                });
            });

            FinalArray.push(result)

            await page.waitForSelector('#tbInfo_last')
            await page.click('#tbInfo_last');
        }
        const merge3 = FinalArray.flat(1); //The depth level specifying how deep a nested array structure should be flattened. Defaults to 1.
        var keyArray = merge3.map(function (item) {
            return item[1];
        });

        const ConditionOneCheck = keyArray.filter(element => OnePayConfig.CONDITIONONE.split(",").includes(element));
        const ConditionTwoCheck = keyArray.filter(element => OnePayConfig.CONDITIONTWO.split(",").includes(element));
        const ConditionThreeCheck = keyArray.filter(element => OnePayConfig.CONDITIONTHREE.split(",").includes(element));

        if (ConditionOneCheck.length > 0) {
            content = '(データ) 異常がありません  [ ' + ConditionOneCheck + ' ]';
            robot.text(content);

        }
        if (ConditionTwoCheck.length > 0) {

            content = '(データ) CPSの失敗件があります。至急対応してください。[ ' + ConditionTwoCheck + ' ]';
            robot.text(content);

        }
        if (ConditionThreeCheck.length > 0) {
            content = '(データ) Hundsunの失敗件があります。至急対応してください。[ ' + ConditionThreeCheck + ' ]';
            robot.text(content);

        }
        if (ConditionOneCheck.length == 0 && ConditionTwoCheck.length == 0 && ConditionThreeCheck.length == 0) {
            content = '(データ) 問題なしリスト以外の失敗件があります。確認してください。';
            robot.text(content);

        }
    }
    await browser.close();
}

log_in()
 

