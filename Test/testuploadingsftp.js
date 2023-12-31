// import ssh2-sftp-client
const SFTPClient = require("ssh2-sftp-client")
// ESM: import SFTPClient from "ssh2-sftp-client"

const SFTPhost= "192.168.0.77";
const SFTPport ='2836';
const SFTPUserName= "onepay";
const SFTPpassword= "onepay001";

async function operationWithSFTP() {
    const sftp = new SFTPClient()
    sftp.connect({
        host: SFTPhost,
        port: SFTPport,
        username: SFTPUserName,
        password: SFTPpassword
    }).then(() => {
        sftp.mkdir("/Public/SquareDownload/Date",false);
        sftp.put('D:\\2023-07-13\\NewDownloadLocation\\【支払明細票 兼 売上票】日本恒生ソフトウェア株式会社-Square株式会社様(20230601_20230615)_20230618025830.zip', '/Public/SquareDownload/date0.zip', false);
        console.log("upload complete")
    }).then((data) => {
        console.log(data, 'the data info');
    }).catch((err) => {
        console.log(err, 'catch error');
    });
}

operationWithSFTP();