const ChatBot = require('dingtalk-robot-sender');
const robot = new ChatBot({
   webhook: 'https://oapi.dingtalk.com/robot/send?access_token=290fc32bb41a5326ad0e44fa59ee2f7c60b58f40820a904268665656ab025f73'
});
robot.text("Report from node server\n" +
    "all true ");