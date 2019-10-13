const nodemailer = require('nodemailer');
const config = require('../configs/config.json');

const transporter = nodemailer.createTransport({
    host: config.emailData.host,
    port: config.emailData.port,
    secure: true,
    secureConnection: true,
    auth: {
        user: config.emailData.user,
        pass: config.emailData.pass
    }
});

module.exports = async (mailData) => {
    let info =  await transporter.sendMail({
        from: `<${config.emailData.user}>`, // sender address
        to: mailData.email,
        subject: mailData.subject,
        text: mailData.text,
        html: mailData.html
    });
    console.log('Message sent: %s', info.messageId);
}

