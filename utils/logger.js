const log4js = require('log4js');
const moment = require('moment');
log4js.configure({
  appenders: { ecms: { type: 'file', filename: `log/ecms.${moment().format('YYYY-MM-DD')}.log` } },
  categories: { default: { appenders: ['ecms'], level: 'INFO' } }
});
module.exports = log4js.getLogger('ecms');