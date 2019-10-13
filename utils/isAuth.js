const jwt = require('express-jwt');
const config = require('../configs/config.json');
const getTokenFromHeader = (req) => {
  if (req.headers.authorization && req.headers.authorization.split(' ')[0] === 'Bearer') {
    return req.headers.authorization.split(' ')[1];
  }
}

module.exports = jwt({
  secret: config.secretKey,
  userProperty: 'token',
  getToken: getTokenFromHeader
})