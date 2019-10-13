const db = require('../models/index');

module.exports = async (req, res, next) => {
    const decodedTokenData =  req.token;
    const userRecord = await db.User.findOne({ login: decodedTokenData.login })
    req.currentUser = userRecord;
   
    if(!userRecord) {
      return res.status(401).send({
        success: false,
        message: 'User not found'
      })
    } else {
      return next();
    }
}