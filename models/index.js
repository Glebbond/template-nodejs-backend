const mongoose = require("mongoose");
const Schema = mongoose.Schema;
const someModelSchema = require("./someModel");

module.exports = {
  SomeModel: mongoose.model('SomeModel', someModelSchema(mongoose, Schema)),
}
