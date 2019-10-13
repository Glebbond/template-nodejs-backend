const express = require('express'),
  app = express(),
  bodyParser = require('body-parser'),
  routes = require('./routes/index'),
  mongoose = require("mongoose"),
  cors = require('cors'),
  passport = require('passport'),
  LocalStrategy = require('passport-local').Strategy;
const host = 'localhost';
const port = 3001;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());

passport.use(new LocalStrategy({
    usernameField: 'login',
    passwordField: 'password',
    passReqToCallback: true
  },
  validateUser
));
passport.serializeUser(function(user, cb) {
  cb(null, user._id);
});

passport.deserializeUser(function(id, cb) {
  cb(null, user);
});

app.use('/api', routes);
app.use(passport.initialize());
app.use(passport.session());

mongoose.connect(`mongodb://${config.dbHostName}:27017/administration`, { useNewUrlParser: true }, function(err){
  if (err) throw err;
  app.listen(port, host, () => console.log(`Server listens http://${host}:${port}`));
  console.log('Successfully connected');
});

