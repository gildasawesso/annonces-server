let express = require('express');
let path = require('path');
let cookieParser = require('cookie-parser');
let logger = require('morgan');
let cors = require('cors');

let uploadRouter = require('./routes/users');
let indexRouter = require('./routes/index');

let app = express();

app.use(cors());
app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());

app.use('/upload', uploadRouter);
app.use('/', indexRouter);

module.exports = app;
