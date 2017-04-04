var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');

var multer = require('multer')
var PythonShell = require('python-shell');
var fsp = require('fs-promise');
var app = express();

app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

app.use(logger('dev'));

app.use(express.static(path.join(__dirname, 'public')));


const upload = multer({ dest: './uploads/' });
app.use(upload.fields(
  [
    { name: 'details' },
    { name: 'lcs' },
    { name: 'capt' }
  ]
))


const runPythonScript = () => {
  return new Promise((resolve, reject) => {
    PythonShell.run('./uploads/process.py', function(err) {
      if (err) reject(err);
      console.log('finished');
      resolve()
    });
  });
}

app.get('/', function(req, res, next) {
  res.render('index', { title: 'Express' });
});

app.post('/upload', function(req, res, next) {
  const files = req.files;
  fsp.readFile('./' + files.details[0].path)
    .then((data) => {
      path = __dirname + '/uploads/details.xlsx'
      return fsp.writeFile(path, data)
    })
    .then(() => {
      return fsp.readFile('./' + files.lcs[0].path)
    })
    .then((data) => {
      path = __dirname + '/uploads/LCS tracker.xlsx'
      return fsp.writeFile(path, data)
    })
    .then(() => {
      return fsp.readFile('./' + files.capt[0].path)
    })
    .then((data) => {
      path = __dirname + '/uploads/CAPT.xlsx';
      return fsp.writeFile(path, data)
    })
    .then(() => {
      return runPythonScript()
    })
    .then(() => {
      res.status(200).sendFile(__dirname + '/RTF_02.xlsx');
    })
    .catch((err) => {
      console.log(err)
    })
});






module.exports = app;
