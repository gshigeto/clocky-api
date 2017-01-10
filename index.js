const express = require('express');
const bodyParser = require('body-parser');
var cors = require('cors');
const app = express();
var google = require('./google.js');

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());

app.get('/', (req, res) => {
  res.send('Hello World!');
});

app.post('/in/:docId', (req, res) => {
  google.createSpreadsheet(createToken(req.body), req.params.docId).then(function (response) {
    google.clockIn(createToken(req.body), req.body.timestamp, response.body.docId).then(function (response) {
      res.status(response.code).send(response.body);
    }).catch(function (err) {
      console.log(err);
      res.status(err.code).send(err.message);
    });
  }).catch(function (err) {
    console.log(err);
    res.status(err.code).send(err.message);
  });
});

app.post('/out/:docId', (req, res) => {
  google.createSpreadsheet(createToken(req.body), req.params.docId).then(function (response) {
    google.clockOut(createToken(req.body), req.body.timestamp, response.body.docId).then(function (response) {
      res.status(response.code).send(response.body);
    }).catch(function (err) {
      console.log(err);
      res.status(err.code).send(err.message);
    });
  }).catch(function (err) {
    console.log(err);
    res.status(err.code).send(err.message);
  });
});

app.post('/export/:docId', (req, res) => {
  google.createSpreadsheet(createToken(req.body), req.params.docId).then(function (response) {
    google.exportTimesheet(createToken(req.body), req.body.shifts, response.body.docId).then(function (response) {
      res.status(response.code).send(response.body);
    }).catch(function (err) {
      console.log(err);
      res.status(err.code).send(err.message);
    });
  }).catch(function (err) {
    console.log(err);
    res.status(err.code).send(err.message);
  });
});

app.post('/spreadsheet/create', (req, res) => {
  google.createSpreadsheet(createToken(req.body), '-1').then(function (response) {
    res.status(response.code).send(response.body);
  }).catch(function (err) {
    console.log(err);
    res.status(err.code).send(err.message);
  });
});

app.listen(3000, () => {
  console.log('App listening on port 3000!');
});

function createToken(body) {
  return {
    access_token: body.access_token,
    refresh_token: body.refresh_token,
    token_type: body.token_type
  };
}