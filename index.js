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
  google.createSpreadsheet(createToken(req.body), req.params.docId).then(function (docId) {
    google.clockIn(createToken(req.body), req.body.timestamp, docId).then(function (response) {
      res.send(response);
    }).catch(function (err) {
      res.status(400).send(err);
    });
  }).catch(function (err) {
    res.status(err.code).send(err.message);
  });
});

app.post('/out/:docId', (req, res) => {
  google.createSpreadsheet(createToken(req.body), req.params.docId).then(function (docId) {
    google.clockOut(createToken(req.body), req.body.timestamp, docId).then(function (response) {
      res.send(response);
    }).catch(function (err) {
      res.status(400).send(err);
    });
  }).catch(function (err) {
    res.status(err.code).send(err.message);
  });
});

app.post('/sheet/create', (req, res) => {
  google.createSpreadsheet(createToken(req.body), '-1').then(function (docId) {
    res.send({docId: docId});
  }).catch(function (err) {
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