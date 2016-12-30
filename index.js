const express = require('express');
const bodyParser = require('body-parser');
var cors = require('cors');
const app = express();

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());

app.get('/', (req, res) => {
  res.send('Hello World!');
});

app.post('/in', (req, res) => {
  res.send({message: 'CLOCKED IN'});
});

app.post('/out', (req, res) => {
  res.send({message: 'CLOCKED OUT'});
});

app.listen(3000, () => {
  console.log('App listening on port 3000!');
});