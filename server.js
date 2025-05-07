const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");

const app = express();

app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.post("/api/submit", (req, res) => {
  const data = req.body;
  console.log("Received data Submit:", JSON.stringify(data, null, 2));
  res.status(200).json({ message: "Data received successfully!" });
});

app.get("/outlook/webhook", (req, res) => {
  const token = req.query.validationToken;
  console.log("Received Token", token);
  res.status(200).send(token);
});

app.post("/outlook/webhook", (req, res) => {
  const data = req.body;
  console.log("Received data:", JSON.stringify(data, null, 2));
  res.status(200).json({ message: "Data received successfully!" });
});

app.listen(3000, () => {
  console.log("App listening on port 3000!");
});
