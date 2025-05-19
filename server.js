const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");

const app = express();

app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.post("/outlook/webhook", (req, res) => {
  if (req?.query && req?.query?.validationToken) {
    res.set("Content-Type", "text/plain");
    res.send(req.query.validationToken);
    return;
  }

  const data = req?.body?.value || null;

  if (!data) return res.status(200).json("No Data!");

  console.log("Received:", data);

  res.status(200).json("Data received successfully!");
});

app.post("/outlook/webhook/lifecycle", (req, res) => {
  if (req?.query && req?.query?.validationToken) {
    res.set("Content-Type", "text/plain");
    res.send(req.query.validationToken);
    return;
  }

  const data = req?.body?.value;

  if (!data) return res.status(200).json("No Data!");

  console.log("Lifecycle Notification", req.body);

  res.status(200).json("Success!");
});

app.listen(3000, () => {
  console.log("App listening on port 3000!");
});
