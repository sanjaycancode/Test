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

app.post("/outlook/webhook/lifecycle", async (req, res) => {
  try {
    if (req?.query && req?.query?.validationToken) {
      res.set("Content-Type", "text/plain");
      res.send(req.query.validationToken);
      return;
    }

    const data = req?.body?.value;

    if (!data) return res.status(200).json("No Data!");

    if (!data?.length) return res.status(500).json("Invalid Lifecycle Data!");

    const reauthorizationCycle = data?.find(
      (c) => c?.lifecycleEvent === "reauthorizationRequired"
    );

    console.log("Reauthorization Notification", reauthorizationCycle);
  } catch (error) {
    console.log({ error });
  }
});

app.listen(3000, () => {
  console.log("App listening on port 3000!");
});
