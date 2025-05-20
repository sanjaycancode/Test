const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");

const app = express();

app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.post("/outlook/webhook", (req, res) => {
  // if (req?.query && req?.query?.validationToken) {
  //   res.set("Content-Type", "text/plain");
  //   res.send(req.query.validationToken);
  //   return;
  // }

  // const data = req?.body?.value || null;

  // if (!data) return res.status(200).json("No Data!");

  if (req.query && req.query.validationToken) {
    // Important: Set content type to text/plain
    res.set("Content-Type", "text/plain");

    // Send back the exact same validationToken with 200 OK status
    res.status(200).send(req.query.validationToken);

    console.log("Validation response sent successfully");
    return;
  }

  const data = req?.body?.value || null;

  if (!data) return res.status(400).send("Invalid Notification Data!");

  console.log(data);

  if (
    data?.changeType === "created" &&
    data?.resourceData?.["@odata.type"] === "#Microsoft.Graph.Message"
  ) {
    res.status(202).send("Notification Received! Will take time in our end.");

    console.log("Received:", data);
  }

  res.status(200).send("Notification Received!");
});

app.post("/outlook/webhook/lifecycle", async (req, res) => {
  try {
    if (req.query && req.query.validationToken) {
      // Important: Set content type to text/plain
      res.set("Content-Type", "text/plain");

      // Send back the exact same validationToken with 200 OK status
      res.status(200).send(req.query.validationToken);

      console.log("Validation response sent successfully");
      return;
    }

    const data = req?.body?.value;

    if (!data) return res.status(400).json("No Data!");

    if (!data?.length) return res.status(500).json("Invalid Lifecycle Data!");

    console.log(reauthorizationCycle);

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
