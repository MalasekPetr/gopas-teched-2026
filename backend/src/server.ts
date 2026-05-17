import "dotenv/config";
import express from "express";
import { webhookRouter } from "./routes/webhook.js";
import { manageSubscriptionsRouter } from "./routes/manageSubscriptions.js";

const app = express();

// SharePoint sends application/json notifications, but the handshake is plain text.
app.use(express.json());
app.use(express.text({ type: "text/plain" }));

app.use("/webhook", webhookRouter);
app.use("/manage-subscriptions", manageSubscriptionsRouter);

app.get("/", (_req, res) => res.send("SP Webhook Demo backend up"));

const port = Number(process.env.PORT) || 8080;
app.listen(port, () => console.log(`Listening on ${port}`));
