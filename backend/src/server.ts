import "dotenv/config";
import express from "express";
import { webhookRouter } from "./routes/webhook.js";
import { manageSubscriptionsRouter } from "./routes/manageSubscriptions.js";

const app = express();

// Graph posts JSON notifications, plus an empty-body POST for the validation handshake.
app.use(express.json());

app.use("/webhook", webhookRouter);
app.use("/manage-subscriptions", manageSubscriptionsRouter);

app.get("/", (_req, res) => res.send("SP Webhook Demo backend up"));

const port = Number(process.env.PORT) || 8080;
app.listen(port, () => console.log(`Listening on ${port}`));
