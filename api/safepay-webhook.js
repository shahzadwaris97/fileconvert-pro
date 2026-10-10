
import crypto from "node:crypto";
import { createClient } from "@supabase/supabase-js";

export const config = {
  api: {
    bodyParser: false,
  },
};

function verifySignature(secret, rawBody, signature) {
  if (
    !secret ||
    typeof signature !== "string" ||
    !/^[a-fA-F0-9]{64}$/.test(signature)
  ) {
    return false;
  }

  const expected = crypto
    .createHmac("sha256", secret)
    .update(rawBody)
    .digest();

  const received = Buffer.from(signature, "hex");

  return crypto.timingSafeEqual(expected, received);
}

async function readRawBody(req) {
  const chunks = [];

  for await (const chunk of req) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }

  return Buffer.concat(chunks);
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).send("Method not allowed");
  }

  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SECRET_KEY;
  const webhookSecret = process.env.SAFEPAY_WEBHOOK_SECRET;

  if (!url || !key || !webhookSecret) {
    return res.status(500).send("Server configuration missing");
  }

  try {
    const rawBody = await readRawBody(req);
    const signature = req.headers["x-sfpy-signature"];

    if (!verifySignature(webhookSecret, rawBody, signature)) {
      return res.status(401).send("Invalid signature");
    }

    let event;

    try {
      event = JSON.parse(rawBody.toString("utf8"));
    } catch {
      return res.status(400).send("Invalid JSON");
    }

    if (
      typeof event.token !== "string" ||
      typeof event.type !== "string"
    ) {
      return res.status(400).send("Missing event ID or type");
    }

    const supabase = createClient(url, key, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    });

    const { error } = await supabase
      .from("webhook_events")
      .insert({
        event_id: event.token,
        event_type: event.type,
        payload: event,
        processed: false,
      });

    if (error?.code === "23505") {
      return res.status(200).send("Duplicate event received");
    }

    if (error) {
      console.error("Could not store webhook event");
      return res.status(500).send("Event storage failed");
    }

    return res.status(200).send("Webhook received");
  } catch {
    console.error("Webhook request failed");
    return res.status(500).send("Webhook processing failed");
  }
}
