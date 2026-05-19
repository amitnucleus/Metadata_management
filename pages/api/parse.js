import { IncomingForm } from "formidable";
import fs from "fs";
import FormData from "form-data";
import fetch from "node-fetch";

export const config = { api: { bodyParser: false } };

const PARSER_API =
  process.env.PARSER_API_INTERNAL_URL ||
  process.env.NEXT_PUBLIC_PARSER_API_URL ||
  "http://parser-api:8000";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const form = new IncomingForm({ keepExtensions: true });

  form.parse(req, async (err, _fields, files) => {
    if (err) return res.status(400).json({ error: "Failed to parse form" });

    const uploaded = files.file?.[0] || files.file;
    if (!uploaded) return res.status(400).json({ error: "No file uploaded" });

    const fd = new FormData();
    fd.append("file", fs.createReadStream(uploaded.filepath), uploaded.originalFilename);

    try {
      const upstream = await fetch(`${PARSER_API}/parse`, { method: "POST", body: fd });
      const data = await upstream.json();
      if (!upstream.ok) return res.status(upstream.status).json(data);
      return res.status(200).json(data);
    } catch (e) {
      return res.status(502).json({ error: "Parser API unavailable: " + e.message });
    }
  });
}
