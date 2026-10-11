import { mkToken, same } from "./_lib.js";
const C = "Path=/; HttpOnly; Secure; SameSite=Strict";
export default async function handler(req, res) {
  if (req.method === "DELETE") { res.setHeader("Set-Cookie", `am=; ${C}; Max-Age=0`); return res.json({ ok: true }); }
  if (req.method !== "POST") return res.status(405).end();
  const p = String((req.body || {}).password || "");
  if (!process.env.ADMIN_PASSWORD || !same(p, process.env.ADMIN_PASSWORD)) {
    await new Promise((r) => setTimeout(r, 700));
    return res.status(401).json({ error: "Kata laluan salah" });
  }
  res.setHeader("Set-Cookie", `am=${mkToken()}; ${C}; Max-Age=604800`);
  res.json({ ok: true });
}
