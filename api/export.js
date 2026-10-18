import { sql } from "./_lib.js";
// One book of one translation per request (keeps each response small enough for Vercel).
export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  const { tid, book } = req.query;
  try {
    res.json(await sql`select ch, v, text from verses where tid = ${String(tid)} and book = ${String(book)} order by ch, v`);
  } catch (e) { res.status(500).json({ error: e.message }); }
}
