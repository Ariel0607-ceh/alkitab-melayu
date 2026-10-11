import { sql } from "./_lib.js";
export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  const { book, ch, t = "" } = req.query;
  const ids = String(t).split(",").filter(Boolean).slice(0, 6);
  try {
    const rows = ids.length
      ? await sql`select tid, v, text from verses where book = ${String(book)} and ch = ${+ch} and tid = any(${ids}::text[]) order by v`
      : [];
    res.json(rows);
  } catch (e) { res.status(500).json({ error: e.message }); }
}
