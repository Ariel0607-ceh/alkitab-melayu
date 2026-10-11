import { sql, isAdmin } from "./_lib.js";
export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  try {
    const [groups, books, translations, names] = await Promise.all([
      sql`select name from groups order by pos, name`,
      sql`select b.code c, b.name n, b.name_en en, b.chapters ch, b.grp g, (select count(*)::int from verses v where v.book = b.code) cnt from books b order by b.pos`,
      sql`select t.id, t.name, t.translator, t.year, t.century, (select count(*)::int from verses v where v.tid = t.id) cnt from translations t order by t.century, t.year, t.name`,
      sql`select tid, book, name from book_names`,
    ]);
    res.json({ groups: groups.map((g) => g.name), books, translations, names, admin: isAdmin(req) });
  } catch (e) { res.status(500).json({ error: "setup" }); }
}
