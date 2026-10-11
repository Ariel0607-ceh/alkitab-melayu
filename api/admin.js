import { sql, isAdmin } from "./_lib.js";
const cen = (c) => String(c).replace("17-18", "17");
const ops = {
  verse: (b) => sql`insert into verses(tid,book,ch,v,text) values(${b.tid},${b.book},${+b.ch},${+b.v},${String(b.text).slice(0, 5000)}) on conflict(tid,book,ch,v) do update set text = excluded.text`,
  delVerse: (b) => sql`delete from verses where tid=${b.tid} and book=${b.book} and ch=${+b.ch} and v=${+b.v}`,
  tr: (b) => sql`insert into translations(id,name,translator,year,century) values(${b.id},${b.name},${b.translator || ""},${b.year || ""},${cen(b.century)}) on conflict(id) do update set name=excluded.name, translator=excluded.translator, year=excluded.year, century=excluded.century`,
  delTr: (b) => sql`delete from translations where id=${b.id}`,
  book: (b) => b.old
    ? sql`update books set code=${b.c}, name=${b.n}, name_en=${b.en || null}, chapters=${+b.ch}, grp=${b.g} where code=${b.old}`
    : sql`insert into books(code,name,name_en,chapters,grp,pos) values(${b.c},${b.n},${b.en || null},${+b.ch},${b.g},(select coalesce(max(pos),0)+1 from books))`,
  delBook: (b) => sql`delete from books where code=${b.c}`,
  bookName: (b) => b.name
    ? sql`insert into book_names(tid,book,name) values(${b.tid},${b.book},${String(b.name).slice(0, 100)}) on conflict(tid,book) do update set name=excluded.name`
    : sql`delete from book_names where tid=${b.tid} and book=${b.book}`,
  async move(b) {
    const [me] = await sql`select pos, grp from books where code=${b.c}`;
    const [o] = b.dir < 0
      ? await sql`select code, pos from books where grp=${me.grp} and pos<${me.pos} order by pos desc limit 1`
      : await sql`select code, pos from books where grp=${me.grp} and pos>${me.pos} order by pos limit 1`;
    if (o) await sql.transaction([
      sql`update books set pos=${o.pos} where code=${b.c}`,
      sql`update books set pos=${me.pos} where code=${o.code}`,
    ]);
  },
  group: (b) => b.old
    ? sql`update groups set name=${b.name} where name=${b.old}`
    : sql`insert into groups(name,pos) values(${b.name},(select coalesce(max(pos),0)+1 from groups))`,
  delGroup: (b) => sql`delete from groups where name=${b.name}`,
  async importMeta(b) {
    const gs = b.groups || [], bs = b.books || [], ts = b.translations || [];
    if (gs.length) await sql`insert into groups(name,pos) select * from unnest(${gs}::text[], ${gs.map((_, i) => i + 100)}::int[]) on conflict do nothing`;
    if (bs.length) await sql`insert into books(code,name,chapters,grp,pos) select * from unnest(${bs.map((x) => x.c)}::text[], ${bs.map((x) => x.n)}::text[], ${bs.map((x) => +x.ch)}::int[], ${bs.map((x) => x.g)}::text[], ${bs.map((_, i) => i + 100)}::int[]) on conflict(code) do update set name=excluded.name, chapters=excluded.chapters`;
    if (ts.length) await sql`insert into translations(id,name,translator,year,century) select * from unnest(${ts.map((t) => t.id)}::text[], ${ts.map((t) => t.name)}::text[], ${ts.map((t) => t.translator || "")}::text[], ${ts.map((t) => String(t.year || ""))}::text[], ${ts.map((t) => cen(t.century))}::text[]) on conflict(id) do update set name=excluded.name, translator=excluded.translator, year=excluded.year, century=excluded.century`;
  },
  async importVerses(b) {
    const r = b.rows || [];
    if (r.length) await sql`insert into verses(tid,book,ch,v,text) select * from unnest(${r.map((x) => x[0])}::text[], ${r.map((x) => x[1])}::text[], ${r.map((x) => +x[2])}::int[], ${r.map((x) => +x[3])}::int[], ${r.map((x) => x[4])}::text[]) on conflict(tid,book,ch,v) do update set text=excluded.text`;
  },
};
export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).end();
  if (!isAdmin(req)) return res.status(401).json({ error: "Sesi tamat. Log masuk semula." });
  const b = req.body || {};
  try {
    if (!Object.hasOwn(ops, b.op)) throw Error("Operasi tidak sah");
    await ops[b.op](b);
    res.json({ ok: true });
  } catch (e) {
    res.status(400).json({ error: /violates|duplicate/.test(e.message) ? "Tidak boleh: kod/nama sudah wujud atau masih digunakan" : e.message });
  }
}
