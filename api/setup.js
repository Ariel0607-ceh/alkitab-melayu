import { sql, isAdmin } from "./_lib.js";
const OT = "GEN|Kejadian|50,EXO|Keluaran|40,LEV|Imamat|27,NUM|Bilangan|36,DEU|Ulangan|34,JOS|Yosua|24,JDG|Hakim-hakim|21,RUT|Rut|4,1SA|1 Samuel|31,2SA|2 Samuel|24,1KI|1 Raja-raja|22,2KI|2 Raja-raja|25,1CH|1 Tawarikh|29,2CH|2 Tawarikh|36,EZR|Ezra|10,NEH|Nehemia|13,EST|Ester|10,JOB|Ayub|42,PSA|Mazmur|150,PRO|Amsal|31,ECC|Pengkhotbah|12,SNG|Kidung Agung|8,ISA|Yesaya|66,JER|Yeremia|52,LAM|Ratapan|5,EZK|Yehezkiel|48,DAN|Daniel|12,HOS|Hosea|14,JOL|Yoel|3,AMO|Amos|9,OBA|Obaja|1,JON|Yunus|4,MIC|Mikha|7,NAM|Nahum|3,HAB|Habakuk|3,ZEP|Zefanya|3,HAG|Hagai|2,ZEC|Zakharia|14,MAL|Maleakhi|4";
const NT = "MAT|Matius|28,MRK|Markus|16,LUK|Lukas|24,JHN|Yohanes|21,ACT|Kisah Para Rasul|28,ROM|Roma|16,1CO|1 Korintus|16,2CO|2 Korintus|13,GAL|Galatia|6,EPH|Efesus|6,PHP|Filipi|4,COL|Kolose|4,1TH|1 Tesalonika|5,2TH|2 Tesalonika|3,1TI|1 Timotius|6,2TI|2 Timotius|4,TIT|Titus|3,PHM|Filemon|1,HEB|Ibrani|13,JAS|Yakobus|5,1PE|1 Petrus|5,2PE|2 Petrus|3,1JN|1 Yohanes|5,2JN|2 Yohanes|1,3JN|3 Yohanes|1,JUD|Yudas|1,REV|Wahyu|22";
const DC = "TOB|Tobit|14,JDT|Yudit|16,WIS|Kebijaksanaan Salomo|19,SIR|Putra Sirakh|51,BAR|Barukh|6,1MA|1 Makabe|16,2MA|2 Makabe|15";
const G = [["Perjanjian Lama", OT], ["Perjanjian Baru", NT], ["Deuterokanonika", DC]];
export default async function handler(req, res) {
  if (!isAdmin(req)) return res.status(401).json({ error: "Log masuk admin dahulu (/#admin), kemudian buka semula /api/setup" });
  try {
    await sql.transaction([
      sql`create table if not exists groups(name text primary key, pos int not null default 0)`,
      sql`create table if not exists books(code text primary key, name text not null, chapters int not null default 1, grp text not null references groups(name) on update cascade, pos int not null default 0)`,
      sql`create table if not exists translations(id text primary key, name text not null, translator text default '', year text default '', century text not null default '19')`,
      sql`create table if not exists verses(tid text references translations(id) on delete cascade, book text references books(code) on update cascade on delete cascade, ch int, v int, text text not null, primary key(tid,book,ch,v))`,
    ]);
    const rows = G.flatMap(([g, s]) => s.split(",").map((x) => { const [c, n, ch] = x.split("|"); return [c, n, +ch, g]; }));
    await sql`insert into groups(name,pos) select * from unnest(${G.map((g) => g[0])}::text[], ${G.map((_, i) => i)}::int[]) on conflict do nothing`;
    await sql`insert into books(code,name,chapters,grp,pos) select * from unnest(${rows.map((r) => r[0])}::text[], ${rows.map((r) => r[1])}::text[], ${rows.map((r) => r[2])}::int[], ${rows.map((r) => r[3])}::text[], ${rows.map((_, i) => i)}::int[]) on conflict do nothing`;
    res.json({ ok: true, msg: "Pangkalan data sedia. Kembali ke laman utama." });
  } catch (e) { res.status(500).json({ error: e.message }); }
}
