import { neon } from "@neondatabase/serverless";
import crypto from "node:crypto";
export const sql = neon(process.env.DATABASE_URL);
const key = () => process.env.SESSION_SECRET || process.env.ADMIN_PASSWORD + "-s";
const sig = (v) => crypto.createHmac("sha256", key()).update(v).digest("hex");
export const mkToken = () => { const e = String(Date.now() + 7 * 864e5); return e + "." + sig(e); };
export const isAdmin = (req) => {
  const m = /(?:^|; )am=([^;]+)/.exec(req.headers.cookie || "");
  if (!m || !process.env.ADMIN_PASSWORD) return false;
  const [e, s = ""] = m[1].split(".");
  const good = sig(e);
  return s.length === good.length && +e > Date.now() && crypto.timingSafeEqual(Buffer.from(s), Buffer.from(good));
};
export const same = (a, b) => {
  const h = (x) => crypto.createHash("sha256").update(x).digest();
  return crypto.timingSafeEqual(h(a), h(b));
};
