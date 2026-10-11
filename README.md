# Alkitab Melayu
Static front end (`public/`) + Vercel serverless API (`api/`) + Neon Postgres.

1. Create a Neon project, copy the connection string.
2. Push this folder to GitHub, import it in Vercel (framework preset: Other).
3. Vercel env vars: `DATABASE_URL`, `ADMIN_PASSWORD`, `SESSION_SECRET`.
4. Deploy, open `/#admin`, log in, then open `/api/setup` once (creates tables + seeds the 73 books).
5. Optional: Admin > "Import data lama" to upload the JSON exported from the localStorage version.
Local: `npm i -g vercel`, copy `.env.example` to `.env`, run `vercel dev`.

Updating an existing deployment? After redeploying, log in and open `/api/setup` once more (adds English book names and the per-translation book names table).
