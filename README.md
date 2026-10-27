# Data Practice Lab

A free learning site for Excel, SQL and Power BI: 188 theory chapters, 197 hands-on lessons and 349 ready-to-copy formulas.

- Sign in with email and password, or Google (Supabase Auth)
- Progress saved to each learner's account (Supabase Postgres, protected by row-level security)
- Completion certificates with a public verification page
- A leaderboard (learners can hide themselves)
- An admin dashboard for the owner
- Self-service "Download my data" and "Delete my account"

It's a static site with no build step, hosted on Vercel.

Licence: GPL-3.0 (see `LICENSE`). It includes HyperFormula (GPL-3.0) and sql.js (MIT).

## Try it on your computer

Run `python -m http.server 8000` in this folder and open <http://localhost:8000>. While `js/config.js` still has the placeholder values, the site runs in **preview mode**: any email and password works, data stays in your browser, and you're an admin.

## Setup

### 1. Supabase

1. Create a project at <https://supabase.com/dashboard>. Choose a region near your users (Mumbai for India).
2. Open **SQL Editor**, paste all of `supabase/schema.sql` and click **Run**.
3. Open **Project Settings → API**. Copy the **Project URL** and the **anon / publishable** key into `js/config.js`. Never use the `service_role` or secret key in the website.
4. Open **Authentication → URL Configuration**:
   - **Site URL**: your live address, for example `https://data-practice-lab.vercel.app`
   - **Redirect URLs**: add `https://data-practice-lab.vercel.app/**` and `http://localhost:8000/**`
5. Optional, **Authentication → Providers → Email → Confirm email**:
   - **On**: people must click a link before they can sign in. Supabase's built-in email sender only sends a few emails per hour, so add your own SMTP (Authentication → Emails → SMTP settings) before you have many sign-ups.
   - **Off**: people sign in immediately.

### 2. Google sign-in (optional)

1. In Google Cloud Console, create an **OAuth client ID** (type: Web application).
2. Add the **Authorized redirect URI** that Supabase shows under **Authentication → Providers → Google**. It looks like `https://<project>.supabase.co/auth/v1/callback`.
3. Paste the Google **Client ID** and **Client secret** into that Supabase page and enable Google.

Until you do this, the Google button shows a friendly message and email sign-in still works.

### 3. Vercel

1. Import this GitHub repository in Vercel. Framework preset: **Other**. No build command and no output directory are needed.
2. Every push to the main branch redeploys automatically.

### 4. Make yourself admin

1. Sign in on the live site once.
2. In Supabase, open **Authentication → Users** and copy your user UID.
3. In **Table Editor → admins**, insert a row with `id` = your UID.
4. Reload the site. **Admin dashboard** appears in your account menu.

## Good to know

- **Free tiers.** Supabase Free (500 MB database, 50,000 monthly active users) and Vercel Hobby are enough to start. A Supabase free project pauses after a week with no activity; it wakes when you open it in the dashboard.
- **The login gate is soft.** Lessons are in `js/app.bundle.js`, which anyone can download, and the code is open source. Login protects each learner's data, not the lessons.
- **Scores can be faked** by a determined user with their own script. The database caps them at a sensible number. Treat the leaderboard as fun, and certificates as a record of completion, not an exam.
- **Privacy and terms.** `privacy.html` and `terms.html` are drafts. Fill in the placeholders and have them reviewed.

## Files

| File | What it is |
| --- | --- |
| `index.html` | Landing page, sign-in and the app |
| `js/config.js` | **Your settings**: Supabase URL and public key, site details |
| `js/cloud.js` | Login, progress sync, certificates, leaderboard, admin |
| `js/app.bundle.js` | The learning app (lessons, practice engines) |
| `supabase/schema.sql` | Tables, security rules and functions |
| `vendor/` | HyperFormula and sql.js |
| `verify.html` | Public certificate check |
| `privacy.html`, `terms.html` | Draft policies |
| `vercel.json` | Hosting headers |
