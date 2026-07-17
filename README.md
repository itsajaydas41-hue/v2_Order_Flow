# OrderFlow — O2C + P2P App (Final Package)

Ye folder aapki poori app hai. Isko GitHub par upload karke Vercel se live karna hai.

## Folder me kya hai

| File | Kya hai | Chhedna hai? |
|---|---|---|
| `index.html` | App ka poora UI (dashboard, orders, gate, collection…) | ❌ Nahi |
| `api-supabase.js` | App ka engine (saara logic + Supabase connection) | ❌ Nahi |
| `config.js` | Supabase ka URL + key (**pehle se bhara hua**) | ✅ Sirf project badle to |
| `_setup/` | Ek baar ke setup scripts (Vercel isse ignore karega) | — |

---

# PART 1 — App ko LIVE karna (ek hi baar, ~10 min)

## Step 1 — Supabase me naye columns add karo (pehle ye karo!)
1. **supabase.com** kholo → apna project → left menu → **SQL Editor** → **New query**
2. `_setup/add-columns.sql` file Notepad me kholo → sab copy karo → SQL Editor me paste karo
3. **Run** dabao → **"Success"** aana chahiye

> Ye zaroori hai. Warna Gate Entry ka Mobile No aur Collection ke naye fields save nahi honge.

## Step 2 — GitHub par naya repository banao
1. **github.com** kholo → login karo
2. Upar right me **+** → **New repository**
3. **Repository name:** `orderflow-app`
4. **Public** select karo → **Create repository**

## Step 3 — Files upload karo
1. Naye repo ke page par → **"uploading an existing file"** link dabao
   (ya **Add file** → **Upload files**)
2. Apne computer me `FINAL-APP` folder kholo
3. In **3 files** ko select karke drag karo:
   - `index.html`
   - `api-supabase.js`
   - `config.js`

   > ⚠️ **Folder mat drag karo — sirf files.** Files repo ke andar sabse upar (root) me honi chahiye.
4. Neeche **Commit changes** dabao

## Step 4 — Vercel me naya project banao
1. **vercel.com** kholo → **Add New** → **Project**
2. **Import Git Repository** me apna `orderflow-app` dikhega → **Import** dabao
   - Agar na dikhe → **Adjust GitHub App Permissions** → repo ko allow karo
3. Settings me kuch **mat** badlo:
   - Framework Preset: **Other**
   - Build Command: **khaali**
   - Root Directory: **./**
4. **Deploy** dabao → 30 second wait karo

## Step 5 — App kholo aur check karo
1. Vercel green screen dikhayega → **Continue to Dashboard** → upar **Domains** me URL milega
   (jaise `orderflow-app.vercel.app`) — **yahi aapka permanent address hai**
2. URL kholo → **Ctrl + Shift + R** dabao
3. Login karo → Dashboard me ye dikhna chahiye:
   - Subtitle me **"build v2"**
   - Rangeen **pipeline flow** with chalti hui animation
   - Neeche **Monthly Orders** chart (month par click karo → weekly chart badlega)

✅ Ho gaya. Ye URL team ko share karo.

---

# PART 2 — Future me code change karna (bahut easy)

Jab bhi main aapko nayi file du (jaise naya `index.html`):

1. **github.com** → apna `orderflow-app` repo kholo
2. **Add file** → **Upload files**
3. Nayi file drag karo (same naam ki)
4. **Commit changes** dabao
5. Bas! Vercel apne aap 30 second me update kar dega — **URL wahi rahega**

> Ctrl + Shift + R dabana mat bhoolna, warna purana version cache me dikhega.

**Kabhi naya Vercel project mat banao** — warna URL badal jayega. Hamesha isi repo me upload karo.

---

# PART 3 — Google Sheet mirror (optional)

Agar Sheet me data dekhna hai:
1. Google Sheet → **Extensions** → **Apps Script**
2. Purana code delete karo → `_setup/sync-to-sheets.gs` paste karo
3. Upar `MIR_URL` aur `MIR_KEY` me wahi values daalo jo `config.js` me hain
4. Function dropdown → **mirSyncNow** → **Run** (permission allow karo)
5. Function dropdown → **mirInstallTrigger** → **Run** (har 5 min auto-update)

---

# Kuch important baatein

- **Data Supabase me hai**, Sheet sirf dekhne ke liye mirror hai (5 min purana ho sakta hai)
- **Purane apps delete kar do** — Vercel me `o2c-app-v3`, `srfm-erp`, `srfmerp` (Settings → Delete Project), taki koi galti se purana version na khole
- **Security:** abhi anon key se koi bhi data padh/likh sakta hai. Internal team ke liye theek hai, par baad me proper login + security lagwa lena (bol dena, laga dunga)
- **Backup:** Supabase → Table Editor → har table CSV export kar sakte ho

# Kuch problem aaye to

| Problem | Solution |
|---|---|
| 404 NOT_FOUND | Vercel → Deployments → newest → ⋯ → **Promote to Production** |
| "build v2" nahi dikh raha | Purani file hai — dubara upload karo + Ctrl+Shift+R |
| "Supabase 404 on table…" | SQL Editor me chalao: `NOTIFY pgrst, 'reload schema';` |
| Naye fields save nahi ho rahe | Step 1 ka `add-columns.sql` nahi chala — chala do |
