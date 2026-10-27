/* Data Practice Lab – website layer: login, cloud progress, certificates, leaderboard, admin.
   Copyright (C) the site owner. Licensed under GPL-3.0 (see LICENSE). */
const SB_VERSION = "2.117.2";
const cfg = window.SUPABASE_CONFIG || {};
const SITE = Object.assign({name:"Data Practice Lab", tagline:"Excel · SQL · Power BI", owner:"", contactEmail:"", sourceUrl:""}, window.SITE || {});
const configured = /^https:\/\//.test(cfg.url || "") && !!cfg.anonKey && !/PASTE/i.test(cfg.url + cfg.anonKey);
const isLocal = location.protocol === "file:" || ["localhost", "127.0.0.1", ""].includes(location.hostname);
const DEMO = !configured && isLocal;
const SYNC_KEYS = ["dpl-progress", "dpl-days", "dpl-libv", "dpl-libo", "dpl-theme", "dpl-ui"];
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const nowIso = () => new Date().toISOString();
let onRecovery = null;

/* ---------------- Backends ---------------- */
/* Paths look like "users/<uid>"; each collection is a table. Field names are camelCase here, snake_case in the database. */
const TABLE = {users:"profiles", progress:"progress", leaderboard:"leaderboard", certificates:"certificates", admins:"admins"};
const toSnake = o => { const r = {}; for (const k in o) r[k.replace(/[A-Z]/g, c => "_" + c.toLowerCase())] = o[k]; return r; };
const toCamel = o => { if (!o) return o; const r = {}; for (const k in o) r[k.replace(/_([a-z])/g, (_, c) => c.toUpperCase())] = o[k]; return r; };
const SB_ERRORS = [[/invalid login credentials/i, "auth/invalid-credential"], [/email not confirmed/i, "auth/email-not-confirmed"], [/already registered|already been registered|user_already_exists/i, "auth/email-already-in-use"],
  [/password should be at least|weak.?password/i, "auth/weak-password"], [/rate limit|too many/i, "auth/too-many-requests"], [/unable to validate email|invalid.*email|email address .* is invalid/i, "auth/invalid-email"],
  [/failed to fetch|network/i, "auth/network-request-failed"], [/provider is not enabled|unsupported provider/i, "auth/provider-off"], [/signups not allowed/i, "auth/signup-off"]];
function sbErr(e) { if (!e) return e; const m = e.message || String(e); const hit = SB_ERRORS.find(([re]) => re.test(m) || re.test(e.code || "")); const err = new Error(m); err.code = hit ? hit[1] : e.code; return err; }
async function supabaseBackend() {
  const {createClient} = await import(`https://cdn.jsdelivr.net/npm/@supabase/supabase-js@${SB_VERSION}/+esm`);
  const sb = createClient(cfg.url, cfg.anonKey, {auth:{persistSession:true, autoRefreshToken:true, detectSessionInUrl:true, flowType:"pkce"}});
  const back = location.origin + location.pathname;
  const mapUser = u => u && {uid:u.id, name:(u.user_metadata || {}).full_name || (u.user_metadata || {}).name || "", email:u.email || "", photo:(u.user_metadata || {}).avatar_url || "",
    verified:!!u.email_confirmed_at, provider:(u.app_metadata || {}).provider || "email"};
  const chk = ({data, error}) => { if (error) throw sbErr(error); return data; };
  const split = p => { const [c, id] = p.split("/"); return [TABLE[c], id]; };
  return {
    kind:"supabase",
    onAuth(cb) { let last; sb.auth.onAuthStateChange((ev, session) => {
      if (ev === "PASSWORD_RECOVERY") setTimeout(() => onRecovery && onRecovery(), 1200);
      const u = session && session.user, id = u ? u.id : null; if (id === last) return; last = id; setTimeout(() => cb(mapUser(u)), 0); }); },
    async google() { chk(await sb.auth.signInWithOAuth({provider:"google", options:{redirectTo:back, queryParams:{prompt:"select_account"}}})); },
    async signUp(name, email, pw) { const d = chk(await sb.auth.signUp({email, password:pw, options:{data:{full_name:name}, emailRedirectTo:back}}));
      if (d.user && Array.isArray(d.user.identities) && !d.user.identities.length) { const e = new Error("exists"); e.code = "auth/email-already-in-use"; throw e; }
      return {confirm:!d.session}; },
    async signIn(email, pw) { chk(await sb.auth.signInWithPassword({email, password:pw})); },
    async reset(email) { chk(await sb.auth.resetPasswordForEmail(email, {redirectTo:back})); },
    async signOut() { await sb.auth.signOut(); },
    async rename(name) { chk(await sb.auth.updateUser({data:{full_name:name}})); },
    async setPassword(pw) { chk(await sb.auth.updateUser({password:pw})); },
    async verify() { const {data:{user}} = await sb.auth.getUser(); if (user) chk(await sb.auth.resend({type:"signup", email:user.email, options:{emailRedirectTo:back}})); },
    async deleteAccount() { chk(await sb.rpc("delete_my_account")); await sb.auth.signOut(); },
    async get(p) { const [t, id] = split(p); return toCamel(chk(await sb.from(t).select("*").eq("id", id).maybeSingle())); },
    async set(p, data, merge = true) { const [t, id] = split(p), row = Object.assign(toSnake(data), {id});
      if (t === "certificates") { chk(await sb.from(t).insert(row)); return; }
      chk(await sb.from(t).upsert(row, {onConflict:"id"})); },
    async del(p) { const [t, id] = split(p); chk(await sb.from(t).delete().eq("id", id)); },
    async top(coll, field, n) { return chk(await sb.from(TABLE[coll]).select("*").order(Object.keys(toSnake({[field]:1}))[0], {ascending:false}).limit(n)).map(toCamel); },
    async all(coll) { const out = []; for (let from = 0; ; from += 1000) { const rows = chk(await sb.from(TABLE[coll]).select("*").range(from, from + 999)); out.push(...rows.map(toCamel)); if (rows.length < 1000) break; } return out; }
  };
}
function demoBackend() {
  const load = () => { try { return JSON.parse(localStorage.getItem("demo-db") || "{}"); } catch (e) { return {}; } };
  const saveDb = d => localStorage.setItem("demo-db", JSON.stringify(d));
  let cb = null;
  const user = () => { try { return JSON.parse(localStorage.getItem("demo-user") || "null"); } catch (e) { return null; } };
  const setUser = u => { if (u) localStorage.setItem("demo-user", JSON.stringify(u)); else localStorage.removeItem("demo-user"); cb && cb(u); };
  return {
    kind:"demo",
    onAuth: f => { cb = f; setTimeout(() => f(user()), 0); },
    async google() { setUser({uid:"demo-google", name:"Demo User", email:"demo@example.com", photo:"", verified:true, provider:"google.com"}); },
    async signUp(name, email) { setUser({uid:"demo-" + email.replace(/\W/g, ""), name, email, photo:"", verified:false, provider:"email"}); return {confirm:false}; },
    async signIn(email) { setUser({uid:"demo-" + email.replace(/\W/g, ""), name:email.split("@")[0], email, photo:"", verified:true, provider:"email"}); },
    async reset() {}, async signOut() { setUser(null); },
    async rename(name) { const u = user(); if (u) { u.name = name; localStorage.setItem("demo-user", JSON.stringify(u)); } },
    async verify() {}, async setPassword() {},
    async deleteAccount() { const u = user(); if (!u) return; const d = load(); Object.keys(d).forEach(k => { if (k.endsWith("/" + u.uid) || (d[k] && d[k].uid === u.uid)) delete d[k]; }); saveDb(d); setUser(null); },
    async get(p) { return load()[p] ?? (p.startsWith("admins/") ? {role:"owner"} : null); },
    async set(p, data, merge = true) { const d = load(); d[p] = merge ? Object.assign({}, d[p] || {}, data) : data; saveDb(d); },
    async del(p) { const d = load(); delete d[p]; saveDb(d); },
    async top(coll, field, n) { const d = load(); return Object.keys(d).filter(k => k.startsWith(coll + "/") && k.split("/").length === 2).map(k => Object.assign({id:k.split("/")[1]}, d[k])).sort((a, b) => (b[field] || 0) - (a[field] || 0)).slice(0, n); },
    async all(coll) { const d = load(); return Object.keys(d).filter(k => k.startsWith(coll + "/") && k.split("/").length === 2).map(k => Object.assign({id:k.split("/")[1]}, d[k])); }
  };
}

/* ---------------- Landing & auth UI ---------------- */
const AUTH_ERRORS = {
  "auth/invalid-email":"That email address doesn't look right.", "auth/missing-password":"Enter your password.",
  "auth/invalid-credential":"Email or password is incorrect.", "auth/wrong-password":"Email or password is incorrect.", "auth/user-not-found":"No account uses this email. Create one instead.",
  "auth/email-already-in-use":"An account already uses this email. Sign in instead.", "auth/weak-password":"Use a password of at least 8 characters.",
  "auth/too-many-requests":"Too many attempts. Wait a few minutes and try again.", "auth/popup-closed-by-user":"The Google window was closed before signing in.",
  "auth/network-request-failed":"No internet connection. Check your network and try again.", "auth/email-not-confirmed":"Please confirm your email first. Open the link we sent you (check spam too).", "auth/provider-off":"Google sign-in isn't switched on yet. Use email and password for now.", "auth/signup-off":"New sign-ups are closed right now."
};
const errText = e => AUTH_ERRORS[e && e.code] || (e && e.message) || "Something went wrong. Try again.";
let B = null, me = null, isAdmin = false, profile = {};

function fillSiteText() {
  document.querySelectorAll("[data-site]").forEach(el => { el.textContent = SITE[el.dataset.site] || el.textContent; });
  document.querySelectorAll("[data-site-href]").forEach(el => { const v = SITE[el.dataset.siteHref]; if (v) el.href = el.dataset.siteHref === "contactEmail" ? "mailto:" + v : v; else el.hidden = true; });
}
function setMsg(t, kind = "bad") { const m = $("#authmsg"); m.textContent = t || ""; m.className = "authmsg " + (t ? kind : ""); }
function authTab(which) {
  document.querySelectorAll("[data-authtab]").forEach(b => b.setAttribute("aria-selected", b.dataset.authtab === which));
  $("#f-name").closest("label").hidden = which !== "signup"; $("#f-name").required = which === "signup";
  $("#authsubmit").textContent = which === "signup" ? "Create account" : which === "reset" ? "Send reset link" : "Sign in";
  $("#f-pw").closest("label").hidden = which === "reset"; $("#f-pw").required = which !== "reset";
  $("#f-pw").autocomplete = which === "signup" ? "new-password" : "current-password";
  $("#forgot").hidden = which !== "signin"; $("#backsignin").hidden = which !== "reset";
  $("#authform").dataset.mode = which; setMsg("");
}
function wireLanding() {
  fillSiteText();
  document.querySelectorAll("[data-authtab]").forEach(b => b.onclick = () => authTab(b.dataset.authtab));
  $("#forgot").onclick = () => authTab("reset"); $("#backsignin").onclick = () => authTab("signin");
  $("#googlebtn").onclick = async () => { setMsg(""); busy(true); try { await B.google(); } catch (e) { setMsg(errText(e)); } busy(false); };
  $("#authform").onsubmit = async ev => {
    ev.preventDefault(); const mode = ev.target.dataset.mode, email = $("#f-email").value.trim(), pw = $("#f-pw").value, name = $("#f-name").value.trim();
    if (mode === "signup" && pw.length < 8) { setMsg("Use a password of at least 8 characters."); return; }
    busy(true);
    try {
      if (mode === "signup") { const r = await B.signUp(name, email, pw); if (r && r.confirm) { authTab("signin"); $("#f-email").value = email; setMsg("We sent a confirmation link to " + email + ". Open it, then sign in here. Check your spam folder too.", "ok"); } }
      else if (mode === "reset") { await B.reset(email); setMsg("If an account uses this email, a reset link is on its way. Check your inbox and spam folder.", "ok"); }
      else await B.signIn(email, pw);
    } catch (e) { setMsg(errText(e)); }
    busy(false);
  };
  authTab("signin");
  if (DEMO) $("#demonote").hidden = false;
  if (!configured && !isLocal) { $("#setupnote").hidden = false; $("#authcard").classList.add("disabled"); }
}
function busy(on) { $("#authcard").classList.toggle("busy", on); document.querySelectorAll("#authcard button, #authcard input").forEach(x => x.disabled = on); }

/* ---------------- Progress sync ---------------- */
let pushT = null, appLoaded = false;
function hookStorage() {
  const orig = Storage.prototype.setItem;
  Storage.prototype.setItem = function (k, v) { orig.call(this, k, v); if (this === window.localStorage && SYNC_KEYS.includes(k)) schedulePush(); };
  document.addEventListener("visibilitychange", () => { if (document.visibilityState === "hidden") pushNow(); });
}
function schedulePush() { clearTimeout(pushT); pushT = setTimeout(pushNow, 1500); }
async function pushNow() {
  clearTimeout(pushT); if (!me || !window.DPL) return;
  const keys = {}; SYNC_KEYS.forEach(k => { const v = localStorage.getItem(k); if (v !== null) keys[k] = v; });
  const st = window.DPL.stats(), score = st.theory + st.practice;
  try {
    await B.set(`progress/${me.uid}`, {keys, updatedAt:nowIso()}, false);
    await B.set(`users/${me.uid}`, {lastActive:nowIso(), score, theory:st.theory, practice:st.practice, per:compactPer(st.per)});
    if (profile.showOnBoard !== false) await B.set(`leaderboard/${me.uid}`, {name:boardName(), score, updatedAt:nowIso()}, false);
  } catch (e) { console.warn("Sync failed", e); }
}
const compactPer = per => { const o = {}; for (const t in per) for (const v in per[t]) o[t + "_" + v] = per[t][v].theory[0] + per[t][v].practice[0]; return o; };
const boardName = () => { const n = (profile.displayName || me.name || "Learner").trim(); const parts = n.split(/\s+/); return (parts[0] + (parts[1] ? " " + parts[1][0] + "." : "")).slice(0, 40); };

async function restoreProgress() {
  const cloud = await B.get(`progress/${me.uid}`);
  const owner = localStorage.getItem("dpl-owner");
  const localDone = (() => { try { return JSON.parse(localStorage.getItem("dpl-progress") || "{}").done || []; } catch (e) { return []; } })();
  if (owner && owner !== me.uid) SYNC_KEYS.forEach(k => localStorage.removeItem(k));
  if (cloud && cloud.keys) for (const k of SYNC_KEYS) if (cloud.keys[k] !== undefined) localStorage.setItem(k, cloud.keys[k]);
  if ((!owner || owner === me.uid) && localDone.length) {
    try { const p = JSON.parse(localStorage.getItem("dpl-progress") || "{}"); p.done = [...new Set([...(p.done || []), ...localDone])]; localStorage.setItem("dpl-progress", JSON.stringify(p)); } catch (e) {}
  }
  localStorage.setItem("dpl-owner", me.uid);
}
function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => rej(new Error("Couldn't load " + src)); document.body.appendChild(s); }); }

async function enterApp(u) {
  me = u;
  $("#landing").hidden = true; $("#loading").hidden = false;
  try {
    profile = (await B.get(`users/${u.uid}`)) || {};
    if (!profile.createdAt) { profile = {createdAt:nowIso(), displayName:u.name || u.email.split("@")[0], email:u.email, showOnBoard:true, provider:u.provider}; await B.set(`users/${u.uid}`, profile); }
    await B.set(`users/${u.uid}`, {email:u.email, photo:u.photo || "", lastActive:nowIso()});
    try { isAdmin = !!(await B.get(`admins/${u.uid}`)); } catch (e) { isAdmin = false; }
    await restoreProgress();
    if (!appLoaded) {
      $("#appwrap").hidden = false;
      await loadScript("vendor/hyperformula.full.min.js"); await loadScript("vendor/sql-asm.js"); await loadScript("js/app.bundle.js");
      appLoaded = true; hookStorage();
    }
    $("#loading").hidden = true; $("#appwrap").hidden = false;
    renderUserMenu(); schedulePush();
  } catch (e) {
    $("#loading").hidden = true; $("#landing").hidden = false; setMsg("Couldn't open your account: " + errText(e));
  }
}

/* ---------------- User menu & panels ---------------- */
const initials = n => (n || "?").trim().split(/\s+/).map(w => w[0]).slice(0, 2).join("").toUpperCase();
function renderUserMenu() {
  const slot = $("#userslot"); if (!slot) return;
  const name = profile.displayName || me.name || me.email;
  slot.innerHTML = `<div class="umenu"><button type="button" class="avatar" id="avatarbtn" aria-haspopup="true" aria-expanded="false" title="${esc(name)}">${me.photo ? `<img src="${esc(me.photo)}" alt="" referrerpolicy="no-referrer">` : esc(initials(name))}</button>
    <div class="pop umenu-pop" id="umenupop" hidden><div class="who"><b>${esc(name)}</b><small>${esc(me.email)}</small></div>
    <button type="button" data-panel="profile">My profile</button><button type="button" data-panel="certs">Certificates</button><button type="button" data-panel="board">Leaderboard</button>
    ${isAdmin ? `<button type="button" data-panel="admin">Admin dashboard</button>` : ""}<hr><button type="button" id="signoutbtn">Sign out</button></div></div>`;
}
document.addEventListener("click", e => {
  const ab = e.target.closest("#avatarbtn"), pop = $("#umenupop");
  if (ab) { pop.hidden = !pop.hidden; ab.setAttribute("aria-expanded", !pop.hidden); return; }
  if (pop && !e.target.closest(".umenu")) pop.hidden = true;
  const pb = e.target.closest("[data-panel]"); if (pb) { if (pop) pop.hidden = true; openPanel(pb.dataset.panel); return; }
  if (e.target.closest("#signoutbtn")) { doSignOut(); return; }
  if (e.target.closest("#panelclose") || e.target.id === "panelwrap") closePanel();
});
document.addEventListener("keydown", e => { if (e.key === "Escape") closePanel(); });
async function doSignOut() { await pushNow(); SYNC_KEYS.concat("dpl-owner", "dpl-view").forEach(k => localStorage.removeItem(k)); await B.signOut(); location.reload(); }

function openPanel(kind) {
  const w = $("#panelwrap"); w.hidden = false; $("#paneltitle").textContent = {profile:"My profile", certs:"Certificates", board:"Leaderboard", admin:"Admin dashboard"}[kind];
  $("#panelbody").innerHTML = `<p class="muted">Loading…</p>`; document.body.classList.add("noscroll");
  ({profile:panelProfile, certs:panelCerts, board:panelBoard, admin:panelAdmin})[kind]().catch(e => { $("#panelbody").innerHTML = `<p class="exr err">Couldn't load: ${esc(errText(e))}</p>`; });
}
function closePanel() { const w = $("#panelwrap"); if (w && !w.hidden) { w.hidden = true; document.body.classList.remove("noscroll"); } }

async function panelProfile() {
  const st = window.DPL.stats(), T = window.DPL.TRACKS;
  $("#panelbody").innerHTML = `<form id="profform" class="pform">
    <label><span>Name on certificates and leaderboard</span><input id="p-name" maxlength="60" required value="${esc(profile.displayName || me.name || "")}"></label>
    <label class="check"><input type="checkbox" id="p-board" ${profile.showOnBoard !== false ? "checked" : ""}> Show me on the leaderboard (first name and initial only)</label>
    <div class="row"><button class="btn primary" type="submit">Save</button><span id="p-msg" class="muted"></span></div></form>
    <div class="pinfo"><p><b>Email:</b> ${esc(me.email)} ${me.provider === "email" ? (me.verified ? `<span class="pill ok">verified</span>` : `<span class="pill">not verified</span> <button type="button" class="btn small" id="p-verify">Send verification email</button>`) : `<span class="pill">Google account</span>`}</p>
    <p><b>Member since:</b> ${esc((profile.createdAt || "").slice(0, 10))}</p></div>
    <h4>Your progress</h4><table class="mini"><tr><th>Track</th><th>Basic</th><th>Advanced</th></tr>${Object.keys(T).map(t => `<tr><td>${T[t]}</td>${["basic","advanced"].map(v => { const x = st.per[t][v]; return `<td>${x.theory[0] + x.practice[0]} / ${x.theory[1] + x.practice[1]}</td>`; }).join("")}</tr>`).join("")}</table>
    ${me.provider === "email" ? `<h4>Password</h4><form id="pwform" class="pform"><label><span>New password (at least 8 characters)</span><input id="p-pw" type="password" autocomplete="new-password" minlength="8" required></label>
    <div class="row"><button class="btn small" type="submit">Change password</button><span id="pw-msg" class="muted"></span></div></form>` : ""}
    <h4>Your data</h4><p class="muted">Download a copy of everything stored about you, or delete your account. Deleting removes your progress, leaderboard entry and certificates for good.</p>
    <div class="row"><button type="button" class="btn small" id="p-export">Download my data</button><button type="button" class="btn small danger" id="p-delete">Delete my account</button><span id="del-msg" class="muted"></span></div>`;
  $("#profform").onsubmit = async ev => { ev.preventDefault(); const n = $("#p-name").value.trim(); profile.displayName = n; profile.showOnBoard = $("#p-board").checked;
    await B.set(`users/${me.uid}`, {displayName:n, showOnBoard:profile.showOnBoard}); await B.rename(n);
    if (!profile.showOnBoard) { try { await B.del(`leaderboard/${me.uid}`); } catch (e) {} } await pushNow(); renderUserMenu(); $("#p-msg").textContent = "Saved."; };
  const vb = $("#p-verify"); if (vb) vb.onclick = async () => { await B.verify(); vb.textContent = "Sent. Check your inbox."; vb.disabled = true; };
  $("#p-export").onclick = async () => { const data = {user:await B.get(`users/${me.uid}`), progress:await B.get(`progress/${me.uid}`), leaderboard:await B.get(`leaderboard/${me.uid}`)};
    downloadBlob(new Blob([JSON.stringify(data, null, 2)], {type:"application/json"}), "my-data-practice-lab-data.json"); };
  const pf = $("#pwform"); if (pf) pf.onsubmit = async ev => { ev.preventDefault(); const pw = $("#p-pw").value; if (pw.length < 8) { $("#pw-msg").textContent = "Use at least 8 characters."; return; }
    try { await B.setPassword(pw); $("#pw-msg").textContent = "Password changed."; $("#p-pw").value = ""; } catch (e) { $("#pw-msg").textContent = errText(e); } };
  const del = $("#p-delete"); let armed = false;
  del.onclick = async () => { if (!armed) { armed = true; del.textContent = "Click again to delete for good"; $("#del-msg").textContent = "This can't be undone."; setTimeout(() => { armed = false; del.textContent = "Delete my account"; $("#del-msg").textContent = ""; }, 6000); return; }
    del.disabled = true; $("#del-msg").textContent = "Deleting…";
    try { clearTimeout(pushT); const keep = me; me = null; try { await B.deleteAccount(); } catch (e) { me = keep; throw e; } SYNC_KEYS.concat("dpl-owner", "dpl-view").forEach(k => localStorage.removeItem(k)); location.reload(); } catch (e) { del.disabled = false; $("#del-msg").textContent = errText(e); } };
}

/* Certificates */
const CERT_LEVELS = [["excel","basic"],["excel","advanced"],["sql","basic"],["sql","advanced"],["pbi","basic"],["pbi","advanced"]];
async function panelCerts() {
  const st = window.DPL.stats(), T = window.DPL.TRACKS, certs = profile.certs || {};
  $("#panelbody").innerHTML = `<p class="muted">Finish every theory chapter and practice lesson of a level to earn its certificate. Each certificate has a public verification link.</p><div class="certs">${CERT_LEVELS.map(([t, v]) => {
    const x = st.per[t][v], d = x.theory[0] + x.practice[0], n = x.theory[1] + x.practice[1], pct = n ? Math.round(d / n * 100) : 0, key = t + "_" + v;
    return `<div class="certrow"><div><b>${T[t]} · ${v === "basic" ? "Basic" : "Advanced"}</b><small>${d} of ${n} done · theory ${x.theory[0]}/${x.theory[1]}, practice ${x.practice[0]}/${x.practice[1]}</small><div class="bar"><i style="width:${pct}%"></i></div></div>
    ${pct === 100 ? `<button type="button" class="btn primary small" data-cert="${key}">${certs[key] ? "Download again" : "Get certificate"}</button>` : `<span class="pill">${pct}%</span>`}</div>`; }).join("")}</div><div id="certout"></div>`;
  document.querySelectorAll("[data-cert]").forEach(b => b.onclick = () => issueCert(b.dataset.cert));
}
const rid = () => Array.from(crypto.getRandomValues(new Uint8Array(10)), x => "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"[x % 32]).join("");
async function issueCert(key) {
  const [t, v] = key.split("_"), st = window.DPL.stats(), x = st.per[t][v];
  if (x.theory[0] < x.theory[1] || x.practice[0] < x.practice[1]) return;
  profile.certs = profile.certs || {};
  let id = profile.certs[key], cert = id ? await B.get(`certificates/${id}`) : null;
  if (!cert) { id = rid(); cert = {uid:me.uid, name:profile.displayName || me.name || me.email, track:t, level:v, issuedAt:nowIso()}; await B.set(`certificates/${id}`, cert, false); profile.certs[key] = id; await B.set(`users/${me.uid}`, {certs:profile.certs}); }
  const url = new URL("verify.html?id=" + id, location.href).href;
  const canvas = drawCert(cert, id, url);
  canvas.toBlob(b => downloadBlob(b, `certificate-${t}-${v}-${id}.png`), "image/png");
  $("#certout").innerHTML = `<div class="callout note"><b>Certificate downloaded</b><p>Verification link: <a href="${esc(url)}" target="_blank" rel="noopener">${esc(url)}</a></p></div>`;
  window.DPL.confetti(120);
}
function drawCert(c, id, url) {
  const W = 2000, H = 1414, cv = document.createElement("canvas"); cv.width = W; cv.height = H; const g = cv.getContext("2d");
  const accent = "#1F5FBF", ink = "#172033", muted = "#5A6478";
  g.fillStyle = "#FFFFFF"; g.fillRect(0, 0, W, H);
  g.strokeStyle = accent; g.lineWidth = 14; g.strokeRect(50, 50, W - 100, H - 100);
  g.strokeStyle = "#DCE1E8"; g.lineWidth = 3; g.strokeRect(84, 84, W - 168, H - 168);
  g.fillStyle = accent; g.fillRect(84, 84, W - 168, 16);
  const T = window.DPL.TRACKS, F = '"IBM Plex Sans", "Segoe UI", Arial, sans-serif';
  g.textAlign = "center";
  g.fillStyle = muted; g.font = `600 34px ${F}`; g.fillText(SITE.name.toUpperCase(), W / 2, 240);
  g.fillStyle = ink; g.font = `600 92px ${F}`; g.fillText("Certificate of Completion", W / 2, 380);
  g.fillStyle = muted; g.font = `400 40px ${F}`; g.fillText("This certifies that", W / 2, 500);
  g.fillStyle = ink; let fs = 104; g.font = `600 ${fs}px ${F}`; while (g.measureText(c.name).width > W - 400 && fs > 50) { fs -= 4; g.font = `600 ${fs}px ${F}`; } g.fillText(c.name, W / 2, 640);
  g.strokeStyle = "#DCE1E8"; g.lineWidth = 3; g.beginPath(); g.moveTo(W / 2 - 520, 680); g.lineTo(W / 2 + 520, 680); g.stroke();
  g.fillStyle = muted; g.font = `400 40px ${F}`; g.fillText("has successfully completed", W / 2, 770);
  g.fillStyle = accent; g.font = `600 64px ${F}`; g.fillText(`${T[c.track]} · ${c.level === "basic" ? "Basic" : "Advanced"} level`, W / 2, 870);
  g.fillStyle = muted; g.font = `400 34px ${F}`; g.fillText("all theory chapters and hands-on practice lessons", W / 2, 935);
  g.textAlign = "left"; g.fillStyle = ink; g.font = `600 32px ${F}`; g.fillText("Date of issue", 200, 1140); g.fillText("Certificate ID", W - 760, 1140);
  g.font = `400 32px ${F}`; g.fillStyle = muted; g.fillText(new Date(c.issuedAt).toLocaleDateString("en-GB", {day:"2-digit", month:"long", year:"numeric"}), 200, 1190); g.fillText(id, W - 760, 1190);
  g.font = `400 26px ${F}`; g.textAlign = "center"; g.fillText("Verify at " + url, W / 2, 1290);
  return cv;
}

/* Leaderboard */
async function panelBoard() {
  const rows = await B.top("leaderboard", "score", 25);
  $("#panelbody").innerHTML = `<p class="muted">Top learners by chapters and lessons completed. ${profile.showOnBoard === false ? "You're hidden: turn it on in My profile." : "Hide yourself any time in My profile."}</p>
  ${rows.length ? `<table class="mini board"><tr><th>#</th><th>Learner</th><th>Score</th></tr>${rows.map((r, i) => `<tr class="${r.id === me.uid ? "me" : ""}"><td>${i + 1}</td><td>${esc(r.name)}${r.id === me.uid ? " (you)" : ""}</td><td>${r.score}</td></tr>`).join("")}</table>` : `<p>No one on the board yet. Complete a lesson to be first.</p>`}`;
}

/* Admin */
async function panelAdmin() {
  if (!isAdmin) { $("#panelbody").innerHTML = "<p>Admins only.</p>"; return; }
  const users = (await B.all("users")).sort((a, b) => (b.createdAt || "").localeCompare(a.createdAt || ""));
  const day = d => (d || "").slice(0, 10), today = new Date(), ago = n => { const d = new Date(today); d.setDate(d.getDate() - n); return d.toISOString().slice(0, 10); };
  const active7 = users.filter(u => day(u.lastActive) >= ago(7)).length, new30 = users.filter(u => day(u.createdAt) >= ago(30)).length;
  const avg = users.length ? Math.round(users.reduce((s, u) => s + (u.score || 0), 0) / users.length) : 0;
  const days = Array.from({length:30}, (_, i) => ago(29 - i)), counts = days.map(d => users.filter(u => day(u.createdAt) === d).length), mx = Math.max(1, ...counts);
  const bars = days.map((d, i) => `<rect x="${i * 20 + 2}" y="${100 - counts[i] / mx * 90}" width="16" height="${counts[i] / mx * 90}" rx="2"><title>${d}: ${counts[i]}</title></rect>`).join("");
  $("#panelbody").innerHTML = `<div class="stats"><div class="stat"><small>Total users</small><b>${users.length}</b></div><div class="stat"><small>Active last 7 days</small><b>${active7}</b></div><div class="stat"><small>New last 30 days</small><b>${new30}</b></div><div class="stat"><small>Average score</small><b>${avg}</b></div></div>
    <h4>Sign-ups, last 30 days</h4><svg class="adminchart" viewBox="0 0 600 112" role="img" aria-label="Sign-ups per day">${bars}<line x1="0" x2="600" y1="100.5" y2="100.5"/><text x="0" y="111">${ago(29)}</text><text x="600" y="111" text-anchor="end">today</text></svg>
    <div class="row"><input id="a-q" type="search" placeholder="Search name or email" aria-label="Search users"><button type="button" class="btn small" id="a-csv">Download CSV</button></div>
    <div class="tscroll"><table class="grid" id="a-table"></table></div>`;
  const cols = ["displayName","email","createdAt","lastActive","score","theory","practice"];
  const draw = q => { const r = users.filter(u => !q || ((u.displayName || "") + " " + (u.email || "")).toLowerCase().includes(q.toLowerCase()));
    $("#a-table").innerHTML = `<thead><tr><th>Name</th><th>Email</th><th>Joined</th><th>Last active</th><th>Score</th><th>Theory</th><th>Practice</th></tr></thead><tbody>${r.map(u => `<tr><td>${esc(u.displayName)}</td><td>${esc(u.email)}</td><td>${esc(day(u.createdAt))}</td><td>${esc(day(u.lastActive))}</td><td class="n">${u.score || 0}</td><td class="n">${u.theory || 0}</td><td class="n">${u.practice || 0}</td></tr>`).join("")}</tbody>`; };
  draw(""); $("#a-q").oninput = e => draw(e.target.value);
  $("#a-csv").onclick = () => { const q = s => `"${String(s ?? "").replace(/"/g, '""')}"`; const csv = [cols.join(",")].concat(users.map(u => cols.map(c => q(u[c])).join(","))).join("\n");
    downloadBlob(new Blob([csv], {type:"text/csv"}), `users-${new Date().toISOString().slice(0, 10)}.csv`); };
}
function downloadBlob(blob, name) { const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000); }

onRecovery = () => { if (!me || !appLoaded || !window.DPL || $("#loading").hidden === false) { setTimeout(onRecovery, 500); return; } openPanel("profile"); setTimeout(() => { const f = $("#p-pw"); if (f) { f.focus(); $("#pw-msg").textContent = "Choose your new password."; } }, 600); };

/* ---------------- Start ---------------- */
let wired = false; function wireLandingSafe() { if (!wired) { wired = true; wireLanding(); } }
(async function start() {
  try { B = DEMO ? demoBackend() : configured ? await supabaseBackend() : demoBackend(); }
  catch (e) { $("#loading").hidden = true; $("#landing").hidden = false; wireLandingSafe(); setMsg("Couldn't reach the login service. Check your internet connection and reload."); return; }
  wireLandingSafe();
  B.onAuth(u => { if (u) enterApp(u); else { if (appLoaded) location.reload(); $("#loading").hidden = true; $("#landing").hidden = false; } });
})();
