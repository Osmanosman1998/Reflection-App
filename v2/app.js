// My Shift Reflection, build 2.0. Three versions chosen by the intern's independence level.

const app = document.getElementById("app");
const PREVIEW = !CONFIG.profileUrl;

let profile = null;     // { token, name, site, level, rotation, targets[] }
let version = "middle"; // supported | middle | independent
let steps = [];
let at = 0;
let answers = {};
let nudged = {};
let dir = 1;
function go(d) { dir = d; at += d; draw(); }
// Steps that only show in some cases
const SKIP = { safe2: () => !answers.safe || answers.safe === "Yes" };

// ---------- Small helpers ----------
function el(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") e.className = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else if (v !== false && v != null) e.setAttribute(k, v);
  }
  kids.flat().forEach((k) => k != null && e.append(k.nodeType ? k : document.createTextNode(k)));
  return e;
}
const words = (s) => (s || "").trim().split(/\s+/).filter(Boolean).length;
const firstName = () => profile.name.split(" ")[0];

// Listen only reads when the intern taps it. Tapping again stops it.
function setListenButton(on) {
  const b = document.getElementById("listenBtn");
  if (!b) return;
  b.setAttribute("aria-pressed", String(on));
  b.setAttribute("aria-label", on ? "Stop reading" : "Read this out loud");
  b.querySelector("span").textContent = on ? "Stop" : "Listen";
  b.firstChild.textContent = on ? "⏹ " : "🔊 ";
}
function stopSpeaking() {
  if ("speechSynthesis" in window) speechSynthesis.cancel();
  setListenButton(false);
}
function speak(text) {
  if (!("speechSynthesis" in window)) return;
  speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(text);
  u.lang = "en-GB"; u.rate = 0.9;
  u.onend = () => setListenButton(false);
  u.onerror = () => setListenButton(false);
  setListenButton(true);
  speechSynthesis.speak(u);
}
function readScreen() {
  const bits = [];
  app.querySelectorAll("h1, .lead, .target-text, .starter, .prompts li, .chip, .tile, .scale-btn, .face b, .option, .tip, .hint:not([hidden]) > span:first-child, .remember")
    .forEach((n) => bits.push(n.textContent.replace("✓ Picked", "").trim()));
  speak(bits.join(". "));
}

// ---------- Getting the intern's profile ----------
// Card codes are lower case letters and numbers. Tidies anything a camera or a person adds.
const cleanToken = (t) => String(t || "").trim().toLowerCase().replace(/[^a-z0-9-]/g, "").slice(0, 20) || null;
function tokenFromLink() {
  const q = new URLSearchParams(location.search).get("id");
  const h = location.hash.replace("#", "");
  return cleanToken(q || h);
}
function remembered() {
  try { return cleanToken(localStorage.getItem("internToken")); } catch (e) { return null; }
}
const cacheKey = (t) => "internProfile:" + t;
// Details saved on the phone are only used offline for 7 days, so an intern who has left
// cannot keep reflecting from an old copy.
const CACHE_DAYS = 7;
function cacheProfile(p) {
  try { localStorage.setItem(cacheKey(p.token), JSON.stringify({ ...p, savedAt: Date.now() })); } catch (e) { /* storage blocked */ }
}
function cachedProfile(t) {
  try {
    const c = JSON.parse(localStorage.getItem(cacheKey(t)));
    return c && Date.now() - (c.savedAt || 0) < CACHE_DAYS * 86400000 ? c : null;
  } catch (e) { return null; }
}
function dropCachedProfile(t) {
  try { localStorage.removeItem(cacheKey(t)); } catch (e) { /* storage blocked */ }
}

// fetch that gives up after a set time, so a slow flow never leaves an intern stuck
async function fetchWithTimeout(url, options, ms) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), ms);
  try { return await fetch(url, { ...options, signal: ctrl.signal }); }
  finally { clearTimeout(timer); }
}
function remember(t) {
  try { localStorage.setItem("internToken", t); } catch (e) { /* storage blocked */ }
}

// Turns a row from the Excel "Questions" table into what the app needs
const QTYPES = { "Typed answer": "text", "Tap one option": "choose", "1 to 5 scale": "scale" };
// Two support levels: Pre-SIP ("pre") and SIP ("sip"). Old numbers still work: 1 is Pre-SIP, 2 and 3 are SIP.
const levelKey = (v) => (/pre/i.test(String(v)) || String(v).trim() === "1" ? "pre" : "sip");
const levelName = (k) => (k === "pre" ? "Pre-SIP" : "SIP");
const WHO = { "Everyone": ["pre", "sip"], "Pre-SIP only": ["pre"], "SIP only": ["sip"],
              "Level 1 only": ["pre"], "Level 2 only": ["sip"], "Level 3 only": ["sip"],
              "Levels 1 and 2": ["pre", "sip"], "Levels 2 and 3": ["sip"] };
function normaliseQuestion(q) {
  // Accepts a raw Excel row (ID, Type, Question, Options, Who sees it) or the preview shape
  const raw = { id: q.ID ?? q.id, type: q.Type ?? q.type, text: q.Question ?? q.text, site: q.Site ?? q.site,
                options: q.Options ?? q.options, who: q["Who sees it"] ?? q.who };
  const type = QTYPES[String(raw.type || "").trim()] || raw.type;
  if (!raw.text || !["text", "choose", "scale"].includes(type)) return null;
  const options = String(raw.options || "").split(";").map((o) => o.trim()).filter(Boolean);
  if (type === "choose" && !options.length) return null;
  return { id: String(raw.id), type, text: String(raw.text), options, levels: WHO[String(raw.who || "").trim()] || ["pre", "sip"],
           site: String(raw.site || "").trim(), required: type !== "text" };
}

// Power Automate sends the intern's Excel row as it is. This turns it into the app's shape.
function fromExcel(p) {
  if (!p || !p.intern) return p;
  const i = p.intern;
  return {
    name: i["Name"], site: i["Site"], level: levelKey(i["Support level"]), rotation: i["Current rotation"], internId: i["Intern ID"] || "",
    targets: CONFIG.areas.map((a) => i[a.column]),
    rotations: (p.rotations || []).map((r) => (typeof r === "string" ? r : r["Rotation"])),
    questions: p.questions || [],
    halfTerms: p.halfTerms || [],
    doneDates: p.doneDates || [],
    wordBanks: p.wordBanks || []
  };
}

async function loadProfile(token) {
  let p;
  if (PREVIEW) {
    p = typeof DEMO_PROFILES !== "undefined" && DEMO_PROFILES[token];
    if (!p) return null;
    p = { ...p, questions: p.questions || (typeof DEMO_QUESTIONS !== "undefined" ? DEMO_QUESTIONS : []),
          halfTerms: p.halfTerms || (typeof DEMO_HALFTERMS !== "undefined" ? DEMO_HALFTERMS : []),
          wordBanks: p.wordBanks || (typeof DEMO_WORDBANKS !== "undefined" ? DEMO_WORDBANKS : []) };
  } else {
    // Power Automate checks the code against Excel and returns this intern only.
    // text/plain keeps this a "simple" request, so the browser sends it straight to the flow.
    try {
      const r = await fetchWithTimeout(CONFIG.profileUrl, {
        method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ token })
      }, CONFIG.profileTimeoutMs || 30000);
      if (r.status === 404 || r.status === 403) { dropCachedProfile(token); return null; }
      if (!r.ok) throw new Error("flow error " + r.status);
      p = fromExcel(await r.json());
    } catch (e) {
      // No signal, slow flow or a flow error: use the details saved last time, if there are any
      const c = cachedProfile(token);
      return c ? { ...c, fromCache: true } : { offline: true };
    }
  }
  if (!p || !p.name) return null;
  const fresh = {
    token,
    internId: String(p.internId || ""),
    name: String(p.name),
    site: p.site || "",
    level: levelKey(p.level),
    rotation: p.rotation || "",
    targets: CONFIG.areas.map((a, i) => (p.targets && p.targets[i]) || ""),
    rotations: (p.rotations || []).filter(Boolean),
    questions: (p.questions || []).map(normaliseQuestion).filter(Boolean),
    halfTerms: (p.halfTerms || []).map(readHalfTerm).filter(Boolean),
    doneDates: (p.doneDates || []).map((d) => { const x = readDate(d); return x ? isoDay(x) : null; }).filter(Boolean),
    wordBanks: readWordBanks(p.wordBanks || [])
  };
  if (!PREVIEW) cacheProfile(fresh);
  return fresh;
}

// ---------- Half-term check-in ----------
// Excel can send dates as serial numbers or as text, so both are read
function readDate(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  if (!isNaN(n) && String(v).trim() !== "") return new Date(Math.round((n - 25569) * 86400000));
  const d = new Date(String(v).trim().replace(" ", "T").replace(/(T\d\d:\d\d:\d\d)$/, "$1Z"));
  return isNaN(d) ? null : d;
}
function readHalfTerm(h) {
  const name = h["Half term"] ?? h.name, start = readDate(h["Starts"] ?? h.start), end = readDate(h["Ends"] ?? h.end);
  return name && start && end ? { name: String(name), start: start.getTime(), end: end.getTime() } : null;
}
const dayOnly = (ms) => { const d = new Date(ms); return Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); };
function todayUtc() { const d = new Date(); return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()); }
const ciKey = (p, term) => "checkIn:" + (p.internId || p.token) + ":" + term;
function currentCheckIn(p) {
  const t = todayUtc();
  const term = (p.halfTerms || []).find((h) => t <= dayOnly(h.end) && t >= dayOnly(h.end) - ((CONFIG.checkInDays || 7) - 1) * DAY);
  if (!term) return null;
  try { if (localStorage.getItem(ciKey(p, term.name))) return null; } catch (e) { /* storage blocked: ask anyway */ }
  return term.name;
}
function markCheckInDone() {
  if (!checkInTerm || PREVIEW) return;
  try { localStorage.setItem(ciKey(profile, checkInTerm), "done"); } catch (e) { /* storage blocked */ }
}
let checkInTerm = null;

async function start(token) {
  answers = {}; nudged = {};
  if (token) app.replaceChildren(el("section", { class: "screen" }, el("p", { class: "lead" }, "Loading...")));
  profile = token ? await loadProfile(token) : null;
  // Remember the card even when there is no signal, so the next open works without scanning
  if (profile && profile.offline) { if (!PREVIEW) remember(token); return showOffline(token); }
  if (!profile) return showNoCard();
  if (!PREVIEW) { remember(token); if (!profile.fromCache) moveWaitingToCard(profile); }
  version = CONFIG.versionForLevel[profile.level] || "independent";
  answers.rotation = profile.rotation;
  // One reflection per day: if today and the last shift are both done, there is nothing to do
  const open = dayChoices().filter((c) => !c.done);
  if (!open.length) { steps = ["alldone"]; at = 0; return draw(); }
  answers.day = open[0].key;
  const targetSteps = [];
  CONFIG.areas.forEach((a, i) => {
    if (!profile.targets[i]) return;
    if (!hasSeen(profile.targets[i])) targetSteps.push("s" + i);
    targetSteps.push("t" + i);
  });
  steps = ["hello", "day", "rotation", "rating", "safe", "safe2", ...targetSteps, "well", "hard", "help", "next"];
  if (version === "independent") steps.push("learned");
  profile.questions
    .filter((q) => q.levels.includes(profile.level))
    .filter((q) => !q.site || /^all/i.test(q.site) || q.site.toLowerCase() === profile.site.toLowerCase())
    .forEach((q) => steps.push("x:" + q.id));
  // Half-term check-in: last 7 days of a half term, once per half term on this phone
  checkInTerm = currentCheckIn(profile);
  if (checkInTerm) {
    const ci = CONFIG.areas.map((a, i) => "c" + i).filter((c, i) => profile.targets[i]);
    if (ci.length) steps.push("ci_intro", ...ci);
  }
  steps.push("thanks");
  at = 0;
  draw();
}

function showOffline(token) {
  app.replaceChildren(
    el("section", { class: "screen" },
      el("div", { class: "hero-face", "aria-hidden": "true" }, "📶"),
      el("h1", {}, "No internet"),
      el("p", { class: "lead" }, "Check your Wi-Fi or data, then try again."),
      el("div", { class: "nav" }, el("button", { class: "btn main", type: "button", onclick: () => start(token) }, "Try again"))
    )
  );
}

function forget() {
  try { localStorage.removeItem("internToken"); } catch (e) { /* blocked */ }
  showNoCard();
}

function showNoCard() {
  app.replaceChildren(
    el("section", { class: "screen" },
      el("div", { class: "hero-face", "aria-hidden": "true" }, "🪪"),
      el("h1", {}, "Scan your QR card"),
      el("p", { class: "lead" }, "Use the camera on your phone to scan your own card. If you have lost it, ask your job coach.")
    )
  );
}

// ---------- One reflection per day, with a catch-up for the last shift ----------
const pad2 = (n) => String(n).padStart(2, "0");
const isoDay = (d) => `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
const localDay = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
function lastShift() {
  const d = new Date();
  do { d.setDate(d.getDate() - 1); } while (d.getDay() === 0 || d.getDay() === 6);
  return d;
}
const niceDay = (d) => d.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" });
// "today", or "on Thursday" when catching up on the last shift
function dayWord() {
  if (!answers.day || answers.day === localDay(new Date())) return "today";
  const [y, m, d] = answers.day.split("-").map(Number);
  return "on " + new Date(y, m - 1, d).toLocaleDateString("en-GB", { weekday: "long" });
}
const MEMORY = new Set();   // used when the phone blocks storage, and in the preview
const doneKey = (day) => "done:" + (profile.internId || profile.token) + ":" + day;
function isDone(day) {
  if ((profile.doneDates || []).includes(day) || MEMORY.has(doneKey(day))) return true;
  if (PREVIEW) return false;
  try { return !!localStorage.getItem(doneKey(day)); } catch (e) { return false; }
}
function markDone(day) {
  MEMORY.add(doneKey(day));
  if (!PREVIEW) try { localStorage.setItem(doneKey(day), "1"); } catch (e) { /* storage blocked */ }
}
function dayChoices() {
  const t = new Date(), l = lastShift();
  return [
    { key: localDay(t), label: "Today", when: niceDay(t) },
    { key: localDay(l), label: "My last shift", when: niceDay(l) }
  ].map((c) => ({ ...c, done: isDone(c.key) }));
}

// ---------- Starting view: the first time an intern sees a new target ----------
const seenKey = (text) => "seen:" + (profile.internId || profile.token) + ":" + text.toLowerCase().replace(/\s+/g, " ").trim();
function hasSeen(text) {
  if (MEMORY.has(seenKey(text))) return true;
  if (PREVIEW) return false;
  try { return !!localStorage.getItem(seenKey(text)); } catch (e) { return false; }
}
function markSeen() {
  profile.targets.forEach((t, i) => {
    if (!t || !answers["start" + (i + 1)]) return;
    MEMORY.add(seenKey(t));
    if (!PREVIEW) try { localStorage.setItem(seenKey(t), "1"); } catch (e) { /* storage blocked */ }
  });
}

// ---------- Word banks for each rotation and PfA area ----------
// From the Word banks tab: one row per rotation and area, phrases separated by ;
function readWordBanks(rows) {
  const out = {};
  rows.forEach((r) => {
    const rot = String(r["Rotation"] ?? r.rotation ?? "").trim(), area = String(r["Area"] ?? r.area ?? "").trim();
    const i = CONFIG.areas.findIndex((a) => a.name.toLowerCase() === area.toLowerCase());
    const phrases = String(r["Phrases"] ?? r.phrases ?? "").split(";").map((x) => x.trim().replace(/\.$/, "")).filter(Boolean);
    const sc = String(r["Score"] ?? r.score ?? "").replace(/\s+/g, " ").trim().toLowerCase();
    const band = sc === "1 to 2" ? "low" : sc === "3" ? "mid" : sc === "4 to 5" ? "high" : "all";
    if (rot && i >= 0 && phrases.length) {
      const a = (out[rot.toLowerCase()] = out[rot.toLowerCase()] || {});
      (a[i] = a[i] || {})[band] = phrases;
    }
  });
  return out;
}
// Score groups: 1 to 2 = low, 3 = mid, 4 to 5 = high
const bandOf = (score) => (score <= 2 ? "low" : score === 3 ? "mid" : "high");
// Phrases for "What did I do?": this rotation's phrases for the score group, then general phrases for the score group
function targetBank(i, band) {
  const rot = profile.wordBanks[String(answers.rotation || "").toLowerCase()];
  const mine = rot && rot[i] ? [...(rot[i][band] || []), ...(rot[i].all || [])] : [];
  const general = CONFIG.banks.targetBand[band] || [];
  const area = rot && rot[i] ? [] : (band === "low" ? [] : CONFIG.banks.target[i] || []);
  return [...new Set([...mine, ...general, ...area])].slice(0, 10);
}

// ---------- Writing check: full sentences, capital letters and full stops ----------
function writingIssues(text) {
  const t = String(text || "").trim();
  if (!t) return [];
  const issues = [];
  // The first sentence carries on from the starter ("I am proud that I..."), so only later sentences need a capital
  if (t.split(/(?<=[.!?])\s+/).slice(1).some((x) => /^[a-z]/.test(x))) issues.push("start each sentence with a capital letter");
  if (/(^|\s)i(\s|'|$)/.test(t)) issues.push("write “I” as a capital letter");
  if (!/[.!?]["”’)]?$/.test(t)) issues.push("end with a full stop");
  const slang = [];
  t.toLowerCase().replace(/[^a-z'\s]/g, " ").split(/\s+/).forEach((w) => { if (CONFIG.slang[w] && !slang.includes(w)) slang.push(w); });
  slang.slice(0, 3).forEach((w) => issues.push(`use “${CONFIG.slang[w]}” instead of “${w}”`));
  return issues;
}

// ---------- Screen layout ----------
function draw() {
  window.scrollTo(0, 0);
  if ("speechSynthesis" in window) speechSynthesis.cancel();
  while (SKIP[steps[at]] && SKIP[steps[at]]()) at += dir;
  const name = steps[at];
  const screen = name.startsWith("x:") ? extraScreen(name.slice(2)) : SCREENS[name]();
  const top = el("header", { class: "top" },
    el("span", { class: "brand" }, "Project SEARCH"),
    progress(),
    el("button", { class: "speak", id: "listenBtn", type: "button", "aria-pressed": "false", "aria-label": "Read this out loud",
      onclick: () => (document.getElementById("listenBtn").getAttribute("aria-pressed") === "true" ? stopSpeaking() : readScreen()) }, "🔊 ", el("span", {}, "Listen"))
  );
  app.replaceChildren(top, screen);
  const h = screen.querySelector("h1");
  if (h) { h.tabIndex = -1; h.focus({ preventScroll: true }); }
}

function progress() {
  const box = el("div", { class: "progress", "aria-hidden": "true" });
  const total = steps.length - 2;
  if (at === 0 || at > total) return box;
  const bar = el("i");
  bar.style.width = Math.round((at / total) * 100) + "%";
  box.append(bar);
  return box;
}

function nav({ onNext, nextLabel = "Next", back = true }) {
  const isLast = steps[at + 1] === "thanks";
  return el("div", { class: "nav" },
    back ? el("button", { class: "btn back", type: "button", onclick: () => { go(-1); } }, "Back") : null,
    el("button", { class: "btn main", type: "button", id: "nextBtn", onclick: async () => {
      if (onNext && !(await onNext())) return;
      if (isLast && !(await send())) return;
      go(1);
    } }, isLast ? "Send" : nextLabel)
  );
}

// ---------- Text question block, shaped by version ----------
// opts: key, starter, bank, prompts, goal, required
function textBlock(opts) {
  const box = el("div", { class: "textblock" });
  const area = el("textarea", { id: opts.key, rows: opts.goal ? 6 : 3, maxlength: 1500 });
  area.value = answers[opts.key] || "";

  if (opts.prompts) {
    box.append(el("div", { class: "prompts" }, el("b", {}, "Think about:"),
      el("ul", {}, opts.prompts.map((p) => el("li", {}, p)))));
  }
  if (opts.bank) {
    const chips = el("div", { class: "chips" });
    opts.bank.forEach((idea) => {
      const used = () => area.value.includes(idea);
      const c = el("button", { class: "chip", type: "button", "aria-pressed": String(used()), onclick: () => {
        let parts = area.value.split(". ").map((p) => p.replace(/\.$/, "").trim()).filter(Boolean);
        parts = used() ? parts.filter((p) => p !== idea) : [...parts, idea];
        area.value = parts.length ? parts.join(". ") + "." : "";
        area.dispatchEvent(new Event("input"));
      } }, idea);
      area.addEventListener("input", () => c.setAttribute("aria-pressed", String(used())));
      chips.append(c);
    });
    box.append(chips);
  }
  box.append(el("label", { class: "starter", for: opts.key }, opts.starter), area);
  if (version !== "supported") box.append(el("p", { class: "tip" }, "✏️ " + CONFIG.writingTip));

  const hint = el("p", { class: "hint", hidden: true });
  let counter = null;
  if (opts.goal) {
    counter = el("div", { class: "counter" }, el("div", { class: "meter" }, el("i")), el("span"));
    box.append(counter);
  }
  box.append(hint);

  const update = () => {
    answers[opts.key] = area.value.trim();
    if (area.value.trim()) hint.hidden = true;
    if (counter) {
      const n = words(area.value);
      counter.querySelector("i").style.width = Math.min(100, (n / opts.goal) * 100) + "%";
      counter.querySelector("span").textContent = `${n} of ${opts.goal} words`;
      counter.classList.toggle("met", n >= opts.goal);
    }
  };
  area.addEventListener("input", update);
  update();

  // Returns true when it is fine to move on
  box.check = () => {
    const n = words(area.value);
    if (opts.required && n === 0) {
      hint.textContent = opts.bank ? "Tap an idea or type something." : "Write your answer first.";
      hint.classList.remove("nudge"); hint.hidden = false; area.focus(); return false;
    }
    // Writing check. Level 3 fixes it before moving on; Levels 1 and 2 can carry on.
    const issues = writingIssues(area.value);
    if (issues.length && !(version !== "independent" && nudged["w_" + opts.key])) {
      const list = issues.join(", ");
      hint.replaceChildren(...[
        el("span", {}, version === "independent"
          ? `✏️ Please fix this before you carry on: ${list}.`
          : `✏️ Check your writing: ${list}.`),
        version === "independent" ? null : el("span", { class: "hint-btns" },
          el("button", { class: "btn back small", type: "button", onclick: () => area.focus() }, "Fix it"),
          el("button", { class: "btn main small", type: "button", onclick: () => { nudged["w_" + opts.key] = true; document.getElementById("nextBtn").click(); } }, "Carry on"))
      ].filter(Boolean));
      hint.classList.add("nudge"); hint.hidden = false; area.focus(); return false;
    }
    if (opts.min && n > 0 && n < opts.min) {
      hint.textContent = `Please write at least ${opts.min} words. You have ${n}. Tell us a little more.`;
      hint.classList.add("nudge"); hint.hidden = false; area.focus(); return false;
    }
    if (opts.goal && n < opts.goal && n > 0 && !nudged[opts.key]) {
      nudged[opts.key] = true;
      hint.textContent = `You have ${n} words. Your goal is ${opts.goal}. Can you add a bit more? Tap Next again to carry on.`;
      hint.classList.add("nudge"); hint.hidden = false; area.focus(); return false;
    }
    return true;
  };
  return box;
}

function questionScreen(title, lead, blocks, navOpts = {}) {
  return el("section", { class: "screen" },
    el("h1", {}, title),
    lead ? el("p", { class: "lead" }, lead) : null,
    blocks,
    nav({ ...navOpts, onNext: () => blocks.every((b) => !b.check || b.check()) })
  );
}

// ---------- The screens ----------
const goals = () => CONFIG.wordGoals[version] || {};

const SCREENS = {
  hello: () => el("section", { class: "screen" },
    el("div", { class: "hero-face", "aria-hidden": "true" }, "👋"),
    el("h1", {}, `Hello ${firstName()}`),
    el("p", { class: "lead" }, version === "independent"
      ? "Take about 15 minutes to think about your shift. Write in your own words."
      : "Let's think about your shift today."),
    profile.fromCache ? el("p", { class: "hint nudge" }, "No internet right now. You can still reflect. It will send later.") : null,
    nav({ back: false, nextLabel: "Start", onNext: () => { answers.started = new Date().toISOString(); return true; } }),
    PREVIEW ? null : el("button", { class: "notme", type: "button", onclick: forget }, `Not ${firstName()}? Tap here`)
  ),

  rotation: () => {
    const list = profile.rotations.length ? profile.rotations : [profile.rotation].filter(Boolean);
    const box = el("div", { class: "options" });
    list.forEach((r) => box.append(el("button", {
      class: "option", type: "button", "aria-pressed": String(answers.rotation === r),
      onclick: (e) => {
        answers.rotation = r;
        box.querySelectorAll(".option").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
      }
    }, r)));
    const hint = el("p", { class: "hint", hidden: true }, "Tap your rotation.");
    return el("section", { class: "screen" },
      el("h1", {}, `Where did you work ${dayWord()}?`),
      el("p", { class: "lead" }, `${profile.site}. Tap your rotation.`),
      box, hint,
      nav({ onNext: () => { if (!answers.rotation) { hint.hidden = false; return false; } return true; } })
    );
  },

  rating: () => el("section", { class: "screen" },
    el("h1", {}, dayWord() === "today" ? "How was your shift?" : `How was your shift ${dayWord()}?`),
    el("p", { class: "lead" }, "Tap one."),
    el("div", { class: "faces" },
      [["Good", "🙂", "I enjoyed it or feel proud", "good"],
       ["Okay", "😐", "Some things went well", "okay"],
       ["Difficult", "🙁", "Some things felt hard", "hard"]].map(([v, e, s, c]) =>
        el("button", { class: `face ${c}`, type: "button", "aria-pressed": String(answers.rating === v),
          onclick: () => { answers.rating = v; go(1); } },
          el("span", { class: "emoji", "aria-hidden": "true" }, e), el("b", {}, v), el("small", {}, s)))
    ),
    el("div", { class: "nav" }, el("button", { class: "btn back", type: "button", onclick: () => { go(-1); } }, "Back"))
  ),

  day: () => {
    const choices = dayChoices();
    const box = el("div", { class: "options days" }, choices.map((c) => el("button", {
      class: "option day", type: "button", disabled: c.done ? "disabled" : null,
      "aria-pressed": String(answers.day === c.key && !c.done),
      onclick: (e) => {
        answers.day = c.key;
        box.querySelectorAll(".option").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
      }
    }, el("b", {}, c.label), el("span", {}, c.when), c.done ? el("small", {}, "✓ Done") : null)));
    return el("section", { class: "screen" },
      el("h1", {}, "Which shift are you reflecting on?"),
      el("p", { class: "lead" }, "Tap one. You can do one reflection for each shift."),
      box,
      nav({ onNext: () => !!answers.day && !choices.find((c) => c.key === answers.day).done })
    );
  },

  alldone: () => el("section", { class: "screen" },
    el("div", { class: "hero-face", "aria-hidden": "true" }, "🌟"),
    el("h1", {}, `Well done, ${firstName()}`),
    el("p", { class: "lead" }, "You have done your reflections for today and your last shift. Come back after your next shift."),
    PREVIEW ? null : el("button", { class: "notme", type: "button", onclick: forget }, `Not ${firstName()}? Tap here`)
  ),

  safe: () => el("section", { class: "screen" },
    el("h1", {}, `Did you feel safe at work ${dayWord()}?`),
    el("p", { class: "lead" }, "Tap one. It is always okay to tell us."),
    el("div", { class: "faces" },
      [["Yes", "🙂", "I felt safe", "good"], ["Not sure", "😐", "I am not sure", "okay"], ["No", "🙁", "I did not feel safe", "hard"]].map(([v, e, t, c]) =>
        el("button", { class: `face ${c}`, type: "button", "aria-pressed": String(answers.safe === v),
          onclick: () => { answers.safe = v; go(1); } },
          el("span", { class: "emoji", "aria-hidden": "true" }, e), el("b", {}, v), el("small", {}, t)))),
    el("div", { class: "nav" }, el("button", { class: "btn back", type: "button", onclick: () => go(-1) }, "Back"))
  ),

  safe2: () => {
    const hint = el("p", { class: "hint", hidden: true }, "Tap one.");
    const box = el("div", { class: "options" }, [["Yes", "Yes, please talk to me"], ["No", "No, thank you"]].map(([v, t]) =>
      el("button", { class: "option", type: "button", "aria-pressed": String(answers.wantsTalk === v), onclick: (e) => {
        answers.wantsTalk = v; hint.hidden = true;
        box.querySelectorAll(".option").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
      } }, t)));
    const note = el("textarea", { id: "safetyNote", rows: 3, maxlength: 800 });
    note.value = answers.safetyNote || "";
    note.addEventListener("input", () => (answers.safetyNote = note.value.trim()));
    return el("section", { class: "screen" },
      el("h1", {}, "Would you like someone from the Project SEARCH team to talk to you?"),
      el("p", { class: "lead" }, "Thank you for telling us. Tap one."),
      box, hint,
      el("label", { class: "starter", for: "safetyNote" }, "You can tell us more (you can leave this empty)"),
      note,
      nav({ onNext: () => { if (!answers.wantsTalk) { hint.hidden = false; return false; } return true; } })
    );
  },

  s0: () => startScreen(0), s1: () => startScreen(1), s2: () => startScreen(2), s3: () => startScreen(3),

  t0: () => targetScreen(0), t1: () => targetScreen(1), t2: () => targetScreen(2), t3: () => targetScreen(3),
  c0: () => checkInScreen(0), c1: () => checkInScreen(1), c2: () => checkInScreen(2), c3: () => checkInScreen(3),
  ci_intro: () => {
    const n = profile.targets.filter(Boolean).length;
    return el("section", { class: "screen" },
      el("div", { class: "hero-face", "aria-hidden": "true" }, "🗓️"),
      el("h1", {}, "End of half term"),
      el("p", { class: "lead" }, `${n} quick question${n === 1 ? "" : "s"} about your targets. Think about the whole half term, not just today.`),
      nav({})
    );
  },

  well: () => {
    if (version === "supported") return questionScreen("Something I did well", "Tap the ideas that are true.", [
      textBlock({ key: "wellDone", starter: "I am proud that I...", bank: CONFIG.banks.well, required: true, min: CONFIG.preSipMinWords })]);
    if (version === "middle") return questionScreen("Something I did well", "Tap an idea or type your own.", [
      textBlock({ key: "wellDone", starter: "I am proud that I...", bank: CONFIG.banks.well.slice(0, 6), required: true }),
      textBlock({ key: "wellMore", starter: "Tell me more. How did you do it?", required: true, goal: goals().follow })]);
    return questionScreen("Something I did well", null, [
      textBlock({ key: "wellDone", starter: "I am proud that I...", required: true, goal: goals().well,
        prompts: ["What did you do?", "How did you do it?", "Why are you proud of it?"] })]);
  },

  hard: () => {
    if (version === "supported") return questionScreen("Something I found hard", "Tap any that are true. You can skip this.", [
      textBlock({ key: "difficult", starter: "I found it hard when...", bank: CONFIG.banks.hard })]);
    if (version === "middle") return questionScreen("Something I found hard", "You can skip this if nothing was hard.", [
      textBlock({ key: "difficult", starter: "I found it hard when...", bank: CONFIG.banks.hard.slice(0, 6) }),
      textBlock({ key: "difficultMore", starter: "What did you do when that happened?", goal: goals().follow })]);
    return questionScreen("Something I found hard", "If nothing was hard, write about something that could have gone better.", [
      textBlock({ key: "difficult", starter: "I found it hard when...", goal: goals().hard,
        prompts: ["What happened?", "What did you do about it?", "What would you do differently?"] })]);
  },

  help: () => {
    const opts = [
      ["Clear instructions", "📋", "Clear instructions"], ["A visual checklist", "✅", "A picture checklist"],
      ["More time to practise", "⏱️", "More time to practise"], ["Help from my supervisor", "🙋", "Help from my supervisor"],
      ["A quieter environment", "🤫", "A quieter place"], ["A warning before changes", "🔔", "A warning before changes"],
      ["A demonstration", "👀", "Show me how"]];
    answers.help = answers.help || [];
    const tiles = el("div", { class: "tiles" }, opts.map(([v, e, t]) =>
      el("button", { class: "tile", type: "button", "aria-pressed": String(answers.help.includes(v)), onclick: (ev) => {
        answers.help = answers.help.includes(v) ? answers.help.filter((x) => x !== v) : [...answers.help, v];
        ev.currentTarget.setAttribute("aria-pressed", String(answers.help.includes(v)));
      } }, el("span", { "aria-hidden": "true" }, e), t)));
    const other = el("input", { id: "helpOther", type: "text", autocomplete: "off", maxlength: 150 });
    other.value = answers.helpOther || "";
    other.addEventListener("input", () => (answers.helpOther = other.value.trim()));
    return el("section", { class: "screen" },
      el("h1", {}, "What could help me next time?"),
      el("p", { class: "lead" }, "Tap all that help."),
      tiles,
      el("label", { class: "starter", for: "helpOther" }, "Something else (you can leave this empty)"),
      other, nav({})
    );
  },

  next: () => {
    if (version === "supported") return questionScreen("What will I try next time?", "Tap one idea.", [
      textBlock({ key: "nextTime", starter: "Next shift, I will...", bank: CONFIG.banks.next, required: true, min: CONFIG.preSipMinWords })]);
    if (version === "middle") return questionScreen("What will I try next time?", "Tap an idea or type your own.", [
      textBlock({ key: "nextTime", starter: "Next shift, I will...", bank: CONFIG.banks.next.slice(0, 6), required: true }),
      textBlock({ key: "nextHow", starter: "How will you do that?", required: true, goal: goals().follow })]);
    return questionScreen("What will I try next time?", null, [
      textBlock({ key: "nextTime", starter: "Next shift, I will...", required: true, goal: goals().next,
        prompts: ["What will you do differently?", "How will you remember to do it?", "How will you know it worked?"] })]);
  },

  learned: () => questionScreen("What did I learn today?", null, [
    textBlock({ key: "learned", starter: "Today I learned...", required: true, goal: goals().learned,
      prompts: ["A new skill or fact", "Something about how you work best", "How you could use it again"] })]),

  thanks: () => el("section", { class: "screen" },
    el("div", { class: "hero-face", "aria-hidden": "true" }, "⭐"),
    el("h1", {}, `Thank you, ${firstName()}`),
    el("p", { class: "lead", id: "saveMsg" }, answers.saveMsg || "Your reflection is sent."),
    answers.safe && answers.safe !== "Yes" ? el("div", { class: "remember" }, el("b", {}, "Thank you for telling us. "),
      answers.wantsTalk === "Yes" ? "Someone from the Project SEARCH team will talk to you soon." : "If you change your mind, talk to your job coach or site lead.") : null,
    el("div", { class: "remember" }, el("b", {}, "Remember: "),
      "it is okay to ask for support. You can talk to your job coach, supervisor or workplace buddy."),
    PREVIEW ? recordPanel() : null,
    el("div", { class: "nav" }, el("button", { class: "btn main", type: "button", onclick: () => start(profile.token) }, "Finish"))
  )
};

function extraScreen(id) {
  const q = profile.questions.find((x) => x.id === id);
  const key = "x_" + id;
  if (q.type === "text") {
    return questionScreen(q.text, null, [textBlock({ key, starter: "My answer",
      bank: version === "independent" ? null : q.options, required: !!q.required,
      goal: version === "independent" ? q.wordGoal : 0 })]);
  }
  const hint = el("p", { class: "hint", hidden: true }, "Tap one.");
  const choices = q.type === "scale" ? CONFIG.scale.map((s) => [s.n, s]) : (q.options || []).map((o) => [o, null]);
  const box = el("div", { class: q.type === "scale" ? "scale" : "options" }, choices.map(([v, s]) =>
    el("button", { class: s ? `scale-btn s${s.n}` : "option", type: "button", "aria-pressed": String(answers[key] === v),
      onclick: (e) => {
        answers[key] = v; hint.hidden = true;
        box.querySelectorAll("button").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
      } }, s ? [el("b", {}, String(s.n)), el("span", {}, s.says)] : v)));
  return el("section", { class: "screen" },
    el("h1", {}, q.text),
    el("p", { class: "lead" }, "Tap one."),
    box, hint,
    nav({ onNext: () => { if (q.required !== false && answers[key] == null) { hint.hidden = false; return false; } return true; } })
  );
}

// Same 1 to 5 scale as the team's half-term review, in simple words
const CHECKIN_CHOICES = [[5, "Every time"], [4, "Most times"], [3, "Sometimes"], [2, "Once or twice"], [1, "Not yet"]];
function checkInScreen(i) {
  const key = "ci" + (i + 1);
  const shown = profile.targets.filter(Boolean).length;
  const num = profile.targets.slice(0, i + 1).filter(Boolean).length;
  const hint = el("p", { class: "hint", hidden: true }, "Tap one.");
  const box = el("div", { class: "options" }, CHECKIN_CHOICES.map(([v, label]) =>
    el("button", { class: "option", type: "button", "aria-pressed": String(answers[key] === v), onclick: (e) => {
      answers[key] = v; hint.hidden = true;
      box.querySelectorAll(".option").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
    } }, label)));
  return el("section", { class: "screen" },
    el("p", { class: "eyebrow" }, `Half-term check-in, ${num} of ${shown}`),
    el("p", { class: "area" }, el("span", { "aria-hidden": "true" }, CONFIG.areas[i].icon), CONFIG.areas[i].name),
    el("div", { class: "target-text" }, profile.targets[i]),
    el("h1", {}, "Over this half term, how often did I do this?"),
    box, hint,
    nav({ onNext: () => { if (!answers[key]) { hint.hidden = false; return false; } return true; } })
  );
}

// The intern's own starting point for a new target, in the same words as the check-in
function startScreen(i) {
  const key = "start" + (i + 1);
  const hint = el("p", { class: "hint", hidden: true }, "Tap one.");
  const box = el("div", { class: "options" }, CHECKIN_CHOICES.map(([v, label]) =>
    el("button", { class: "option", type: "button", "aria-pressed": String(answers[key] === v), onclick: (e) => {
      answers[key] = v; hint.hidden = true;
      box.querySelectorAll(".option").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
    } }, label)));
  return el("section", { class: "screen" },
    el("p", { class: "eyebrow" }, "New target"),
    el("p", { class: "area" }, el("span", { "aria-hidden": "true" }, CONFIG.areas[i].icon), CONFIG.areas[i].name),
    el("div", { class: "target-text" }, profile.targets[i]),
    el("h1", {}, "Before we start: how often can I do this now?"),
    el("p", { class: "lead" }, "Be honest. This is your starting point."),
    box, hint,
    nav({ onNext: () => { if (!answers[key]) { hint.hidden = false; return false; } return true; } })
  );
}

function targetScreen(i) {
  const key = "t" + (i + 1);
  const scale = el("div", { class: "scale" }, CONFIG.scale.map((s) =>
    el("button", { class: `scale-btn s${s.n}`, type: "button", "aria-pressed": String(answers[key + "Score"] === s.n),
      onclick: (e) => {
        answers[key + "Score"] = s.n;
        showNote(s.n);
        scale.querySelectorAll(".scale-btn").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
        hint.hidden = true;
      } },
      el("b", {}, String(s.n)), el("span", {}, s.says))));
  const hint = el("p", { class: "hint", hidden: true }, "Tap a number.");
  // The "What did I do?" box changes with the score the intern gives: struggle words for 1 to 2, some help for 3, doing well for 4 to 5
  const holder = el("div", { class: "note-holder" });
  const noteBox = { current: null, check: () => (noteBox.current ? noteBox.current.check() : true) };
  const showNote = (score) => {
    const band = bandOf(score), pre = version === "supported";
    const when = dayWord() === "today" ? "today" : "that shift";
    const q = CONFIG.targetNotes[pre ? "pre" : "sip"][band];
    answers[key + "Starter"] = pre ? "" : q.starter;
    const opts = pre
      ? { key: key + "Note", starter: q.ask, bank: targetBank(i, band), required: true, min: CONFIG.preSipMinWords }
      : { key: key + "Note", starter: q.starter, required: true, goal: goals().target, prompts: q.prompts.map((t) => t.replace("{when}", when)) };
    noteBox.current = textBlock(opts);
    holder.replaceChildren(noteBox.current);
  };
  if (answers[key + "Score"]) showNote(answers[key + "Score"]);
  const blocks = [noteBox];
  return el("section", { class: "screen" },
    el("p", { class: "eyebrow" }, `My target ${profile.targets.slice(0, i + 1).filter(Boolean).length} of ${profile.targets.filter(Boolean).length}`),
    el("p", { class: "area" }, el("span", { "aria-hidden": "true" }, CONFIG.areas[i].icon), CONFIG.areas[i].name),
    el("div", { class: "target-text" }, profile.targets[i] || "Target to be added"),
    el("h1", {}, `How did I do ${dayWord()}?`),
    scale, hint, holder,
    nav({ onNext: () => {
      if (!answers[key + "Score"]) { hint.hidden = false; return false; }
      return noteBox.check();
    } })
  );
}

// ---------- Sending to Excel ----------
// Excel treats text that starts with = + - @ as a formula. A leading space keeps it as plain text.
const safe = (v) => {
  const t = String(v ?? "").replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").trim().slice(0, 2000);
  return /^[=+\-@]/.test(t) ? " " + t : t;
};
// Saves answers as full sentences, for example "Next shift, I will... check my list."
const withStarter = (starter, text) => (text ? `${starter} ${safe(text)}` : "");
// Each reflection gets its own ID, so a resend after lost signal is never saved twice
function newId() {
  const d = new Date(), p = (n) => String(n).padStart(2, "0");
  const rand = Math.random().toString(36).slice(2, 6).toUpperCase().replace(/[^A-Z0-9]/g, "X").padEnd(4, "X");
  return `R${String(d.getFullYear()).slice(2)}${p(d.getMonth() + 1)}${p(d.getDate())}${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}${rand}`;
}

function buildRecord() {
  const finished = new Date();
  const help = [...(answers.help || [])];
  if (answers.helpOther) help.push(answers.helpOther);
  const textKeys = ["wellDone", "wellMore", "difficult", "difficultMore", "nextTime", "nextHow", "learned", "t1Note", "t2Note", "t3Note", "t4Note"];
  return {
    // One reflection per intern per shift day: the flow keeps the first row with this ID
    id: profile.internId && answers.day ? profile.internId + "-" + answers.day : newId(),
    token: profile.token,
    internId: profile.internId || "",
    appVersion: CONFIG.version,
    started: answers.started,
    finished: finished.toISOString(),
    minutes: Math.max(1, Math.round((finished - new Date(answers.started)) / 60000)),
    name: profile.name,
    site: profile.site,
    level: levelName(profile.level),
    appMode: version,
    rotation: answers.rotation,
    rating: answers.rating,
    shiftDate: answers.day,
    safe: answers.safe || "", wantsTalk: answers.safe && answers.safe !== "Yes" ? (answers.wantsTalk || "") : "",
    safetyNote: answers.safe && answers.safe !== "Yes" ? safe(answers.safetyNote) : "",
    start1: answers.start1 || "", start2: answers.start2 || "", start3: answers.start3 || "", start4: answers.start4 || "",
    target1: profile.targets[0] || "", t1Score: answers.t1Score || "", t1Note: withStarter(answers.t1Starter, answers.t1Note) || safe(answers.t1Note),
    target2: profile.targets[1] || "", t2Score: answers.t2Score || "", t2Note: withStarter(answers.t2Starter, answers.t2Note) || safe(answers.t2Note),
    target3: profile.targets[2] || "", t3Score: answers.t3Score || "", t3Note: withStarter(answers.t3Starter, answers.t3Note) || safe(answers.t3Note),
    target4: profile.targets[3] || "", t4Score: answers.t4Score || "", t4Note: withStarter(answers.t4Starter, answers.t4Note) || safe(answers.t4Note),
    checkIn1: answers.ci1 || "", checkIn2: answers.ci2 || "", checkIn3: answers.ci3 || "", checkIn4: answers.ci4 || "",
    wellDone: withStarter("I am proud that I...", answers.wellDone), wellMore: safe(answers.wellMore),
    difficult: withStarter("I found it hard when...", answers.difficult), difficultMore: safe(answers.difficultMore),
    help: help.map(safe).join(";"),
    nextTime: withStarter("Next shift, I will...", answers.nextTime), nextHow: safe(answers.nextHow),
    learned: safe(answers.learned),
    wordCount: textKeys.reduce((n, k) => n + words(answers[k]), 0),
    extraAnswers: profile.questions.filter((q) => answers["x_" + q.id] != null)
      .map((q) => ({ id: q.id, question: safe(q.text), answer: typeof answers["x_" + q.id] === "number" ? answers["x_" + q.id] : safe(answers["x_" + q.id]) }))
  };
}

const QUEUE = "reflectionQueue";
const readQ = () => { try { return JSON.parse(localStorage.getItem(QUEUE)) || []; } catch (e) { return []; } };
const writeQ = (q) => { try { localStorage.setItem(QUEUE, JSON.stringify(q)); } catch (e) { /* blocked */ } };

async function post(record) {
  const r = await fetchWithTimeout(CONFIG.submitUrl, {
    method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(record)
  }, CONFIG.saveTimeoutMs || 45000);
  // 403 means the card code is no longer recognised, usually because a lost card was replaced
  if (r.status === 403) return "rejected";
  if (!r.ok) throw new Error("save failed " + r.status);
  return "saved";
}
function queue(record) {
  const q = readQ().filter((x) => x.id !== record.id);
  writeQ([...q, record].slice(-50));
}

async function send() {
  const b = document.getElementById("nextBtn");
  b.disabled = true; b.textContent = "Sending...";
  const record = buildRecord();
  answers.record = record;
  markCheckInDone();
  markDone(answers.day);
  markSeen();
  if (PREVIEW) {
    answers.saveMsg = "Preview mode, so nothing was saved.";
  } else {
    try {
      const res = await post(record);
      answers.saveMsg = res === "saved" ? "Your reflection is sent." : "This card is not active. Please ask your job coach.";
    } catch (e) {
      queue(record);
      answers.saveMsg = "No internet right now. Your reflection is saved on this phone and will send next time you open the app.";
    }
  }
  return true;
}

let sending = false, again = false;
const DAY = 86400000;
async function sendWaiting() {
  if (PREVIEW) return;
  if (sending) { again = true; return; }
  sending = true;
  const done = new Set();
  for (const r of readQ()) {
    try {
      const res = await post(r);
      // A refused reflection is kept for 30 days, in case this phone scans the intern's new card
      const old = Date.now() - new Date(r.finished || 0).getTime() > 30 * DAY;
      if (res === "saved" || !r.internId || old) done.add(r.id);
    } catch (e) { /* stays queued for next time */ }
  }
  writeQ(readQ().filter((x) => !done.has(x.id)));
  sending = false;
  if (again) { again = false; sendWaiting(); }
}
// After a lost card is replaced: reflections waiting on this phone for the same intern move to the new card
function moveWaitingToCard(p) {
  if (!p.internId) return;
  const q = readQ();
  let moved = 0;
  const next = q.map((x) => (x.internId === p.internId && x.token !== p.token ? (moved++, { ...x, token: p.token }) : x));
  if (moved) { writeQ(next); sendWaiting(); }
}

// Preview only: shows the row Excel would receive
function recordPanel() {
  const rec = answers.record || {};
  const rows = Object.entries(rec).map(([k, v]) => el("tr", {}, el("th", {}, k), el("td", {}, typeof v === "object" ? JSON.stringify(v) : String(v))));
  return el("details", { class: "record" },
    el("summary", {}, "Preview: what Excel would receive"),
    el("div", { class: "table-wrap" }, el("table", {}, el("tbody", {}, rows))));
}

// ---------- Preview switcher ----------
function previewBar() {
  if (!PREVIEW) return;
  const labels = { "demo-a7k2": "Pre-SIP", "demo-x3r8": "SIP" };
  const bar = el("div", { class: "preview-bar" },
    el("span", {}, "Preview as:"),
    Object.entries(labels).map(([t, l]) => el("button", { type: "button", "data-t": t, onclick: (e) => {
      bar.querySelectorAll("button[data-t]").forEach((b) => b.setAttribute("aria-pressed", String(b === e.currentTarget)));
      start(t);
    } }, l)));
  bar.append(el("button", { type: "button", class: "reset", onclick: () => { MEMORY.clear(); start(profile ? profile.token : token); } }, "Reset trial"));
  document.body.prepend(bar);
  bar.querySelector("button").setAttribute("aria-pressed", "true");
}

// ---------- Start up ----------
previewBar();
// The card code stays in the address, so a home screen shortcut made from it keeps working
const token = tokenFromLink() || (PREVIEW ? "demo-a7k2" : remembered());
start(token);
sendWaiting();
window.addEventListener("online", sendWaiting);
