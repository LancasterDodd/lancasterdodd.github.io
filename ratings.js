/* ratings.js  --  shared by index.html (the ratings) and lab.html (the Lab).
   It holds: the settings (names, file), reading the Excel file, small helpers, the fun remark lines and album covers.
   Edit the names, REMARKS and PERSONAL here; both pages pick up the change. */

/* ---------- SETTINGS ---------- */
const DATA_FILE = "albums.xlsx";   // the Excel file sitting next to this page
const NAME_1 = "Nimbus";           // the names SHOWN on the website
const NAME_2 = "Korin";
const SHEET_NAME_1 = "Nimbus";      // FALLBACK only: the page reads the names straight from the Excel headers ("O.R. <name>"): first name = NAME_1, second = NAME_2. The names as written in the Excel headers ("O.R. Nigel", "Top Song Nigel"...) and in "Picked By".
const SHEET_NAME_2 = "Korin";      // (If the Excel says Nimbus/Korin, put Nimbus/Korin here.) The site only shows NAME_1 / NAME_2.
const LOOK_UP_COVERS = true;       // false = never contact Apple, show colored placeholders only
const CRITIC_SOURCE = "Pitchfork";  // name shown in the Lab for the optional "Critic Score" column (e.g. "Pitchfork", "Metacritic", "the critics")

/* The fun lines under an album when you two disagree (or agree). Add, remove or rewrite as many as you like!
   {hi} = who rated higher, {lo} = who rated lower, {d} = points apart, {hs}/{ls} = the high/low score,
   {hg}/{lg} = the high/low grade name (Jimmey, S, A... Bean), {picker} = who picked the album (only if one of you did) */
const REMARKS = {
  perfect: ["Double Jimmey. Nothing more to say.", "A perfect 10 from both of us. Frame it."],
  bean:    ["Bean for both of us. We survived.", "We agree: never again."],
  same:    ["For once, we actually agree.", "Same score. No argument tonight."],
  small:   ["{hi} liked it a bit more.", "{lo} is holding back. {hi} is not."],
  medium:  ["{hi} is vibing, {lo} is not quite there.", "We heard two different albums. {hi} wins by {d}.", "{hi} says {hg}, {lo} says {lg}."],
  big:     ["Argument starter: {hi} gave it a {hs}, {lo} gave it a {ls}.", "{lo} and {hi} will not be settling this one.", "{hi} says {hg}. {lo} says {lg}. Same album.", "{d} points apart. We need a tiebreaker."],
  // extra lines that only apply when {picker} picked the album and the two of us disagree by 1 point or more:
  pickerHigh: ["{picker} picked it and defends it. {lo} is not buying it.", "Good pick by {picker}, according to {picker}."],
  pickerLow:  ["{picker} picked it and still rated it lower than {hi} did.", "Own goal by {picker}: {hi} liked it more."],
};
/* Lines that belong to one person. They are added when that person rated the album higher ("high") or lower ("low").
   This is the place for inside jokes! Keys must be exactly the names in NAME_1 and NAME_2. */
const PERSONAL = {
  Nimbus: { high: ["Nimbus is on a victory lap."], low: ["Nimbus is not convinced."] },
  Korin:  { high: ["Korin is on a victory lap."],  low: ["Korin is not convinced."] },
};

/* ---------- 5. READ THE EXCEL FILE ---------- */
let albums = [];      // rated albums
let listening = [];   // albums in the sheet that have a title but no rating yet = what you are listening to now

async function loadAlbums() {
  const res = await fetch(DATA_FILE + "?v=" + Date.now());   // ?v= avoids an old cached copy
  if (!res.ok) throw new Error("Could not load " + DATA_FILE + " (" + res.status + ")");
  const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];          // first sheet in the workbook
  const range = XLSX.utils.decode_range(ws["!ref"]);

  // find the header row (the row that contains "Album")
  let headerRow = -1;
  for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 10) && headerRow < 0; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      if (text(ws[XLSX.utils.encode_cell({ r, c })]).toLowerCase() === "album") { headerRow = r; break; }
    }
  }
  if (headerRow < 0) throw new Error('No header row with an "Album" column found.');

  // map column name -> column number, so the order of columns doesn't matter
  const col = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const hd = text(ws[XLSX.utils.encode_cell({ r: headerRow, c })]).toLowerCase();
    if (hd) col[hd] = c;
  }
  const has = name => col[name.toLowerCase()] !== undefined;
  const cell = (r, name) => has(name) ? ws[XLSX.utils.encode_cell({ r, c: col[name.toLowerCase()] })] : undefined;
  const sheetNames = [];                           // read the two names from the headers "O.R. <name>" (so the Excel may say Nimbus/Korin or real names)
  for (const k of Object.keys(col)) { const m = k.match(/^o\.r\.\s+(.+?)(\s+j-be)?$/); if (m && !sheetNames.includes(m[1])) sheetNames.push(m[1]); }
  const S1 = sheetNames[0] || SHEET_NAME_1, S2 = sheetNames[1] || SHEET_NAME_2;
  const isFav = (r, name) => fold(text(cell(r, "O.R. " + name + " J-Be"))) === "fav" || fold(text(cell(r, "O.R. " + name))) === "fav";   // FAV typed in the letter column (or straight in the rating column)

  const out = [];
  for (let r = headerRow + 1; r <= range.e.r; r++) {
    const title = text(cell(r, "Album"));
    if (!title) continue;                          // skip empty rows
    const f1 = isFav(r, S1), f2 = isFav(r, S2);
    const n1 = score(cell(r, "O.R. " + S1)) ?? (f1 ? 10 : null), n2 = score(cell(r, "O.R. " + S2)) ?? (f2 ? 10 : null);   // FAV without a number = 10
    const a = {
      row: r,
      pending: n1 == null && n2 == null,           // no overall rating yet
      title, artist: text(cell(r, "Artist")),
      year: num(cell(r, "Year")),
      date: excelDate(cell(r, "Dates")),
      by: text(cell(r, "Picked By")).replace(new RegExp(S1.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), NAME_1).replace(new RegExp(S2.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), NAME_2),   // real names in the sheet -> shown names
      comment: text(cell(r, "Comments")),
      coverUrl: text(cell(r, "Cover URL")),        // optional column: your own image link or file path
      spotify: text(cell(r, "Spotify URL")),       // optional column: a direct Spotify link (otherwise we link to a search)
      genres: text(cell(r, "Genre")).split(/\s*[,;/]\s*/).filter(Boolean),   // optional column: "Hip-hop, R&B"
      fav: [f1, f2],   // is it in Nimbus's / Korin's top 4?
      r1: n1, r2: n2,
      avg: n1 != null && n2 != null ? (n1 + n2) / 2 : (n1 ?? n2),
      diff: n1 != null && n2 != null ? Math.abs(n1 - n2) : null,
      song:  [text(cell(r, "Top Song " + S1)), text(cell(r, "Top Song " + S2))],
      feat:  [text(cell(r, "Best Feature " + S1)), text(cell(r, "Best Feature " + S2))],
      instr: [text(cell(r, "Best Instrumentals " + S1)), text(cell(r, "Best Instrumentals " + S2))],
      cover: [score(cell(r, "Cover Rating " + S1)), score(cell(r, "Cover Rating " + S2))],
      critic: criticScore(cell(r, "Critic Score") || cell(r, CRITIC_SOURCE)),     // optional column "Critic Score" (0-10, or 0-100)
      ...chefFields(text(cell(r, "Chef's Choice") || cell(r, "Rec"))),            // optional column "Chef's Choice", e.g. "[Debut, Production] Why it deserves more ears"
    };
    a.hayMain  = fold(a.title + " " + a.artist);                                    // what the search box looks at
    a.hayExtra = fold([...a.song, ...a.feat, ...a.instr, a.comment].join(" "));     // only if the checkbox is ticked
    out.push(a);
  }
  return out;
}

/* helpers that turn messy spreadsheet cells into clean values */
function text(c) { return c ? String(c.w ?? c.v ?? "").trim() : ""; }       // .w = text as shown in Excel
function num(c)  { return c && typeof c.v === "number" ? c.v : null; }
function score(c) {                                  // 9.6, "8,0/10", "Jimmey" -> number
  if (!c || c.v == null) return null;
  if (typeof c.v === "number") return c.v;
  const s = String(c.v).trim();
  if (s.toLowerCase() === "jimmey" || s.toLowerCase() === "fav") return 10;      // FAV counts as a 10
  const m = s.replace(",", ".").match(/^(\d+(?:\.\d+)?)/);
  return m ? parseFloat(m[1]) : null;
}
function criticScore(c) {                            // 7.8 or 78 -> 7.8 (scores above 10 are treated as out of 100)
  const n = score(c);
  return n == null ? null : (n > 10 ? n / 10 : n);
}
function chefFields(raw) {                           // "[Debut, Production] Great drums" -> note + tags
  const m = raw.match(/^\[([^\]]*)\]\s*([\s\S]*)$/);
  return m ? { rec: m[2].trim(), recTags: m[1].split(/\s*[,;\/]\s*/).filter(Boolean), recPick: true } : { rec: raw, recTags: [], recPick: !!raw };
}
function excelDate(c) { return c && typeof c.v === "number" ? Math.round((c.v - 25569) * 86400000) : null; }
function fold(s) { return String(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); }   // "Beyoncé" -> "beyonce"
function grade(n) {                                   // same rules as the Excel formula
  if (n == null) return "";
  if (n === 10) return "Jimmey";
  if (n > 8.9) return "S";
  if (n > 8.4) return "A";
  if (n > 7.9) return "B";
  if (n > 6.9) return "C";
  if (n > 5.9) return "D";
  return "Bean";
}
function hash(s) { let x = 0; for (const ch of s) x = (x * 31 + ch.codePointAt(0)) >>> 0; return x; }

/* ---------- 6. SMALL BUILDING BLOCKS ---------- */
function h(tag, cls, content) {                       // tiny helper to build elements safely
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (content != null) e.textContent = content;
  return e;
}
const fmt = n => n == null ? "–" : (Number.isInteger(n) ? n.toFixed(0) : String(+n.toFixed(2)));
const fmtDate = ms => ms == null ? "" : new Date(ms).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric", timeZone: "UTC" });
const blank = s => !s || s === "/" || s === "-";
const pickedByOne = a => [NAME_1, NAME_2].find(n => fold(a.by) === fold(n)) || "";

function spotifyUrl(a) {                              // direct link from the sheet, otherwise a Spotify search for this album
  return a.spotify || "https://open.spotify.com/search/" + encodeURIComponent(a.artist + " " + a.title);
}
function spotifyLink(a) {
  const l = h("a", "spotify", "Listen on Spotify ↗");
  l.href = spotifyUrl(a); l.target = "_blank"; l.rel = "noopener noreferrer";
  l.addEventListener("click", e => e.stopPropagation());
  return l;
}

function remark(a) {                                  // the personalized line
  if (a.r1 == null || a.r2 == null) return "";
  const d = a.diff, hiIsFirst = a.r1 >= a.r2;
  const hi = hiIsFirst ? NAME_1 : NAME_2, lo = hiIsFirst ? NAME_2 : NAME_1, hs = Math.max(a.r1, a.r2), ls = Math.min(a.r1, a.r2);
  const picker = pickedByOne(a);
  const vars = { hi, lo, d: fmt(d), hs: fmt(hs), ls: fmt(ls), hg: grade(hs), lg: grade(ls), picker };
  let pool;
  if (d === 0) pool = a.r1 === 10 ? REMARKS.perfect : a.r1 < 6 ? REMARKS.bean : REMARKS.same;
  else if (d < 1) return "";                          // tiny difference: say nothing
  else {
    pool = d < 2 ? REMARKS.small : d < 4 ? REMARKS.medium : REMARKS.big;
    pool = [...pool];
    if (picker) pool.push(...(picker === hi ? REMARKS.pickerHigh : REMARKS.pickerLow));   // who picked it matters
    pool.push(...(PERSONAL[hi]?.high || []), ...(PERSONAL[lo]?.low || []));              // and lines for each of you
  }
  if (!pool.length) return "";
  return pool[hash(a.title) % pool.length].replace(/\{(\w+)\}/g, (_, k) => vars[k] ?? "");
}

function pairInto(body, a, b) {                       // one value if you agree, two lines if you don't
  if (blank(a) && blank(b)) { body.append(h("span", "none", "–")); return; }
  if (!blank(a) && fold(a) === fold(b)) { body.append(a, h("span", "both", "both")); return; }
  for (const [name, val] of [[NAME_1, a], [NAME_2, b]]) {
    const line = h("div", "line");
    line.append(h("span", "who", name), blank(val) ? "–" : val);
    body.append(line);
  }
}
function tile(kind, label, fill) {
  const t = h("div", "tile k-" + kind);
  const body = h("div", "tile-body");
  fill(body);
  t.append(h("div", "tag", label), body);
  return t;
}
function scoreBox(name, n, fav) {
  const g = grade(n);
  const box = h("div", "score" + (g ? " g-" + g : ""));
  box.append(h("div", "who", name), h("div", "num", fmt(n)));
  if (g) box.append(h("span", "grade", fav ? "FAV ★" : g));
  return box;
}


/* a name as a comparable key, and a stable link key for one album (used by links like index.html?album=...) */
const nk = s => fold(s).replace(/^the\s+/, "").replace(/[^a-z0-9]+/g, " ").trim();      // a name as a comparable key
const albumKey = a => nk(a.artist) + "|" + nk(a.title);

async function loadAll() {                            // read the sheet and split it into rated albums and "now listening"
  const data = await loadAlbums();
  albums = data.filter(a => !a.pending);
  listening = data.filter(a => a.pending);
  warmFavorites(albums);
  return { albums, listening };
}

/* ---------- 8. ALBUM COVERS (looked up by the visitor's browser, remembered afterwards) ---------- */
const PLACEHOLDER_COLORS = ["#bcc3ff", "#dadeff", "#8f9aff", "#ebedff"];
const coversAllowed = () => !window.PRIVACY || PRIVACY.covers;      // the visitor decides in the privacy pop-up (privacy.js)
let coverCache = {};
const loadCache = () => { try { coverCache = JSON.parse(localStorage.getItem("coverCache.v1") || "{}"); } catch (e) {} };
if (coversAllowed()) loadCache();
const saveCovers = () => { if (!coversAllowed()) return; try { localStorage.setItem("coverCache.v1", JSON.stringify(coverCache)); } catch (e) {} };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const coverKey = a => fold(a.artist + "|" + a.title);
const simple = s => fold(s).replace(/\(.*?\)|\[.*?\]/g, " ").replace(/[^a-z0-9]+/g, " ").trim();

const smallCover = url => url.replace(/\/\d+x\d+bb\.(jpg|png)/, "/300x300bb.$1").replace(/\/front-(500|1200)$/, "/front-250");   // phones need far fewer pixels: about 6x lighter than 600px
/* Covers load through a small queue, 6 at a time, always the one highest on the screen first.
   Albums you have already scrolled past are skipped, so fast scrolling doesn't pile up downloads. */
const imgQ = []; let imgActive = 0;
const yOf = it => it.box ? it.box.getBoundingClientRect().top + scrollY : 1e12;      // position on the page; no box = background warm-up, goes last
function pumpImgs() {
  while (imgActive < 6 && imgQ.length) {
    for (let i = imgQ.length - 1; i >= 0; i--) if (imgQ[i].box && !imgQ[i].box.dataset.visible) { delete imgQ[i].box.dataset.requested; imgQ.splice(i, 1); }   // scrolled past: ask again when it comes back
    if (!imgQ.length) break;
    imgQ.forEach(it => { it.y = yOf(it); }); imgQ.sort((x, y) => x.y - y.y);
    const it = imgQ.shift(); imgActive++;
    const img = new Image(); img.decoding = "async";
    const fin = ok => { imgActive--; if (ok && it.box) it.done(img); pumpImgs(); };
    img.onload = () => fin(true); img.onerror = () => fin(false);
    img.src = it.url;
  }
}
function queueImg(url, box, done) { imgQ.push({ url, box, done }); pumpImgs(); }
function showCover(box, url) {
  const al = box._album;
  queueImg(smallCover(url), box, img => { img.alt = al ? "Album cover: " + al.title + " by " + al.artist : ""; box.append(img); });
}
function warmFavorites(albums) {                    // favourites further down are fetched quietly in the background, after what is on screen
  if (!coversAllowed()) return;
  coverMapReady.then(() => setTimeout(() => {          // wait a moment so the covers on screen always get the first slots
    for (const a of albums) if (a.fav && (a.fav[0] || a.fav[1])) {
      const u = a.coverUrl || coverMap[coverKey(a)] || coverCache[coverKey(a)];
      if (u) queueImg(smallCover(u), null, () => {});
    }
  }, 1500));
}
window.warmFavorites = warmFavorites;
async function lookupCover(a) {
  const q = encodeURIComponent(a.artist + " " + a.title);
  const res = await fetch("https://itunes.apple.com/search?media=music&entity=album&limit=8&term=" + q);
  if (!res.ok) throw new Error("iTunes " + res.status);
  const wantTitle = simple(a.title), wantArtist = simple(a.artist);
  const hit = (await res.json()).results.find(r => {
    const t = simple(r.collectionName || ""), ar = simple(r.artistName || "");
    return t && (t.includes(wantTitle) || wantTitle.includes(t)) && (ar.includes(wantArtist) || wantArtist.includes(ar));
  });                                                // no good match = no cover (a wrong cover is worse than none)
  return hit ? hit.artworkUrl100.replace("100x100bb", "600x600bb") : null;
}
/* covers.json = a ready-made list (made once with covers-builder.html), so visitors don't have to look up every cover themselves */
let coverMap = {};
const coverMapReady = fetch("covers.json").then(r => r.ok ? r.json() : {}).then(m => { coverMap = m || {}; }).catch(() => {}).then(() => { coverMapReady.settled = true; });
const queue = []; let pumping = false;
function wantCover(a, box) {
  if (!coversAllowed()) { delete box.dataset.requested; return; }      // no consent (yet): nothing leaves the browser, the placeholder stays
  if (a.coverUrl) { showCover(box, a.coverUrl); return; }     // your own link from the spreadsheet wins
  if (!LOOK_UP_COVERS) return;
  const key = coverKey(a);
  if (key in coverMap) { if (coverMap[key]) showCover(box, coverMap[key]); return; }     // known from covers.json ("" = no cover exists)
  if (!coverMapReady.settled) { coverMapReady.then(() => { coverMapReady.settled = true; wantCover(a, box); }); return; }
  if (key in coverCache) { if (coverCache[key]) showCover(box, coverCache[key]); return; }
  queue.push({ a, box, key }); pump();
}
async function pump() {
  if (pumping) return; pumping = true;
  while (queue.length) {
    const job = queue.shift();
    if (!job.box.isConnected || !job.box.dataset.visible) continue;       // skip covers that were removed or scrolled away
    if (job.key in coverCache) { if (coverCache[job.key]) showCover(job.box, coverCache[job.key]); continue; }
    try {
      const url = await lookupCover(job.a);
      coverCache[job.key] = url || ""; saveCovers();
      if (url) showCover(job.box, url);
      await sleep(3200);                                                 // Apple allows about 20 lookups a minute; going faster gets us blocked
    } catch (e) { await sleep(8000); }                                   // blocked or offline: back off, keep the placeholder
  }
  pumping = false;
}
function onSeen(entries) {
  for (const e of entries) {
    const box = e.target;
    box.dataset.visible = e.isIntersecting ? "1" : "";
    if (e.isIntersecting && !box.dataset.requested) { box.dataset.requested = "1"; wantCover(box._album, box); }
  }
}
window.addEventListener("privacy-change", () => {
  queue.length = 0; imgQ.length = 0;
  if (!coversAllowed()) { coverCache = {}; document.querySelectorAll(".cover img").forEach(i => i.remove()); return; }
  loadCache(); warmFavorites(albums);
  document.querySelectorAll(".cover").forEach(c => { if (c._album && c.dataset.visible && !c.querySelector("img")) { c.dataset.requested = "1"; wantCover(c._album, c); } });
});
const io = new IntersectionObserver(onSeen, { rootMargin: "300px" });      // for the big list (reset on every search/sort)
const ioTop = new IntersectionObserver(onSeen, { rootMargin: "300px" });   // for everything at the top

function makeCover(a, observer) {
  const cover = h("div", "cover");
  cover.style.setProperty("--c", PLACEHOLDER_COLORS[hash(a.title) % PLACEHOLDER_COLORS.length]);
  cover.append(h("span", "initial", [...a.title][0] || "?"));
  cover._album = a; observer.observe(cover);
  return cover;
}

