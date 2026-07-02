import React, { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

// ============================================================
// Delhi Hockey — white-dominant, blue accents. Public site + admin.
// Data persists via Supabase.
// ============================================================

const BG = "#FCFDFE";        // page background — barely-there white
const BG2 = "#F7F9FB";        // soft off-white for sections / alt rows
const ACCENT = "#003366";     // primary blue (text, accents, buttons)
const ACCENT2 = "#00264D";    // deeper blue (hover)
const BLUE_SOFT = "#EEF3F8";  // pale blue wash
const GOLD = "#6FA8DC";       // light blue accent (was yellow)
const GOLD_DEEP = "#2E6CA6";  // mid blue for small labels / stats text
const TEXT = "#1A2A3A";       // dark ink for body text
const MUTE = "#5A6B82";       // muted blue-grey
const LINE = "#E4E9EE";       // soft neutral-white border
const ORANGE = "#FFA500";     // highlight accent (live, new, leader, open)
const ORANGE_DEEP = "#B5660A"; // orange text on light bg
const ORANGE_SOFT = "#FFF1DB"; // soft orange tint for badges


// per-page SEO metadata (title + meta description)
const SEO = {
  Home: { title: "Delhi Hockey — Official State Hockey Association", desc: "Delhi Hockey is the recognised state association affiliated to Hockey India and the Delhi Olympic Association, governing hockey across the National Capital Territory of Delhi." },
  About: { title: "About Us — Delhi Hockey", desc: "About Delhi Hockey: the official governing body for hockey in Delhi. View the selection committee, office bearers and financial audits." },
  League: { title: "Leagues & Standings — Delhi Hockey", desc: "Delhi Hockey league standings, fixtures, results, match reports and player statistics, season by season." },
  Documents: { title: "Documents & Notifications — Delhi Hockey", desc: "Official announcements, notices and downloadable publications from Delhi Hockey." },
  Calendar: { title: "Calendar — Delhi Hockey", desc: "The official Delhi Hockey calendar of tournaments, trials and events. Download the latest schedules." },
  Partners: { title: "Partners & Sponsors — Delhi Hockey", desc: "The partners, sponsors and supporters who help grow the game of hockey in Delhi." },
  Support: { title: "Support Us — Sponsor Delhi Hockey", desc: "Support Delhi Hockey through sponsorship or donation. Fund grassroots development, athlete support and equipment, and help grow hockey across the National Capital." },
  Contact: { title: "Contact — Delhi Hockey", desc: "Contact Delhi Hockey. Address: A-69/1 Okhla Phase 2, New Delhi 110020. Email contact.delhihockey@gmail.com." },
  FIH: { title: "FIH Rules — Delhi Hockey", desc: "Official FIH Rules of Hockey followed by Delhi Hockey, with links to the latest outdoor and indoor rulebooks." },
};

// nav: label shown -> internal page
const NAV = [
  { label: "HOME", page: "Home" },
  { label: "ABOUT DELHI HOCKEY", page: "About" },
  { label: "LEAGUE", page: "League" },
  { label: "DOCUMENTS", page: "Documents" },
  { label: "CALENDAR", page: "Calendar" },
  { label: "PARTNERS", page: "Partners" },
  { label: "SUPPORT US", page: "Support" },
  { label: "CONTACT", page: "Contact" },
  { label: "FIH RULES", page: "FIH" },
];

// ============================================================
// Supabase database connection (replaces preview-only storage)
// These two values come from your Supabase project's API page.
// The anon key is safe to expose in a public website.
// ============================================================
const SUPABASE_URL = "https://ebjlezchorfntbdtkuqu.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImViamxlemNob3JmbnRiZHRrdXF1Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODIwNDM2OTUsImV4cCI6MjA5NzYxOTY5NX0.FeTLXx2BHxDAoRphfPfSComcSCyp8Ac4GRhJH0ajdUQ";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

// optional global hook so a failed save can surface a warning to the user
let onSaveError = null;

// Load one value by key from the site_data table.
async function loadJSON(key, fallback) {
  try {
    const { data, error } = await supabase
      .from("site_data")
      .select("value")
      .eq("key", key)
      .maybeSingle();
    if (error) throw error;
    if (!data || data.value == null) return fallback;
    return JSON.parse(data.value);
  } catch (e) {
    console.error("Load failed for", key, e);
    return fallback;
  }
}

// Save one value by key (insert or update) into the site_data table.
async function saveJSON(key, value) {
  try {
    const { error } = await supabase
      .from("site_data")
      .upsert({ key, value: JSON.stringify(value), updated_at: new Date().toISOString() });
    if (error) throw error;
    return true;
  } catch (e) {
    console.error("Save failed for", key, e);
    if (onSaveError) onSaveError("Couldn't save — please check your connection and try again.");
    return false;
  }
}
function fmtDate(iso) { if (!iso) return ""; const d = new Date(iso); return isNaN(d) ? "" : d.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" }); }
// numeric value for sorting; invalid/empty dates sort to 0
function dateVal(iso) { if (!iso) return 0; const t = new Date(iso).getTime(); return isNaN(t) ? 0 : t; }
function fmtTime(t) { if (!t) return ""; const [h, m] = t.split(":").map(Number); if (isNaN(h)) return t; const ap = h >= 12 ? "PM" : "AM"; const h12 = ((h + 11) % 12) + 1; return `${h12}:${String(m).padStart(2, "0")} ${ap}`; }
function quarterOf(min) { const n = Number(min); if (!n || n < 1) return null; if (n <= 15) return 1; if (n <= 30) return 2; if (n <= 45) return 3; return 4; }
function fileKind(name = "", type = "") {
  const n = name.toLowerCase();
  if (n.endsWith(".pdf") || type.includes("pdf")) return "PDF";
  if (n.match(/\.(xlsx|xls|csv)$/) || type.includes("sheet") || type.includes("excel")) return "Excel";
  if (n.match(/\.(doc|docx)$/) || type.includes("word")) return "Word";
  return "File";
}
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
// Max raw file size. Base64 inflates ~33%, and storage caps at 5MB/key,
// so ~3.6MB raw keeps the encoded value safely under the limit.
const MAX_FILE_BYTES = 3.6 * 1024 * 1024;
function fileToDataURL(file) {
  return new Promise((res, rej) => {
    if (file && file.size > MAX_FILE_BYTES) {
      rej(new Error("FILE_TOO_LARGE"));
      return;
    }
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(new Error("READ_FAILED"));
    r.readAsDataURL(file);
  });
}

// icons
const I = {
  dl: (p) => <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="M7 10l5 5 5-5"/><path d="M12 15V3"/></svg>,
  arrow: (p) => <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M5 12h14"/><path d="M12 5l7 7-7 7"/></svg>,
  trash: (p) => <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M3 6h18"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>,
  up: (p) => <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M18 15l-6-6-6 6"/></svg>,
  down: (p) => <svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M6 9l6 6 6-6"/></svg>,
  ext: (p) => <svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" {...p}><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><path d="M15 3h6v6"/><path d="M10 14L21 3"/></svg>,
};
const SOC = {
  ig: <svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" strokeWidth="2"><rect x="2" y="2" width="20" height="20" rx="5"/><circle cx="12" cy="12" r="4"/><circle cx="17.5" cy="6.5" r="1" fill="currentColor" stroke="none"/></svg>,
  fb: <svg viewBox="0 0 24 24" width="17" height="17" fill="currentColor"><path d="M22 12a10 10 0 1 0-11.6 9.9v-7H7.9V12h2.5V9.8c0-2.5 1.5-3.9 3.8-3.9 1.1 0 2.2.2 2.2.2v2.5h-1.2c-1.2 0-1.6.8-1.6 1.5V12h2.7l-.4 2.9h-2.3v7A10 10 0 0 0 22 12z"/></svg>,
  li: <svg viewBox="0 0 24 24" width="17" height="17" fill="currentColor"><path d="M4.98 3.5A2.5 2.5 0 1 1 5 8.5a2.5 2.5 0 0 1 0-5zM3 9h4v12H3zM9 9h3.8v1.7h.05c.53-1 1.8-2 3.7-2 4 0 4.7 2.6 4.7 6V21H20v-5.3c0-1.3 0-2.9-1.8-2.9s-2 1.4-2 2.8V21H9z"/></svg>,
};

// ============================================================
export default function App() {
  const [route, setRoute] = useState("Home");
  const [admin, setAdmin] = useState(false);
  const [loaded, setLoaded] = useState(false);

  const [docs, setDocs] = useState([]);
  const [aboutItems, setAboutItems] = useState([]);
  const [calendarItems, setCalendarItems] = useState([]);
  const [partners, setPartners] = useState([]);
  const [aboutIntro, setAboutIntro] = useState("");
  const [fihItems, setFihItems] = useState([]);
  const [reg, setReg] = useState(null);
  const [players, setPlayers] = useState([]);
  const [leagues, setLeagues] = useState([]);
  const [globalSponsors, setGlobalSponsors] = useState([]);

  useEffect(() => {
    (async () => {
      // load everything in parallel for a faster first paint
      const [
        d, ab, cal, par, intro, fih, r, pl, gs, lgRaw,
      ] = await Promise.all([
        loadJSON("dh_docs", []),
        loadJSON("dh_about", DEFAULT_ABOUT),
        loadJSON("dh_calendar", []),
        loadJSON("dh_partners", []),
        loadJSON("dh_about_intro", DEFAULT_INTRO),
        loadJSON("dh_fih", DEFAULT_FIH),
        loadJSON("dh_reg", DEFAULT_REG),
        loadJSON("dh_players", []),
        loadJSON("dh_global_sponsors", []),
        loadJSON("dh_leagues", null),
      ]);
      setDocs(d);
      setAboutItems(ab);
      setCalendarItems(cal);
      setPartners(par);
      setAboutIntro(intro);
      setFihItems(fih);
      setReg(r);
      setPlayers(pl);
      setGlobalSponsors(gs);

      // load leagues; migrate old flat teams/matches if present and no leagues yet
      let lg = lgRaw;
      if (!lg) {
        const [oldTeams, oldMatches] = await Promise.all([
          loadJSON("dh_teams", []),
          loadJSON("dh_matches", []),
        ]);
        if (oldTeams.length || oldMatches.length) {
          lg = [{
            id: uid(), name: "Delhi Hockey League", sponsors: [],
            seasons: [{ id: uid(), year: String(new Date().getFullYear()), teams: oldTeams, matches: oldMatches }],
          }];
        } else {
          lg = [];
        }
        await saveJSON("dh_leagues", lg);
      }
      // migrate any old league-level sponsors down into each season (one-time)
      let migrated = false;
      lg = lg.map((l) => {
        if (l.sponsors && l.sponsors.length) {
          migrated = true;
          return { ...l, sponsors: [], seasons: (l.seasons || []).map((s) => ({ ...s, sponsors: [...(s.sponsors || []), ...l.sponsors] })) };
        }
        return l;
      });
      if (migrated) await saveJSON("dh_leagues", lg);
      setLeagues(lg);
      setLoaded(true);
    })();
  }, []);

  const go = (r) => { setRoute(r); window.scrollTo({ top: 0, behavior: "smooth" }); };

  // ---- SEO: keep document title, meta description and structured data in sync ----
  useEffect(() => {
    const meta = SEO[route] || SEO.Home;
    document.title = meta.title;

    const setMeta = (selector, attr, key, content) => {
      let el = document.head.querySelector(selector);
      if (!el) { el = document.createElement("meta"); el.setAttribute(attr, key); document.head.appendChild(el); }
      el.setAttribute("content", content);
    };
    setMeta('meta[name="description"]', "name", "description", meta.desc);
    setMeta('meta[property="og:title"]', "property", "og:title", meta.title);
    setMeta('meta[property="og:description"]', "property", "og:description", meta.desc);
    setMeta('meta[property="og:type"]', "property", "og:type", "website");
    setMeta('meta[name="twitter:card"]', "name", "twitter:card", "summary");

    // language + viewport (only once)
    if (document.documentElement.lang !== "en") document.documentElement.lang = "en";
    if (!document.head.querySelector('meta[name="viewport"]')) {
      const v = document.createElement("meta");
      v.setAttribute("name", "viewport");
      v.setAttribute("content", "width=device-width, initial-scale=1");
      document.head.appendChild(v);
    }
  }, [route]);

  // inject Organization structured data once
  useEffect(() => {
    if (document.getElementById("dh-jsonld")) return;
    const ld = document.createElement("script");
    ld.type = "application/ld+json";
    ld.id = "dh-jsonld";
    ld.text = JSON.stringify({
      "@context": "https://schema.org",
      "@type": "SportsOrganization",
      "name": "Delhi Hockey",
      "sport": "Field hockey",
      "description": "The official state association governing hockey in the National Capital Territory of Delhi, affiliated with Hockey India.",
      "address": {
        "@type": "PostalAddress",
        "streetAddress": "A-69/1 Okhla Phase – 2",
        "addressLocality": "New Delhi",
        "postalCode": "110020",
        "addressCountry": "IN",
      },
      "email": "contact.delhihockey@gmail.com",
      "sameAs": [
        "https://www.facebook.com/delhihockey/",
        "https://www.instagram.com/delhihockey/",
        "https://www.linkedin.com/company/delhi-hockey/",
      ],
    });
    document.head.appendChild(ld);
  }, []);

  return (
    <div style={{ fontFamily: "'Inter',system-ui,sans-serif", color: TEXT, background: BG, minHeight: "100vh" }}>
      <style>{CSS}</style>
      <Header route={route} go={go} />

      {!loaded ? <div style={{ padding: "140px 0", textAlign: "center", color: MUTE }}>Loading…</div> : (
        <main>
          {route === "Home" && <Home docs={docs} reg={reg} go={go} />}
          {route === "About" && <About intro={aboutIntro} items={aboutItems} />}
          {route === "League" && <League leagues={leagues} players={players} globalSponsors={globalSponsors} />}
          {route === "Documents" && <Documents docs={docs} />}
          {route === "Calendar" && <Calendar items={calendarItems} />}
          {route === "Partners" && <Partners partners={partners} />}
          {route === "Support" && <Support />}
          {route === "Contact" && <Contact />}
          {route === "FIH" && <FIH items={fihItems} />}
        </main>
      )}

      <Footer go={go} onAdmin={() => setAdmin(true)} />
      {admin && <Admin onClose={() => setAdmin(false)}
        state={{ docs, aboutItems, calendarItems, partners, aboutIntro, fihItems, reg, players, leagues, globalSponsors }}
        setters={{ setDocs, setAboutItems, setCalendarItems, setPartners, setAboutIntro, setFihItems, setReg, setPlayers, setLeagues, setGlobalSponsors }} />}
    </div>
  );
}

// ---------- header ----------
function Header({ route, go }) {
  const [open, setOpen] = useState(false);
  return (
    <header className="hdr">
      <div className="wrap hdr-in">
        <button className="wordmark-sm" onClick={() => go("Home")}>
          DELHI<span> HOCKEY</span>
        </button>
        <nav className="dnav">
          {NAV.map((n) => (
            <button key={n.page} onClick={() => go(n.page)} className={"nav" + (route === n.page ? " on" : "")}>{n.label}</button>
          ))}
        </nav>
        <button className="burger" onClick={() => setOpen(!open)} aria-label="Menu"><span/><span/><span/></button>
      </div>
      {open && <nav className="mnav">{NAV.map((n) => (
        <button key={n.page} onClick={() => { go(n.page); setOpen(false); }} className={"mnav-l" + (route === n.page ? " on" : "")}>{n.label}</button>
      ))}</nav>}
    </header>
  );
}

// ---------- home ----------
function Home({ docs, reg, go }) {
  const sorted = [...docs].sort((a, b) => dateVal(b.date) - dateVal(a.date));
  return (
    <>
      <section className="hero">
        <div className="wrap hero-grid">
          <div className="hero-left">
            <h1 className="mega">DELHI<br/>HOCKEY</h1>
            <p className="hero-copy">
              Delhi Hockey is the recognised state association affiliated to the Hockey India (HI) & Delhi Olympic Association.
            </p>
            <div className="socs">
              <a className="soc" href="https://www.instagram.com/delhihockey/" target="_blank" rel="noreferrer">{SOC.ig}</a>
              <a className="soc" href="https://www.facebook.com/delhihockey/" target="_blank" rel="noreferrer">{SOC.fb}</a>
              <a className="soc" href="https://www.linkedin.com/company/delhi-hockey/" target="_blank" rel="noreferrer">{SOC.li}</a>
            </div>
            <button className="cta" onClick={() => go("Documents")}>VIEW DOCUMENTS</button>
          </div>

          <LiveDocuments docs={sorted} go={go} />
        </div>
      </section>

      <Registration reg={reg} />
    </>
  );
}

// ---------- registration section ----------
function Registration({ reg }) {
  if (!reg) return null;
  return (
    <section className="reg">
      <div className="wrap">
        <h2 className="reg-h">Registration</h2>
        <p className="reg-desc">{reg.desc}</p>
        <div className="reg-grid">
          {reg.forms.map((f) => <RegCard key={f.key} f={f} />)}
        </div>
      </div>
    </section>
  );
}

function RegCard({ f }) {
  const live = f.status === "live" && f.url;
  return (
    <div className={"reg-card" + (live ? "" : " soon")}>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 }}>
        <div className="reg-name">{f.label}</div>
        {live && <span className="pill-open">OPEN</span>}
      </div>
      {live ? (
        <>
          <a className="reg-btn" href={f.url} target="_blank" rel="noreferrer">Register {I.ext()}</a>
          {f.guide && <a className="reg-guide" href={f.guide} download={f.guideName || "guide.pdf"}>How to fill — guide {I.dl()}</a>}
        </>
      ) : (
        <div className="reg-soon">Coming soon</div>
      )}
    </div>
  );
}

function LiveDocuments({ docs, go }) {
  const ref = useRef(null);
  const [paused, setPaused] = useState(false);
  useEffect(() => {
    const el = ref.current;
    if (!el || docs.length <= 4) return;
    // respect users who prefer reduced motion — no auto-scroll for them
    const reduce = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduce || paused) return;

    let raf;
    const step = () => {
      // pause when the browser tab isn't visible (saves CPU/battery)
      if (!document.hidden && el) {
        el.scrollTop += 0.4;
        if (el.scrollTop >= el.scrollHeight - el.clientHeight - 1) el.scrollTop = 0;
      }
      raf = requestAnimationFrame(step);
    };
    raf = requestAnimationFrame(step);
    return () => cancelAnimationFrame(raf);
  }, [paused, docs.length]);

  return (
    <div className="live">
      <button className="live-head" onClick={() => go("Documents")} aria-label="View all documents">
        <span className="live-dot" aria-hidden="true" /> LIVE DOCUMENTS
        <span className="live-all" aria-hidden="true">{I.arrow()}</span>
      </button>
      <div className="live-track" ref={ref} onMouseEnter={() => setPaused(true)} onMouseLeave={() => setPaused(false)}>
        {docs.length === 0 && <div className="live-empty">No documents yet. Announcements will appear here.</div>}
        {docs.map((d) => <LiveRow key={d.id} d={d} />)}
        {docs.length > 4 && docs.map((d) => <LiveRow key={d.id + "_l"} d={d} aria-hidden="true" />)}
      </div>
    </div>
  );
}
function LiveRow({ d }) {
  const isNew = (() => {
    if (!d.date) return false;
    const t = new Date(d.date).getTime();
    if (isNaN(t)) return false;
    return (Date.now() - t) < 14 * 24 * 60 * 60 * 1000 && t <= Date.now() + 24 * 60 * 60 * 1000;
  })();
  return (
    <div className="live-row">
      <div style={{ flex: 1, minWidth: 0 }}>
        <div className="live-date">{fmtDate(d.date)}</div>
        <div className="live-title">{d.title}</div>
      </div>
      {isNew && <span className="badge-new">NEW</span>}
      <a className="live-dl" href={d.dataUrl} download={d.fileName || d.title} title={"Download " + d.kind} aria-label={"Download " + d.title}>{I.dl()}</a>
    </div>
  );
}

// ============================================================
// LEAGUE — helpers
// ============================================================
function teamName(teams, id) { const t = teams.find((x) => x.id === id); return t ? t.name : "—"; }
// resolve a Unique ID to a registered name; falls back to showing the ID
function playerName(players, pid) {
  if (!pid) return "—";
  const p = (players || []).find((x) => x.id === pid);
  return p ? p.name : pid;
}

function computeStandings(teams, matches) {
  const table = {};
  teams.forEach((t) => { table[t.id] = { id: t.id, name: t.name, P: 0, W: 0, D: 0, L: 0, GF: 0, GA: 0, GD: 0, Pts: 0 }; });
  matches.filter((m) => m.status === "played").forEach((m) => {
    const h = table[m.home], a = table[m.away];
    if (!h || !a) return;
    const hg = Number(m.homeGoals) || 0, ag = Number(m.awayGoals) || 0;
    h.P++; a.P++; h.GF += hg; h.GA += ag; a.GF += ag; a.GA += hg;
    if (hg > ag) { h.W++; a.L++; h.Pts += 3; }
    else if (hg < ag) { a.W++; h.L++; a.Pts += 3; }
    else { h.D++; a.D++; h.Pts++; a.Pts++; }
  });
  return Object.values(table).map((r) => ({ ...r, GD: r.GF - r.GA }))
    .sort((a, b) => b.Pts - a.Pts || b.GD - a.GD || b.GF - a.GF || a.name.localeCompare(b.name));
}

function computePlayerStats(teams, matches) {
  const stats = {}; // key: teamId|player
  const bump = (team, player, field, n = 1) => {
    if (!player) return;
    const k = team + "|" + player;
    if (!stats[k]) stats[k] = { player, team, goals: 0, assists: 0, green: 0, yellow: 0, red: 0 };
    stats[k][field] += n;
  };
  matches.filter((m) => m.status === "played").forEach((m) => {
    (m.events || []).forEach((e) => {
      if (e.type === "goal") bump(e.team, e.player, "goals");
      if (e.type === "assist") bump(e.team, e.player, "assists");
      if (e.type === "green") bump(e.team, e.player, "green");
      if (e.type === "yellow") bump(e.team, e.player, "yellow");
      if (e.type === "red") bump(e.team, e.player, "red");
    });
  });
  return Object.values(stats);
}

// ============================================================
// EXPORT — build a multi-sheet Excel workbook from all data
// ============================================================
function quarterLabel(min) { const q = quarterOf(min); return q ? `Q${q}` : ""; }

function buildExportRows(leagues, players, globalSponsors, scope) {
  // scope = { all:true } or { leagueId, seasonId }
  const reg = (players || []).map((p) => ({ "Unique ID": p.id, "Name": p.name }));

  const teamRows = [];     // squad members
  const coachRows = [];
  const schedRows = [];
  const eventRows = [];
  const refRows = [];
  const standingRows = [];
  const playerStatRows = [];
  const sponsorRows = [];

  // global sponsors
  (globalSponsors || []).forEach((s) => sponsorRows.push({ "Scope": "Global", "League": "—", "Season": "—", "Sponsor": s.name, "Title": s.title || "", "Website": s.url || "" }));

  const targetLeagues = scope.all ? leagues : leagues.filter((l) => l.id === scope.leagueId);

  targetLeagues.forEach((l) => {
    const seasons = scope.all ? (l.seasons || []) : (l.seasons || []).filter((s) => s.id === scope.seasonId);
    seasons.forEach((s) => {
      const teams = s.teams || [];
      const matches = s.matches || [];
      const tname = (id) => { const t = teams.find((x) => x.id === id); return t ? t.name : "—"; };

      // sponsors for this season
      (s.sponsors || []).forEach((sp) => sponsorRows.push({ "Scope": "Season", "League": l.name, "Season": s.year, "Sponsor": sp.name, "Title": sp.title || "", "Website": sp.url || "" }));

      // squads + coaches
      teams.forEach((t) => {
        (t.players || []).forEach((pid) => {
          teamRows.push({ "League": l.name, "Season": s.year, "Team": t.name, "Player Unique ID": pid, "Player Name": playerName(players, pid) });
        });
        (t.coaches || []).forEach((c) => {
          coachRows.push({ "League": l.name, "Season": s.year, "Team": t.name, "Coach Name": c.name, "Coach ID": c.regId || "" });
        });
      });

      // schedule / results + referees + events
      [...matches].sort((a, b) => dateVal(a.date) - dateVal(b.date)).forEach((m) => {
        const played = m.status === "played";
        schedRows.push({
          "League": l.name, "Season": s.year,
          "Date": m.date || "", "Time": m.time ? fmtTime(m.time) : "", "Venue": m.venue || "",
          "Home Team": tname(m.home), "Away Team": tname(m.away),
          "Status": played ? "Played" : "Scheduled",
          "Home Goals": played ? (m.homeGoals || 0) : "", "Away Goals": played ? (m.awayGoals || 0) : "",
          "Home Penalty Corners": m.homePC || 0, "Away Penalty Corners": m.awayPC || 0,
          "Home Circle Entries": m.homeCE || 0, "Away Circle Entries": m.awayCE || 0,
          "Player of the Match": m.pom ? playerName(players, m.pom) : "",
          "Referees": (m.referees || []).join("; "),
        });
        (m.referees || []).forEach((r) => refRows.push({ "League": l.name, "Season": s.year, "Match": `${tname(m.home)} vs ${tname(m.away)}`, "Date": m.date || "", "Referee": r }));
        (m.events || []).forEach((e) => eventRows.push({
          "League": l.name, "Season": s.year, "Match": `${tname(m.home)} vs ${tname(m.away)}`, "Date": m.date || "",
          "Minute": e.min || "", "Quarter": quarterLabel(e.min), "Event": e.type,
          "Player": playerName(players, e.player), "Player Unique ID": e.player, "Team": tname(e.team),
        }));
      });

      // computed standings
      computeStandings(teams, matches).forEach((r, i) => standingRows.push({
        "League": l.name, "Season": s.year, "Position": i + 1, "Team": r.name,
        "Played": r.P, "Won": r.W, "Drawn": r.D, "Lost": r.L, "GF": r.GF, "GA": r.GA, "GD": r.GD, "Points": r.Pts,
      }));

      // computed player stats
      computePlayerStats(teams, matches).forEach((st) => playerStatRows.push({
        "League": l.name, "Season": s.year, "Player": playerName(players, st.player), "Player Unique ID": st.player,
        "Team": tname(st.team), "Goals": st.goals, "Assists": st.assists,
        "Green Cards": st.green, "Yellow Cards": st.yellow, "Red Cards": st.red,
      }));
    });
  });

  return {
    "Player Registry": reg,
    "Squads": teamRows,
    "Coaches": coachRows,
    "Schedule & Results": schedRows,
    "Match Events": eventRows,
    "Referees": refRows,
    "Standings": standingRows,
    "Player Stats": playerStatRows,
    "Sponsors": sponsorRows,
  };
}

function downloadWorkbook(sheets, filename) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    const data = rows.length ? rows : [{ "—": "No data" }];
    const ws = XLSX.utils.json_to_sheet(data);
    // auto column widths
    const cols = Object.keys(data[0] || {}).map((k) => ({ wch: Math.min(40, Math.max(k.length + 2, ...data.map((r) => String(r[k] ?? "").length + 2))) }));
    ws["!cols"] = cols;
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

// ============================================================
// LEAGUE — public page
// ============================================================
function League({ leagues, players, globalSponsors }) {
  const [tab, setTab] = useState("Standings");
  const [leagueId, setLeagueId] = useState(leagues[0]?.id || "");
  const TABS = ["Standings", "Fixtures & Results", "Player Stats", "Teams"];

  const league = leagues.find((l) => l.id === leagueId) || leagues[0];
  // seasons sorted newest year first
  const seasons = league ? [...(league.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year)) : [];
  const [seasonId, setSeasonId] = useState(seasons[0]?.id || "");
  // keep season valid when league changes
  const season = seasons.find((s) => s.id === seasonId) || seasons[0];

  const onLeague = (id) => {
    setLeagueId(id);
    const l = leagues.find((x) => x.id === id);
    const ss = [...(l?.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year));
    setSeasonId(ss[0]?.id || "");
  };

  const sponsors = [...(globalSponsors || []), ...((season && season.sponsors) || [])];
  const teams = season?.teams || [];
  const matches = season?.matches || [];

  return (
    <Page title="League" sub="Delhi Hockey leagues — standings, fixtures, results and player statistics, season by season.">
      {leagues.length === 0 ? (
        <Empty>No leagues have been set up yet. They will appear here once Delhi Hockey adds a competition.</Empty>
      ) : (
        <>
          <div className="lg-selectors">
            <label className="lg-sel"><span>League</span>
              <select className="inp" value={league?.id || ""} onChange={(e) => onLeague(e.target.value)}>
                {leagues.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
              </select>
            </label>
            <label className="lg-sel"><span>Season</span>
              <select className="inp" value={season?.id || ""} onChange={(e) => setSeasonId(e.target.value)} disabled={seasons.length === 0}>
                {seasons.length === 0 ? <option>No seasons</option> : seasons.map((s) => <option key={s.id} value={s.id}>{s.year}</option>)}
              </select>
            </label>
          </div>

          {sponsors.length > 0 && <SponsorStrip sponsors={sponsors} />}

          {!season ? (
            <Empty>This league has no seasons yet.</Empty>
          ) : (teams.length === 0 && matches.length === 0) ? (
            <Empty>Nothing has been published for {league.name} — {season.year} yet.</Empty>
          ) : (
            <>
              <div className="ltabs">
                {TABS.map((t) => <button key={t} className={"ltab" + (tab === t ? " on" : "")} onClick={() => setTab(t)}>{t}</button>)}
              </div>
              <div style={{ marginTop: 24 }}>
                {tab === "Standings" && <StandingsView teams={teams} matches={matches} />}
                {tab === "Fixtures & Results" && <FixturesView teams={teams} matches={matches} players={players} />}
                {tab === "Player Stats" && <PlayerStatsView teams={teams} matches={matches} players={players} />}
                {tab === "Teams" && <TeamsView teams={teams} players={players} />}
              </div>
            </>
          )}
        </>
      )}
    </Page>
  );
}

function SponsorStrip({ sponsors }) {
  return (
    <div className="sp-strip">
      <div className="sp-strip-h">Our Sponsors</div>
      <div className="sp-grid">
        {sponsors.map((s) => (
          <a key={s.id} className="sp-card" href={s.url || undefined} target={s.url ? "_blank" : undefined} rel="noreferrer">
            {s.logo ? <img src={s.logo} alt={s.name} /> : <div className="sp-fallback">{s.name?.[0] || "?"}</div>}
            <div className="sp-name">{s.name}</div>
            {s.title && <div className="sp-title">{s.title}</div>}
          </a>
        ))}
      </div>
    </div>
  );
}

function StandingsView({ teams, matches }) {
  const rows = computeStandings(teams, matches);
  if (rows.length === 0) return <Empty>No teams added yet.</Empty>;
  return (
    <div className="tbl-wrap">
      <table className="tbl">
        <thead>
          <tr><th className="l">#</th><th className="l">Team</th><th>P</th><th>W</th><th>D</th><th>L</th><th>GF</th><th>GA</th><th>GD</th><th className="pts">Pts</th></tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={r.id} className={i === 0 ? "leader" : ""}>
              <td className="rank">{i + 1}</td>
              <td className="l team">{r.name}</td>
              <td>{r.P}</td><td>{r.W}</td><td>{r.D}</td><td>{r.L}</td>
              <td>{r.GF}</td><td>{r.GA}</td><td>{r.GD > 0 ? "+" + r.GD : r.GD}</td>
              <td className="pts">{r.Pts}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function FixturesView({ teams, matches, players }) {
  const [openId, setOpenId] = useState(null);
  const sorted = [...matches].sort((a, b) => dateVal(a.date) - dateVal(b.date));
  if (sorted.length === 0) return <Empty>No matches scheduled yet.</Empty>;
  return (
    <div className="fix-list">
      {sorted.map((m) => {
        const played = m.status === "played";
        const open = openId === m.id;
        return (
          <div className="fix" key={m.id}>
            <div className="fix-meta">{m.date ? fmtDate(m.date) : "TBD"}{m.time ? ` · ${fmtTime(m.time)}` : ""}{m.venue ? ` · ${m.venue}` : ""}</div>
            <div className="fix-row">
              <span className="fix-team h">{teamName(teams, m.home)}</span>
              <span className="fix-score">{played ? `${m.homeGoals} – ${m.awayGoals}` : "vs"}</span>
              <span className="fix-team a">{teamName(teams, m.away)}</span>
            </div>
            {played && (
              <div className="fix-report-toggle">
                <button className="report-btn" onClick={() => setOpenId(open ? null : m.id)}>
                  {open ? "Hide match report" : "View match report"} {open ? I.down() : I.arrow()}
                </button>
              </div>
            )}
            {played && open && <MatchReport m={m} teams={teams} players={players} />}
          </div>
        );
      })}
    </div>
  );
}

function MatchReport({ m, teams, players }) {
  const events = [...(m.events || [])].sort((a, b) => (a.min || 999) - (b.min || 999));
  const byQuarter = { 1: [], 2: [], 3: [], 4: [], 0: [] };
  events.forEach((e) => { const q = quarterOf(e.min) || 0; byQuarter[q].push(e); });
  const EV_LABELS = { goal: "Goal", assist: "Assist", green: "Green card", yellow: "Yellow card", red: "Red card" };
  const homeName = teamName(teams, m.home), awayName = teamName(teams, m.away);
  const pom = m.pom ? playerName(players, m.pom) : "";
  const refs = m.referees || [];
  const hasEvents = events.length > 0;

  return (
    <div className="report">
      {pom && <div className="report-pom">★ Player of the Match — <strong>{pom}</strong></div>}

      <div className="report-stats">
        <div className="rs-row"><span className="rs-v">{m.homePC || 0}</span><span className="rs-l">Penalty Corners</span><span className="rs-v">{m.awayPC || 0}</span></div>
        <div className="rs-row"><span className="rs-v">{m.homeCE || 0}</span><span className="rs-l">Circle Entries</span><span className="rs-v">{m.awayCE || 0}</span></div>
        <div className="rs-teams"><span>{homeName}</span><span>{awayName}</span></div>
      </div>

      {hasEvents ? (
        <div className="report-timeline">
          {[1, 2, 3, 4].map((q) => byQuarter[q].length > 0 && (
            <div className="rq" key={q}>
              <div className="rq-head">Quarter {q} <span>({(q - 1) * 15 + 1}–{q * 15} min)</span></div>
              {byQuarter[q].map((e) => (
                <div className={"rq-ev " + (e.team === m.away ? "away" : "home")} key={e.id}>
                  <span className="rq-min">{e.min ? `${e.min}'` : ""}</span>
                  <span className={"ev-tag " + e.type}>{EV_LABELS[e.type]}</span>
                  <span className="rq-name">{playerName(players, e.player)}</span>
                  <span className="rq-team">{teamName(teams, e.team)}</span>
                </div>
              ))}
            </div>
          ))}
          {byQuarter[0].length > 0 && (
            <div className="rq">
              <div className="rq-head">Other events</div>
              {byQuarter[0].map((e) => (
                <div className={"rq-ev " + (e.team === m.away ? "away" : "home")} key={e.id}>
                  <span className="rq-min">—</span>
                  <span className={"ev-tag " + e.type}>{EV_LABELS[e.type]}</span>
                  <span className="rq-name">{playerName(players, e.player)}</span>
                  <span className="rq-team">{teamName(teams, e.team)}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      ) : <p className="muted2" style={{ margin: "4px 0" }}>No individual events were recorded for this match.</p>}

      {refs.length > 0 && (
        <div className="report-refs">
          <span className="rr-h">{refs.length > 1 ? "Match Officials" : "Match Official"}</span>
          <span>{refs.join(" · ")}</span>
        </div>
      )}
    </div>
  );
}

function PlayerStatsView({ teams, matches, players }) {
  const stats = computePlayerStats(teams, matches);
  if (stats.length === 0) return <Empty>No player statistics recorded yet.</Empty>;
  const scorers = [...stats].filter((s) => s.goals > 0).sort((a, b) => b.goals - a.goals || b.assists - a.assists).slice(0, 15);
  const assisters = [...stats].filter((s) => s.assists > 0).sort((a, b) => b.assists - a.assists).slice(0, 10);
  const cards = [...stats].filter((s) => s.green + s.yellow + s.red > 0).sort((a, b) => (b.red * 3 + b.yellow * 2 + b.green) - (a.red * 3 + a.yellow * 2 + a.green)).slice(0, 12);
  return (
    <div className="stat-cols">
      <div>
        <h3 className="stat-h">Top Scorers</h3>
        <div className="stat-list">
          {scorers.map((s, i) => (
            <div className="stat-row" key={s.team + s.player}>
              <span className="stat-rank">{i + 1}</span>
              <div style={{ flex: 1 }}><div className="stat-name">{playerName(players, s.player)}</div><div className="stat-team">{teamName(teams, s.team)}</div></div>
              <span className="stat-val">{s.goals}<small>G</small></span>
            </div>
          ))}
        </div>
      </div>
      <div>
        <h3 className="stat-h">Top Assists</h3>
        <div className="stat-list">
          {assisters.length === 0 ? <p className="muted2">No assists yet.</p> : assisters.map((s, i) => (
            <div className="stat-row" key={s.team + s.player}>
              <span className="stat-rank">{i + 1}</span>
              <div style={{ flex: 1 }}><div className="stat-name">{playerName(players, s.player)}</div><div className="stat-team">{teamName(teams, s.team)}</div></div>
              <span className="stat-val">{s.assists}<small>A</small></span>
            </div>
          ))}
        </div>
        <h3 className="stat-h" style={{ marginTop: 26 }}>Cards</h3>
        <div className="stat-list">
          {cards.length === 0 ? <p className="muted2">No cards issued.</p> : cards.map((s) => (
            <div className="stat-row" key={s.team + s.player}>
              <div style={{ flex: 1 }}><div className="stat-name">{playerName(players, s.player)}</div><div className="stat-team">{teamName(teams, s.team)}</div></div>
              <span className="cards">
                {!!s.green && <span className="card g">{s.green}</span>}
                {!!s.yellow && <span className="card y">{s.yellow}</span>}
                {!!s.red && <span className="card r">{s.red}</span>}
              </span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function TeamsView({ teams, players }) {
  if (teams.length === 0) return <Empty>No teams added yet.</Empty>;
  return (
    <div className="sq-grid">
      {teams.map((t) => (
        <div className="sq-card" key={t.id}>
          <div className="sq-name">{t.name}</div>
          {(t.coaches && t.coaches.length > 0) && (
            <div className="sq-coaches">
              <div className="sq-coach-h">{t.coaches.length > 1 ? "Coaches" : "Coach"}</div>
              {t.coaches.map((c) => <div className="sq-coach" key={c.cid}>{c.name}{c.regId ? <span className="mono"> · {c.regId}</span> : ""}</div>)}
            </div>
          )}
          {(t.players && t.players.length) ? (
            <ol className="sq-list">{t.players.map((pid) => (
              <li key={pid}><span>{playerName(players, pid)}</span><span className="sq-id">{pid}</span></li>
            ))}</ol>
          ) : <p className="muted2">Squad not added yet.</p>}
        </div>
      ))}
    </div>
  );
}

// ---------- about ----------
function About({ intro, items }) {
  return (
    <Page title="About Delhi Hockey" sub="The official governing body for Hockey in the National Capital Territory of Delhi.">
      <p className="lead">{intro}</p>
      <DocList items={items} numbered />
    </Page>
  );
}

// ---------- documents (was notifications) ----------
function Documents({ docs }) {
  const sorted = [...docs].sort((a, b) => dateVal(b.date) - dateVal(a.date));
  return (
    <Page title="Documents" sub="Announcements, notices and publications from Delhi Hockey.">
      <DocList items={sorted} showDate />
    </Page>
  );
}

// ---------- calendar ----------
function Calendar({ items }) {
  return (
    <Page title="Calendar" sub="Tournaments, trials and events. Download the latest schedules below.">
      <DocList items={[...items].sort((a, b) => dateVal(b.date) - dateVal(a.date))} showDate />
    </Page>
  );
}

// ---------- partners ----------
function Partners({ partners }) {
  return (
    <Page title="Partners" sub="The supporters and sponsors who help grow the game in Delhi.">
      {partners.length === 0 ? <Empty>Partners will be featured here soon.</Empty> : (
        <div className="pgrid">
          {partners.map((p) => (
            <a key={p.id} className="pcard" href={p.url || undefined} target="_blank" rel="noreferrer">
              {p.logo ? <img src={p.logo} alt={p.name}/> : <div className="pfall">{p.name?.[0] || "?"}</div>}
              <div className="pname">{p.name}</div>
              {p.url && <div className="plink">Visit site {I.ext()}</div>}
            </a>
          ))}
        </div>
      )}
    </Page>
  );
}

// ---------- contact ----------
function Contact() {
  return (
    <Page title="Contact" sub="Get in touch with the association.">
      <div className="cgrid">
        <div className="ccard">
          <div className="clabel">Address</div>
          <p className="cbody">DELHI HOCKEY<br/>A-69/1 Okhla Phase – 2<br/>New Delhi – 110020</p>
        </div>
        <div className="ccard">
          <div className="clabel">Email</div>
          <p className="cbody">
            <a className="clink" href="mailto:contact.delhihockey@gmail.com">contact.delhihockey@gmail.com</a><br/>
            <a className="clink" href="mailto:delhihockey@hockeyindia.org">delhihockey@hockeyindia.org</a>
          </p>
        </div>
        <div className="ccard">
          <div className="clabel">Follow</div>
          <div className="socs" style={{ marginTop: 12 }}>
            <a className="soc" href="https://www.instagram.com/delhihockey/" target="_blank" rel="noreferrer">{SOC.ig}</a>
            <a className="soc" href="https://www.facebook.com/delhihockey/" target="_blank" rel="noreferrer">{SOC.fb}</a>
            <a className="soc" href="https://www.linkedin.com/company/delhi-hockey/" target="_blank" rel="noreferrer">{SOC.li}</a>
          </div>
        </div>
      </div>
    </Page>
  );
}

// ---------- FIH ----------
function FIH({ items }) {
  return (
    <Page title="FIH Rules" sub="Delhi Hockey follows the official Rules of Hockey published by the International Hockey Federation.">
      {(!items || items.length === 0) ? <Empty>Rules & regulations will be published here soon.</Empty> : (
        <div className="dlist">
          {items.map((it) => {
            const href = it.dataUrl || it.url;
            const isFile = !!it.dataUrl;
            return (
              <a className="drow link" key={it.id} href={href || undefined} target="_blank" rel="noreferrer"
                 download={isFile ? (it.fileName || it.title) : undefined}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div className="dtitle">{it.title}</div>
                  {it.sub && <div className="dsub">{it.sub}</div>}
                </div>
                {href ? (
                  <span className="pill solid">{isFile ? <>Download {I.dl()}</> : <>Open {I.ext()}</>}</span>
                ) : <span className="pill ghost">No link</span>}
              </a>
            );
          })}
        </div>
      )}
    </Page>
  );
}

// ---------- Support / Sponsorship ----------
const SUPPORT_MAIL = "mailto:contact.delhihockey@gmail.com?subject=Supporting%20Delhi%20Hockey";

function Support() {
  const pillars = [
    { title: "Grassroots & youth development", body: "Coaching camps, school programmes and talent scouting that put sticks in the hands of the next generation across Delhi's neighbourhoods." },
    { title: "Athlete support & travel", body: "Kit, nutrition, coaching and travel costs so our players can compete at state and national tournaments without financial barriers holding them back." },
    { title: "Equipment & ground upkeep", body: "Balls, goals, gear and ground bookings for training — the everyday essentials that keep the game running safely and professionally." },
  ];
  const tiers = [
    { name: "Principal Partner", scope: "Season-long lead sponsorship", perks: ["Logo on team kit & event branding", "Featured on the Partners page & socials", "Named association with all Delhi Hockey events"] },
    { name: "Supporting Partner", scope: "Programme or tournament sponsor", perks: ["Branding at sponsored events", "Recognition on our website & socials", "CSR-aligned impact reporting"] },
    { name: "Community Backer", scope: "Individuals & small businesses", perks: ["Your name among our supporters", "Updates on the players you help", "The pride of growing Delhi hockey"] },
  ];
  return (
    <Page title="Support Delhi Hockey" sub="Help us build the future of hockey in the National Capital. Whether you're a brand seeking meaningful CSR impact or an individual who loves the game, your support powers everything we do.">

      <div className="sup-vision">
        <div className="sup-vision-tag">OUR VISION</div>
        <p className="sup-vision-text">A Delhi where every child with the talent and the will to play hockey has a pathway — from their first stick at a school ground to representing the nation. As the official state association affiliated with Hockey India, we are the custodians of that pathway. But nurturing champions and running the daily business of a sport takes resources, and that's where you come in.</p>
      </div>

      <h2 className="sup-h2">Where your support goes</h2>
      <div className="sup-grid">
        {pillars.map((p, i) => (
          <div className="sup-card" key={i}>
            <div className="sup-card-bar"></div>
            <div className="sup-card-title">{p.title}</div>
            <div className="sup-card-body">{p.body}</div>
          </div>
        ))}
      </div>

      <h2 className="sup-h2">Why partner with us</h2>
      <div className="sup-why">
        <div className="sup-why-item"><span className="sup-why-num">01</span><div><b>Official & credible.</b> The recognised governing body for hockey across the National Capital Territory of Delhi, affiliated with Hockey India.</div></div>
        <div className="sup-why-item"><span className="sup-why-num">02</span><div><b>Real grassroots reach.</b> Your support touches players, schools and grounds directly — not layers of overhead.</div></div>
        <div className="sup-why-item"><span className="sup-why-num">03</span><div><b>Visible impact.</b> Association with a growing, respected sport and a transparent partner that publishes its activities openly.</div></div>
      </div>

      <h2 className="sup-h2">Ways to contribute</h2>
      <p className="sup-intro">Everyone can play a part — from national brands to individuals who simply love the game. Choose the level that fits you.</p>
      <div className="sup-tiers">
        {tiers.map((t, i) => (
          <div className={"sup-tier" + (i === 0 ? " feat" : "")} key={i}>
            {i === 0 && <div className="sup-tier-flag">MOST IMPACT</div>}
            <div className="sup-tier-name">{t.name}</div>
            <div className="sup-tier-scope">{t.scope}</div>
            <ul className="sup-tier-perks">
              {t.perks.map((perk, j) => <li key={j}>{perk}</li>)}
            </ul>
          </div>
        ))}
      </div>

      <div className="sup-cta">
        <div className="sup-cta-title">Ready to make an impact?</div>
        <div className="sup-cta-body">We're currently building partnerships for the coming season. Tell us a little about you or your brand, and we'll share our detailed sponsorship deck and how your contribution will be put to work.</div>
        <a className="cta sup-cta-btn" href={SUPPORT_MAIL}>Get in touch to support us</a>
        <div className="sup-cta-note">Prefer email? Write to <a className="clink" href="mailto:contact.delhihockey@gmail.com">contact.delhihockey@gmail.com</a></div>
      </div>

    </Page>
  );
}

// ---------- shared ----------
function Page({ title, sub, children }) {
  return (
    <section className="page">
      <div className="wrap">
        <h1 className="ptitle">{title}</h1>
        {sub && <p className="psub">{sub}</p>}
        <div style={{ marginTop: 34 }}>{children}</div>
      </div>
    </section>
  );
}

function DocList({ items, numbered, showDate }) {
  if (!items || items.length === 0) return <Empty>Nothing here yet.</Empty>;
  return (
    <div className="dlist">
      {items.map((d, i) => (
        <div className="drow" key={d.id}>
          {numbered && <div className="dnum">{String(i + 1).padStart(2, "0")}</div>}
          <div style={{ flex: 1, minWidth: 0 }}>
            <div className="dtitle">{d.title}</div>
            {(showDate || d.date) && <div className="dsub">{showDate ? "Published: " : "Updated "}{fmtDate(d.date)}{d.kind && d.dataUrl ? ` · ${d.kind}` : ""}</div>}
          </div>
          {d.dataUrl ? (
            <a className="pill solid" href={d.dataUrl} download={d.fileName || d.title}><span className="pill-t">Download</span>{I.dl()}</a>
          ) : <span className="pill ghost">No file</span>}
        </div>
      ))}
    </div>
  );
}

function Empty({ children }) { return <div className="empty">{children}</div>; }

function Footer({ go, onAdmin }) {
  return (
    <footer className="ftr">
      <div className="wrap ftr-in">
        <div>
          <div className="wordmark-sm" style={{ cursor: "default" }}>DELHI<span> HOCKEY</span></div>
          <p className="ftr-copy">The official governing body for Hockey in the National Capital Territory of Delhi. Affiliated with Hockey India.</p>
          <div className="socs" style={{ marginTop: 16 }}>
            <a className="soc sm" href="https://www.instagram.com/delhihockey/" target="_blank" rel="noreferrer">{SOC.ig}</a>
            <a className="soc sm" href="https://www.facebook.com/delhihockey/" target="_blank" rel="noreferrer">{SOC.fb}</a>
            <a className="soc sm" href="https://www.linkedin.com/company/delhi-hockey/" target="_blank" rel="noreferrer">{SOC.li}</a>
          </div>
        </div>
        <div className="ftr-nav">
          {NAV.map((n) => <button key={n.page} className="ftr-l" onClick={() => go(n.page)}>{n.label}</button>)}
        </div>
      </div>
      <div className="ftr-bot wrap">
        <span>© {new Date().getFullYear()} Delhi Hockey. All rights reserved.</span>
        <button className="admin-t" onClick={onAdmin}>Admin</button>
      </div>
    </footer>
  );
}

// ============================================================
// ADMIN
// ============================================================
function Admin({ onClose, state, setters }) {
  const [authed, setAuthed] = useState(false);
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [busy, setBusy] = useState(false);
  const [tab, setTab] = useState("Documents");
  const [toast, setToast] = useState("");
  const flash = (m) => { setToast(m); setTimeout(() => setToast(""), 2600); };

  // surface storage-save failures (e.g. data too large) while the admin panel is open
  useEffect(() => {
    onSaveError = (msg) => flash(msg);
    return () => { onSaveError = null; };
  }, []);

  // if an admin session already exists (e.g. page was reloaded), restore it
  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      if (data?.session) setAuthed(true);
    });
  }, []);

  const signIn = async () => {
    if (busy) return;
    setBusy(true);
    try {
      const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password: pw });
      if (error) { flash("Wrong email or password"); setBusy(false); return; }
      setAuthed(true);
    } catch {
      flash("Couldn't sign in — check your connection");
    }
    setBusy(false);
  };

  const signOut = async () => {
    await supabase.auth.signOut();
    setAuthed(false);
    onClose();
  };

  if (!authed) {
    return (
      <Overlay onClose={onClose}>
        <h2 className="ah">Admin access</h2>
        <p className="amut">Sign in to manage the website.</p>
        <input className="inp" type="email" placeholder="Email" value={email}
          onChange={(e) => setEmail(e.target.value)} autoComplete="username" />
        <input className="inp" type="password" placeholder="Password" value={pw}
          onChange={(e) => setPw(e.target.value)} autoComplete="current-password"
          onKeyDown={(e) => e.key === "Enter" && signIn()} />
        <button className="btn" style={{ width: "100%", marginTop: 4 }} disabled={busy} onClick={signIn}>
          {busy ? "Signing in…" : "Sign in"}
        </button>
        {toast && <Toast>{toast}</Toast>}
      </Overlay>
    );
  }

  const TABS = ["Documents", "About page", "Calendar", "Partners", "FIH Rules", "Registration", "League"];
  return (
    <Overlay onClose={onClose} wide>
      <div className="arow"><h2 className="ah" style={{ margin: 0 }}>Manage website</h2><div style={{ display: "flex", gap: 8, alignItems: "center" }}><button className="linkbtn" style={{ color: MUTE, marginLeft: 0 }} onClick={signOut}>Sign out</button><button className="x" onClick={onClose}>✕</button></div></div>
      <div className="atabs">{TABS.map((t) => <button key={t} className={"atab" + (tab === t ? " on" : "")} onClick={() => setTab(t)}>{t}</button>)}</div>
      <div style={{ marginTop: 16 }}>
        {tab === "Documents" && <DocManager label="document" items={state.docs} setItems={setters.setDocs} storageKey="dh_docs" flash={flash} requireFile />}
        {tab === "Calendar" && <DocManager label="calendar file" items={state.calendarItems} setItems={setters.setCalendarItems} storageKey="dh_calendar" flash={flash} requireFile />}
        {tab === "About page" && <AboutManager intro={state.aboutIntro} setIntro={setters.setAboutIntro} items={state.aboutItems} setItems={setters.setAboutItems} flash={flash} />}
        {tab === "Partners" && <PartnerManager items={state.partners} setItems={setters.setPartners} flash={flash} />}
        {tab === "FIH Rules" && <FihManager items={state.fihItems} setItems={setters.setFihItems} flash={flash} />}
        {tab === "Registration" && <RegManager reg={state.reg} setReg={setters.setReg} flash={flash} />}
        {tab === "League" && <LeagueManager leagues={state.leagues} setLeagues={setters.setLeagues} globalSponsors={state.globalSponsors} setGlobalSponsors={setters.setGlobalSponsors} players={state.players} setPlayers={setters.setPlayers} flash={flash} />}
      </div>
      {toast && <Toast>{toast}</Toast>}
    </Overlay>
  );
}

// Safely read a file to a data URL, returning {ok, dataUrl} or {ok:false, msg}.
async function readFileSafe(file) {
  try {
    const dataUrl = await fileToDataURL(file);
    return { ok: true, dataUrl };
  } catch (e) {
    if (e && e.message === "FILE_TOO_LARGE") return { ok: false, msg: "File is too large — keep it under 3.5 MB." };
    return { ok: false, msg: "Could not read that file. Please try another." };
  }
}

function DocManager({ label, items, setItems, storageKey, flash, requireFile }) {
  const [title, setTitle] = useState("");
  const [date, setDate] = useState(new Date().toISOString().slice(0, 10));
  const [file, setFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const fr = useRef(null);
  const add = async () => {
    if (!title.trim()) return flash("Add a title");
    if (requireFile && !file) return flash("Choose a file");
    setBusy(true);
    let dataUrl = "", fileName = "", kind = "File";
    if (file) {
      const r = await readFileSafe(file);
      if (!r.ok) { setBusy(false); return flash(r.msg); }
      dataUrl = r.dataUrl; fileName = file.name; kind = fileKind(file.name, file.type);
    }
    const next = [{ id: uid(), title: title.trim(), date, dataUrl, fileName, kind }, ...items];
    setItems(next);
    const ok = await saveJSON(storageKey, next);
    setTitle(""); setFile(null); if (fr.current) fr.current.value = ""; setBusy(false);
    flash(ok ? "Published" : "Saved locally, but storage may be full");
  };
  const remove = async (id) => { const n = items.filter((x) => x.id !== id); setItems(n); await saveJSON(storageKey, n); flash("Deleted"); };
  return (
    <div>
      <div className="fcard">
        <div className="fct">Upload a new {label}</div>
        <input className="inp" placeholder="Title" value={title} onChange={(e) => setTitle(e.target.value)} />
        <div className="r2">
          <label className="fld"><span>Publication date</span><input className="inp" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></label>
          <label className="fld"><span>File (PDF, Excel, Word)</span><input ref={fr} className="inp" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.csv" onChange={(e) => setFile(e.target.files[0])} /></label>
        </div>
        <button className="btn" onClick={add} disabled={busy}>{busy ? "Uploading…" : "Publish"}</button>
        <p className="hint">Keep files under ~4 MB so they save reliably.</p>
      </div>
      <ManageList items={items} onRemove={remove} />
    </div>
  );
}

function AboutManager({ intro, setIntro, items, setItems, flash }) {
  const [title, setTitle] = useState("");
  const [file, setFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const fr = useRef(null);
  const saveIntro = async () => { await saveJSON("dh_about_intro", intro); flash("Intro saved"); };
  const add = async () => {
    if (!title.trim()) return flash("Add a heading");
    setBusy(true);
    let dataUrl = "", fileName = "", kind = "File";
    if (file) {
      const r = await readFileSafe(file);
      if (!r.ok) { setBusy(false); return flash(r.msg); }
      dataUrl = r.dataUrl; fileName = file.name; kind = fileKind(file.name, file.type);
    }
    const next = [...items, { id: uid(), title: title.trim(), date: new Date().toISOString().slice(0, 10), dataUrl, fileName, kind }];
    setItems(next); await saveJSON("dh_about", next);
    setTitle(""); setFile(null); if (fr.current) fr.current.value = ""; setBusy(false); flash("Added");
  };
  const remove = async (id) => { const n = items.filter((x) => x.id !== id); setItems(n); await saveJSON("dh_about", n); flash("Deleted"); };
  const move = async (i, dir) => { const j = i + dir; if (j < 0 || j >= items.length) return; const n = [...items]; [n[i], n[j]] = [n[j], n[i]]; setItems(n); await saveJSON("dh_about", n); };
  return (
    <div>
      <div className="fcard">
        <div className="fct">Page introduction</div>
        <textarea className="inp" rows={3} value={intro} onChange={(e) => setIntro(e.target.value)} />
        <button className="btn" onClick={saveIntro}>Save intro</button>
      </div>
      <div className="fcard">
        <div className="fct">Add a heading & document</div>
        <input className="inp" placeholder="Heading (e.g. The Selection Committee)" value={title} onChange={(e) => setTitle(e.target.value)} />
        <label className="fld"><span>Document (optional)</span><input ref={fr} className="inp" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.csv" onChange={(e) => setFile(e.target.files[0])} /></label>
        <button className="btn" onClick={add} disabled={busy}>{busy ? "Saving…" : "Add to list"}</button>
      </div>
      <p className="amut" style={{ fontSize: 12.5 }}>Use the arrows to reorder — this is the order visitors see.</p>
      <div className="mlist">
        {items.map((d, i) => (
          <div className="mrow" key={d.id}>
            <div className="obtns"><button className="ib" onClick={() => move(i, -1)} disabled={i === 0}>{I.up()}</button><button className="ib" onClick={() => move(i, 1)} disabled={i === items.length - 1}>{I.down()}</button></div>
            <span className="ordn">{i + 1}</span>
            <div style={{ flex: 1, minWidth: 0 }}><div className="mt">{d.title}</div><div className="ms">{d.dataUrl ? `${d.kind} attached` : "No file"}</div></div>
            <button className="ib danger" onClick={() => remove(d.id)}>{I.trash()}</button>
          </div>
        ))}
      </div>
    </div>
  );
}

function PartnerManager({ items, setItems, flash }) {
  const [name, setName] = useState(""); const [url, setUrl] = useState(""); const [logo, setLogo] = useState(null);
  const [busy, setBusy] = useState(false); const fr = useRef(null);
  const add = async () => {
    if (!name.trim()) return flash("Add a name"); setBusy(true);
    let logoData = "";
    if (logo) {
      const r = await readFileSafe(logo);
      if (!r.ok) { setBusy(false); return flash(r.msg); }
      logoData = r.dataUrl;
    }
    const next = [...items, { id: uid(), name: name.trim(), url: url.trim(), logo: logoData }];
    setItems(next); await saveJSON("dh_partners", next);
    setName(""); setUrl(""); setLogo(null); if (fr.current) fr.current.value = ""; setBusy(false); flash("Partner added");
  };
  const remove = async (id) => { const n = items.filter((x) => x.id !== id); setItems(n); await saveJSON("dh_partners", n); flash("Removed"); };
  return (
    <div>
      <div className="fcard">
        <div className="fct">Add a partner / sponsor</div>
        <input className="inp" placeholder="Partner name" value={name} onChange={(e) => setName(e.target.value)} />
        <input className="inp" placeholder="Website link (https://…)" value={url} onChange={(e) => setUrl(e.target.value)} />
        <label className="fld"><span>Logo (PNG/JPG)</span><input ref={fr} className="inp" type="file" accept="image/*" onChange={(e) => setLogo(e.target.files[0])} /></label>
        <button className="btn" onClick={add} disabled={busy}>{busy ? "Saving…" : "Add partner"}</button>
      </div>
      <div className="mlist">
        {items.map((p) => (
          <div className="mrow" key={p.id}>
            {p.logo ? <img src={p.logo} alt="" style={{ width: 34, height: 34, objectFit: "contain", borderRadius: 6 }} /> : <div className="pfall sm">{p.name[0]}</div>}
            <div style={{ flex: 1, minWidth: 0 }}><div className="mt">{p.name}</div><div className="ms">{p.url || "No link"}</div></div>
            <button className="ib danger" onClick={() => remove(p.id)}>{I.trash()}</button>
          </div>
        ))}
      </div>
    </div>
  );
}

function LeagueManager({ leagues, setLeagues, globalSponsors, setGlobalSponsors, players, setPlayers, flash }) {
  const [sub, setSub] = useState("Player Registry");
  const SUBS = ["Player Registry", "Leagues & Seasons", "Sponsors", "Teams & Squads", "Schedule", "Results", "Export Data"];

  // selected league + season for the editing tabs
  const [leagueId, setLeagueId] = useState(leagues[0]?.id || "");
  const league = leagues.find((l) => l.id === leagueId) || leagues[0];
  const seasons = league ? [...(league.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year)) : [];
  const [seasonId, setSeasonId] = useState(seasons[0]?.id || "");
  const season = seasons.find((s) => s.id === seasonId) || seasons[0];

  const persist = async (next) => { setLeagues(next); await saveJSON("dh_leagues", next); };

  // writes teams/matches into the selected season.
  // Uses functional setState so it always reads the latest leagues (avoids stale-closure overwrites).
  const updateSeason = (patch) => {
    setLeagues((prev) => {
      const next = prev.map((l) => l.id !== (league?.id) ? l : {
        ...l, seasons: (l.seasons || []).map((s) => s.id !== (season?.id) ? s : { ...s, ...patch }),
      });
      saveJSON("dh_leagues", next);
      return next;
    });
  };
  // setter shims so the existing managers work on the selected season
  const teams = season?.teams || [];
  const matches = season?.matches || [];
  const setTeams = (updater) => {
    const value = typeof updater === "function" ? updater(teams) : updater;
    updateSeason({ teams: value });
  };
  const setMatches = (updater) => {
    const value = typeof updater === "function" ? updater(matches) : updater;
    updateSeason({ matches: value });
  };

  const needSeason = ["Teams & Squads", "Schedule", "Results"].includes(sub);

  return (
    <div>
      <div className="lsubs">
        {SUBS.map((s) => <button key={s} className={"lsub" + (sub === s ? " on" : "")} onClick={() => setSub(s)}>{s}</button>)}
      </div>

      {needSeason && (
        leagues.length === 0 ? (
          <div className="fcard" style={{ marginTop: 14 }}><p className="amut" style={{ margin: 0, fontSize: 13.5 }}>Create a league first in <b>Leagues & Seasons</b>.</p></div>
        ) : (
          <div className="ctx-bar">
            <label className="lg-sel"><span>Editing league</span>
              <select className="inp" value={league?.id || ""} onChange={(e) => { setLeagueId(e.target.value); const l = leagues.find((x) => x.id === e.target.value); const ss = [...(l?.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year)); setSeasonId(ss[0]?.id || ""); }}>
                {leagues.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
              </select>
            </label>
            <label className="lg-sel"><span>Season</span>
              <select className="inp" value={season?.id || ""} onChange={(e) => setSeasonId(e.target.value)} disabled={seasons.length === 0}>
                {seasons.length === 0 ? <option>No seasons</option> : seasons.map((s) => <option key={s.id} value={s.id}>{s.year}</option>)}
              </select>
            </label>
          </div>
        )
      )}

      <div style={{ marginTop: 14 }}>
        {sub === "Player Registry" && <RegistryManager players={players} setPlayers={setPlayers} leagues={leagues} setLeagues={setLeagues} flash={flash} />}
        {sub === "Leagues & Seasons" && <LeagueSeasonManager leagues={leagues} persist={persist} flash={flash} />}
        {sub === "Sponsors" && <SponsorManager leagues={leagues} persist={persist} globalSponsors={globalSponsors} setGlobalSponsors={setGlobalSponsors} flash={flash} />}
        {needSeason && leagues.length > 0 && !season && <div className="fcard"><p className="amut" style={{ margin: 0, fontSize: 13.5 }}>This league has no season yet. Add one in <b>Leagues & Seasons</b>.</p></div>}
        {sub === "Teams & Squads" && season && <TeamManager teams={teams} setTeams={setTeams} matches={matches} setMatches={setMatches} players={players} flash={flash} />}
        {sub === "Schedule" && season && <ScheduleManager teams={teams} matches={matches} setMatches={setMatches} flash={flash} />}
        {sub === "Results" && season && <ResultsManager teams={teams} matches={matches} setMatches={setMatches} players={players} flash={flash} />}
        {sub === "Export Data" && <ExportManager leagues={leagues} players={players} globalSponsors={globalSponsors} league={league} flash={flash} />}
      </div>
    </div>
  );
}

function ExportManager({ leagues, players, globalSponsors, league, flash }) {
  const stamp = () => new Date().toISOString().slice(0, 10);

  // independent selectors for export (don't disturb the editing tabs)
  const [lid, setLid] = useState(league?.id || leagues[0]?.id || "");
  const lg = leagues.find((l) => l.id === lid) || leagues[0];
  const seasons = lg ? [...(lg.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year)) : [];
  const [sid, setSid] = useState("");
  const sel = seasons.find((s) => s.id === sid) || seasons[0];

  const onLeague = (id) => {
    setLid(id);
    const l = leagues.find((x) => x.id === id);
    const ss = [...(l?.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year));
    setSid(ss[0]?.id || "");
  };

  const exportSeason = () => {
    if (!lg || !sel) return flash("Pick a league and season");
    const sheets = buildExportRows(leagues, players, globalSponsors, { leagueId: lg.id, seasonId: sel.id });
    const safe = (lg.name + "_" + sel.year).replace(/[^\w-]+/g, "_");
    downloadWorkbook(sheets, `DelhiHockey_${safe}_${stamp()}.xlsx`);
    flash("Exporting season…");
  };
  const exportRegistry = () => {
    const sheets = { "Player Registry": (players || []).map((p) => ({ "Unique ID": p.id, "Name": p.name })) };
    downloadWorkbook(sheets, `DelhiHockey_PlayerRegistry_${stamp()}.xlsx`);
    flash("Exporting registry…");
  };

  return (
    <div>
      <div className="fcard">
        <div className="fct">Export a league season</div>
        <p className="hint" style={{ marginTop: 0, marginBottom: 12 }}>
          Choose a league and season to download. The Excel file has separate sheets for squads, coaches, schedule & results, individual match events, referees, standings, player stats and sponsors — all for the selected season.
        </p>
        {leagues.length === 0 ? (
          <p className="amut" style={{ fontSize: 13, margin: 0 }}>No leagues yet. Create one in <b>Leagues & Seasons</b>.</p>
        ) : (
          <>
            <div className="r2">
              <label className="fld"><span>League</span>
                <select className="inp" value={lg?.id || ""} onChange={(e) => onLeague(e.target.value)}>
                  {leagues.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
                </select>
              </label>
              <label className="fld"><span>Season</span>
                <select className="inp" value={sel?.id || ""} onChange={(e) => setSid(e.target.value)} disabled={seasons.length === 0}>
                  {seasons.length === 0 ? <option>No seasons</option> : seasons.map((s) => <option key={s.id} value={s.id}>{s.year}</option>)}
                </select>
              </label>
            </div>
            <button className="btn" onClick={exportSeason} disabled={!lg || !sel}>Download {sel ? `${lg.name} — ${sel.year}` : "season"} (.xlsx)</button>
            {seasons.length === 0 && <p className="hint">This league has no seasons yet.</p>}
          </>
        )}
      </div>

      <div className="fcard">
        <div className="fct">Export player registry</div>
        <p className="hint" style={{ marginTop: 0, marginBottom: 12 }}>The full list of registered players and their Unique IDs.</p>
        <button className="btn ghost" onClick={exportRegistry} disabled={players.length === 0}>Download registry (.xlsx)</button>
      </div>
    </div>
  );
}

function LeagueSeasonManager({ leagues, persist, flash }) {
  const [newLeague, setNewLeague] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editName, setEditName] = useState("");
  const [yearInputs, setYearInputs] = useState({});

  const addLeague = async () => {
    if (!newLeague.trim()) return flash("Enter a league name");
    const next = [...leagues, { id: uid(), name: newLeague.trim(), sponsors: [], seasons: [] }];
    await persist(next); setNewLeague(""); flash("League created");
  };
  const renameLeague = async (id) => {
    if (!editName.trim()) return flash("Enter a name");
    const next = leagues.map((l) => l.id === id ? { ...l, name: editName.trim() } : l);
    await persist(next); setEditingId(null); flash("Renamed");
  };
  const removeLeague = async (id) => {
    if (!window.confirm("Delete this league and all its seasons, teams and matches? This cannot be undone.")) return;
    await persist(leagues.filter((l) => l.id !== id)); flash("League deleted");
  };
  const addSeason = async (leagueId) => {
    const yr = (yearInputs[leagueId] || "").trim();
    if (!/^\d{4}$/.test(yr)) return flash("Enter a 4-digit year");
    const lg = leagues.find((l) => l.id === leagueId);
    if ((lg.seasons || []).some((s) => s.year === yr)) return flash("That season already exists");
    const next = leagues.map((l) => l.id === leagueId ? { ...l, seasons: [...(l.seasons || []), { id: uid(), year: yr, teams: [], matches: [], sponsors: [] }] } : l);
    await persist(next); setYearInputs({ ...yearInputs, [leagueId]: "" }); flash("Season added");
  };
  const removeSeason = async (leagueId, seasonId) => {
    if (!window.confirm("Delete this season with all its teams and matches?")) return;
    const next = leagues.map((l) => l.id === leagueId ? { ...l, seasons: l.seasons.filter((s) => s.id !== seasonId) } : l);
    await persist(next); flash("Season removed");
  };

  return (
    <div>
      <div className="fcard">
        <div className="fct">Create a league</div>
        <div className="r2">
          <input className="inp" placeholder="e.g. Senior Men's League" value={newLeague} onChange={(e) => setNewLeague(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addLeague()} />
          <button className="btn" onClick={addLeague}>Add league</button>
        </div>
      </div>

      {leagues.length === 0 && <p className="amut" style={{ fontSize: 13 }}>No leagues yet.</p>}
      {leagues.map((l) => {
        const ss = [...(l.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year));
        return (
          <div className="fcard" key={l.id}>
            <div className="mrow" style={{ border: "none", background: "transparent", padding: 0, marginBottom: 10 }}>
              {editingId === l.id ? (
                <>
                  <input className="inp" style={{ margin: 0, flex: 1 }} value={editName} onChange={(e) => setEditName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && renameLeague(l.id)} />
                  <button className="btn" style={{ padding: "8px 12px", fontSize: 13 }} onClick={() => renameLeague(l.id)}>Save</button>
                </>
              ) : (
                <>
                  <div style={{ flex: 1, minWidth: 0 }}><div className="mt" style={{ fontSize: 15 }}>{l.name}</div><div className="ms">{ss.length} season{ss.length !== 1 ? "s" : ""}</div></div>
                  <button className="linkbtn" style={{ color: MUTE }} onClick={() => { setEditingId(l.id); setEditName(l.name); }}>rename</button>
                  <button className="ib danger" onClick={() => removeLeague(l.id)}>{I.trash()}</button>
                </>
              )}
            </div>
            <div className="season-add">
              <input className="inp" style={{ margin: 0 }} placeholder="Add season year (e.g. 2025)" value={yearInputs[l.id] || ""} onChange={(e) => setYearInputs({ ...yearInputs, [l.id]: e.target.value })} onKeyDown={(e) => e.key === "Enter" && addSeason(l.id)} />
              <button className="btn ghost" onClick={() => addSeason(l.id)}>Add season</button>
            </div>
            <div className="chip-row" style={{ marginTop: 10 }}>
              {ss.length === 0 && <span className="amut" style={{ fontSize: 12.5 }}>No seasons yet.</span>}
              {ss.map((s) => <span className="chip" key={s.id}>{s.year}<button onClick={() => removeSeason(l.id, s.id)}>✕</button></span>)}
            </div>
          </div>
        );
      })}
    </div>
  );
}

function SponsorManager({ leagues, persist, globalSponsors, setGlobalSponsors, flash }) {
  const [scopeLeague, setScopeLeague] = useState("global"); // 'global' or league id
  const [scopeSeason, setScopeSeason] = useState(""); // season id when a league is chosen

  const isGlobal = scopeLeague === "global";
  const lg = isGlobal ? null : leagues.find((l) => l.id === scopeLeague);
  const seasons = lg ? [...(lg.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year)) : [];
  const season = seasons.find((s) => s.id === scopeSeason) || (isGlobal ? null : seasons[0]);
  const list = isGlobal ? globalSponsors : (season?.sponsors || []);

  const [name, setName] = useState("");
  const [title, setTitle] = useState("");
  const [url, setUrl] = useState("");
  const [logo, setLogo] = useState(null);
  const [busy, setBusy] = useState(false);
  const fr = useRef(null);

  const onLeagueScope = (val) => {
    setScopeLeague(val);
    if (val === "global") { setScopeSeason(""); return; }
    const l = leagues.find((x) => x.id === val);
    const ss = [...(l?.seasons || [])].sort((a, b) => Number(b.year) - Number(a.year));
    setScopeSeason(ss[0]?.id || "");
  };

  const saveList = async (next) => {
    if (isGlobal) { setGlobalSponsors(next); await saveJSON("dh_global_sponsors", next); }
    else {
      if (!season) return flash("Pick a season first");
      await persist(leagues.map((l) => l.id !== scopeLeague ? l : {
        ...l, seasons: (l.seasons || []).map((s) => s.id !== season.id ? s : { ...s, sponsors: next }),
      }));
    }
  };
  const add = async () => {
    if (!name.trim()) return flash("Enter a sponsor name");
    if (!isGlobal && !season) return flash("Pick a season first");
    setBusy(true);
    let logoData = "";
    if (logo) {
      const r = await readFileSafe(logo);
      if (!r.ok) { setBusy(false); return flash(r.msg); }
      logoData = r.dataUrl;
    }
    await saveList([...list, { id: uid(), name: name.trim(), title: title.trim(), url: url.trim(), logo: logoData }]);
    setName(""); setTitle(""); setUrl(""); setLogo(null); if (fr.current) fr.current.value = ""; setBusy(false); flash("Sponsor added");
  };
  const remove = async (id) => { await saveList(list.filter((s) => s.id !== id)); flash("Removed"); };

  return (
    <div>
      <div className="fcard">
        <div className="fct">Sponsor scope</div>
        <div className="r2">
          <label className="fld"><span>Show on…</span>
            <select className="inp" value={scopeLeague} onChange={(e) => onLeagueScope(e.target.value)}>
              <option value="global">All leagues (global)</option>
              {leagues.map((l) => <option key={l.id} value={l.id}>{l.name}</option>)}
            </select>
          </label>
          {!isGlobal && (
            <label className="fld"><span>Season</span>
              <select className="inp" value={season?.id || ""} onChange={(e) => setScopeSeason(e.target.value)} disabled={seasons.length === 0}>
                {seasons.length === 0 ? <option>No seasons</option> : seasons.map((s) => <option key={s.id} value={s.id}>{s.year}</option>)}
              </select>
            </label>
          )}
        </div>
        <p className="hint" style={{ marginTop: 0 }}>
          {isGlobal
            ? "Global sponsors appear on every league and season, until you remove them."
            : "These sponsors appear only for the selected season, alongside any global sponsors. Removing them won't affect other seasons."}
        </p>
        {!isGlobal && seasons.length === 0 && <p className="amut" style={{ fontSize: 12.5, margin: 0 }}>This league has no seasons — add one in <b>Leagues & Seasons</b>.</p>}
      </div>

      <div className="fcard">
        <div className="fct">Add a sponsor {isGlobal ? "(global)" : season ? `to ${lg.name} — ${season.year}` : ""}</div>
        <input className="inp" placeholder="Sponsor name" value={name} onChange={(e) => setName(e.target.value)} />
        <input className="inp" placeholder="Title (e.g. Title Sponsor, Official Partner)" value={title} onChange={(e) => setTitle(e.target.value)} />
        <input className="inp" placeholder="Website link (optional)" value={url} onChange={(e) => setUrl(e.target.value)} />
        <label className="fld"><span>Logo (PNG/JPG)</span><input ref={fr} className="inp" type="file" accept="image/*" onChange={(e) => setLogo(e.target.files[0])} /></label>
        <button className="btn" onClick={add} disabled={busy || (!isGlobal && !season)}>{busy ? "Saving…" : "Add sponsor"}</button>
      </div>

      <div className="mlist">
        {list.length === 0 && <p className="amut" style={{ fontSize: 13 }}>No sponsors in this scope yet.</p>}
        {list.map((s) => (
          <div className="mrow" key={s.id}>
            {s.logo ? <img src={s.logo} alt="" style={{ width: 36, height: 36, objectFit: "contain", borderRadius: 6 }} /> : <div className="pfall sm">{s.name[0]}</div>}
            <div style={{ flex: 1, minWidth: 0 }}><div className="mt">{s.name}</div><div className="ms">{s.title || "—"}</div></div>
            <button className="ib danger" onClick={() => remove(s.id)}>{I.trash()}</button>
          </div>
        ))}
      </div>
    </div>
  );
}

function RegistryManager({ players, setPlayers, leagues, setLeagues, flash }) {
  const [busy, setBusy] = useState(false);
  const [search, setSearch] = useState("");
  const [mid, setMid] = useState("");
  const [mname, setMname] = useState("");
  const fr = useRef(null);

  const importFile = async (file) => {
    if (!file) return;
    setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
      // find the ID column and name column by header keywords
      let headerRow = 0;
      const norm = (v) => String(v || "").trim().toLowerCase();
      let idCol = -1, nameCol = -1;
      for (let r = 0; r < Math.min(rows.length, 5); r++) {
        const cells = rows[r].map(norm);
        const ic = cells.findIndex((c) => c.includes("id") || c.includes("unique"));
        const nc = cells.findIndex((c) => c.includes("name"));
        if (ic !== -1 && nc !== -1) { headerRow = r; idCol = ic; nameCol = nc; break; }
      }
      if (idCol === -1) {
        // assume col 0 = ID, col 1 = name, no header
        idCol = 0; nameCol = 1; headerRow = -1;
      }
      const found = [];
      const seen = new Set(players.map((p) => p.id));
      for (let r = headerRow + 1; r < rows.length; r++) {
        const row = rows[r]; if (!row) continue;
        const id = String(row[idCol] || "").trim();
        const name = String(row[nameCol] || "").trim();
        if (!id || !name) continue;
        if (seen.has(id)) continue;
        seen.add(id);
        found.push({ id, name });
      }
      if (found.length === 0) { flash("No new players found — check the columns"); setBusy(false); return; }
      const next = [...players, ...found];
      setPlayers(next); await saveJSON("dh_players", next);
      flash(`Imported ${found.length} player${found.length > 1 ? "s" : ""}`);
    } catch (e) {
      console.error(e); flash("Could not read that file");
    }
    if (fr.current) fr.current.value = "";
    setBusy(false);
  };

  const addManual = async () => {
    if (!mid.trim() || !mname.trim()) return flash("Enter both ID and name");
    if (players.some((p) => p.id === mid.trim())) return flash("That ID already exists");
    const next = [...players, { id: mid.trim(), name: mname.trim() }];
    setPlayers(next); await saveJSON("dh_players", next);
    setMid(""); setMname(""); flash("Player added");
  };

  const remove = async (id) => {
    const next = players.filter((p) => p.id !== id);
    setPlayers(next); await saveJSON("dh_players", next);
    // pull them from every squad in every league + season
    const lg2 = (leagues || []).map((l) => ({
      ...l,
      seasons: (l.seasons || []).map((s) => ({
        ...s,
        teams: (s.teams || []).map((t) => ({ ...t, players: (t.players || []).filter((pid) => pid !== id) })),
      })),
    }));
    setLeagues(lg2); await saveJSON("dh_leagues", lg2);
    flash("Removed from registry");
  };

  const clearAll = async () => {
    if (!window.confirm("Clear the entire player registry? This cannot be undone.")) return;
    setPlayers([]); await saveJSON("dh_players", []);
    flash("Registry cleared");
  };

  const filtered = search.trim()
    ? players.filter((p) => p.id.toLowerCase().includes(search.toLowerCase()) || p.name.toLowerCase().includes(search.toLowerCase()))
    : players;

  return (
    <div>
      <div className="fcard">
        <div className="fct">Import players from spreadsheet</div>
        <p className="hint" style={{ marginTop: 0, marginBottom: 10 }}>
          Upload an Excel (.xlsx) or CSV with two columns: <b>Unique ID</b> (e.g. DH-ATH-15082006-004213) and <b>Name</b>. A header row is detected automatically. Re-importing skips IDs already in the registry.
        </p>
        <input ref={fr} className="inp" type="file" accept=".xlsx,.xls,.csv" disabled={busy} onChange={(e) => importFile(e.target.files[0])} />
        {busy && <p className="hint">Reading file…</p>}
      </div>

      <div className="fcard">
        <div className="fct">Or add one player manually</div>
        <div className="r2">
          <input className="inp" placeholder="DH-ATH-DDMMYYYY-XXXXXX" value={mid} onChange={(e) => setMid(e.target.value)} />
          <input className="inp" placeholder="Player name" value={mname} onChange={(e) => setMname(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addManual()} />
        </div>
        <button className="btn" onClick={addManual}>Add player</button>
      </div>

      <div className="reg-head">
        <span className="fct" style={{ margin: 0 }}>Registry — {players.length} player{players.length !== 1 ? "s" : ""}</span>
        {players.length > 0 && <button className="linkbtn" onClick={clearAll}>clear all</button>}
      </div>
      {players.length > 0 && (
        <input className="inp" placeholder="Search by name or ID…" value={search} onChange={(e) => setSearch(e.target.value)} />
      )}
      <div className="mlist" style={{ maxHeight: 320, overflowY: "auto" }}>
        {players.length === 0 && <p className="amut" style={{ fontSize: 13 }}>Registry is empty. Import a spreadsheet to begin.</p>}
        {filtered.map((p) => (
          <div className="mrow" key={p.id}>
            <div style={{ flex: 1, minWidth: 0 }}><div className="mt">{p.name}</div><div className="ms mono">{p.id}</div></div>
            <button className="ib danger" onClick={() => remove(p.id)}>{I.trash()}</button>
          </div>
        ))}
        {players.length > 0 && filtered.length === 0 && <p className="amut" style={{ fontSize: 13 }}>No matches.</p>}
      </div>
    </div>
  );
}

function TeamManager({ teams, setTeams, matches, players, flash }) {
  const [name, setName] = useState("");
  const [expanded, setExpanded] = useState(null);
  const [pick, setPick] = useState(""); // search text for player picker
  const [coachName, setCoachName] = useState("");
  const [coachId, setCoachId] = useState("");

  const addTeam = async () => {
    if (!name.trim()) return flash("Add a team name");
    const next = [...teams, { id: uid(), name: name.trim(), players: [], coaches: [] }];
    setTeams(next); setName(""); flash("Team added");
  };
  const removeTeam = async (id) => {
    if (matches.some((m) => m.home === id || m.away === id)) return flash("Remove this team's matches first");
    const next = teams.filter((t) => t.id !== id);
    setTeams(next); flash("Team removed");
  };
  const addPlayer = async (teamId, pid) => {
    const team = teams.find((t) => t.id === teamId);
    if ((team.players || []).includes(pid)) return flash("Already in this squad");
    const next = teams.map((t) => t.id === teamId ? { ...t, players: [...(t.players || []), pid] } : t);
    setTeams(next); setPick("");
  };
  const removePlayer = async (teamId, pid) => {
    const next = teams.map((t) => t.id === teamId ? { ...t, players: (t.players || []).filter((x) => x !== pid) } : t);
    setTeams(next);
  };
  const addCoach = async (teamId) => {
    if (!coachName.trim()) return flash("Enter the coach's name");
    const coach = { cid: uid(), name: coachName.trim(), regId: coachId.trim() };
    const next = teams.map((t) => t.id === teamId ? { ...t, coaches: [...(t.coaches || []), coach] } : t);
    setTeams(next); setCoachName(""); setCoachId(""); flash("Coach added");
  };
  const removeCoach = async (teamId, cid) => {
    const next = teams.map((t) => t.id === teamId ? { ...t, coaches: (t.coaches || []).filter((c) => c.cid !== cid) } : t);
    setTeams(next);
  };

  return (
    <div>
      <div className="fcard">
        <div className="fct">Add a team</div>
        <div className="r2">
          <input className="inp" placeholder="Team name" value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addTeam()} />
          <button className="btn" onClick={addTeam}>Add team</button>
        </div>
      </div>
      {players.length === 0 && (
        <div className="fcard"><p className="amut" style={{ margin: 0, fontSize: 13 }}>Import players in the <b>Player Registry</b> tab first — squads are built from registered Unique IDs.</p></div>
      )}
      <div className="mlist">
        {teams.length === 0 && <p className="amut" style={{ fontSize: 13 }}>No teams yet.</p>}
        {teams.map((t) => {
          const squadIds = t.players || [];
          const coaches = t.coaches || [];
          const available = players.filter((p) => !squadIds.includes(p.id) &&
            (pick.trim() ? (p.id.toLowerCase().includes(pick.toLowerCase()) || p.name.toLowerCase().includes(pick.toLowerCase())) : true));
          return (
            <div key={t.id} className="team-block">
              <div className="mrow" style={{ border: "none", background: "transparent", padding: "4px 0" }}>
                <button className="ib" onClick={() => { setExpanded(expanded === t.id ? null : t.id); setPick(""); setCoachName(""); setCoachId(""); }} title="Squad">{expanded === t.id ? I.down() : I.arrow()}</button>
                <div style={{ flex: 1, minWidth: 0 }}><div className="mt">{t.name}</div><div className="ms">{squadIds.length} players{coaches.length ? ` · ${coaches.length} coach${coaches.length > 1 ? "es" : ""}` : ""}</div></div>
                <button className="ib danger" onClick={() => removeTeam(t.id)}>{I.trash()}</button>
              </div>
              {expanded === t.id && (
                <div className="squad-edit">
                  <div className="sq-section">Coaching staff</div>
                  <div className="r3">
                    <input className="inp" placeholder="Coach name" value={coachName} onChange={(e) => setCoachName(e.target.value)} />
                    <input className="inp" placeholder="Unique ID (optional)" value={coachId} onChange={(e) => setCoachId(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addCoach(t.id)} />
                    <button className="btn" onClick={() => addCoach(t.id)}>Add coach</button>
                  </div>
                  <ol className="squad-ol">
                    {coaches.length === 0 && <li><span className="amut" style={{ fontSize: 12.5 }}>No coaches added.</span></li>}
                    {coaches.map((c) => (
                      <li key={c.cid}><span>{c.name}</span>{c.regId && <span className="mono sq-id">{c.regId}</span>}<button className="linkbtn" onClick={() => removeCoach(t.id, c.cid)}>remove</button></li>
                    ))}
                  </ol>

                  <div className="sq-section" style={{ marginTop: 16 }}>Players</div>
                  {players.length > 0 && (
                    <>
                      <input className="inp" placeholder="Search registry by name or ID to add…" value={pick} onChange={(e) => setPick(e.target.value)} />
                      {pick.trim() && (
                        <div className="pick-list">
                          {available.length === 0 ? <p className="amut" style={{ fontSize: 12.5, padding: "4px 2px" }}>No matches.</p> :
                            available.slice(0, 8).map((p) => (
                              <button key={p.id} className="pick-row" onClick={() => addPlayer(t.id, p.id)}>
                                <span>{p.name}</span><span className="mono">{p.id}</span>
                              </button>
                            ))}
                        </div>
                      )}
                    </>
                  )}
                  <ol className="squad-ol">
                    {squadIds.length === 0 && <li><span className="amut" style={{ fontSize: 12.5 }}>No players added.</span></li>}
                    {squadIds.map((pid) => (
                      <li key={pid}><span>{playerName(players, pid)}</span><span className="mono sq-id">{pid}</span><button className="linkbtn" onClick={() => removePlayer(t.id, pid)}>remove</button></li>
                    ))}
                  </ol>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

function ScheduleManager({ teams, matches, setMatches, flash }) {
  const [home, setHome] = useState("");
  const [away, setAway] = useState("");
  const [date, setDate] = useState("");
  const [time, setTime] = useState("");
  const [venue, setVenue] = useState("");
  const [refInput, setRefInput] = useState("");
  const [refs, setRefs] = useState([]);

  const addRef = () => {
    if (!refInput.trim()) return;
    setRefs([...refs, refInput.trim()]); setRefInput("");
  };
  const removeRef = (i) => setRefs(refs.filter((_, x) => x !== i));

  const add = async () => {
    if (!home || !away) return flash("Pick both teams");
    if (home === away) return flash("Teams must differ");
    const next = [...matches, { id: uid(), home, away, date, time, venue: venue.trim(), referees: refs, status: "scheduled", homeGoals: 0, awayGoals: 0, events: [], homePC: 0, awayPC: 0, homeCE: 0, awayCE: 0, pom: "" }];
    setMatches(next);
    setHome(""); setAway(""); setDate(""); setTime(""); setVenue(""); setRefs([]); setRefInput(""); flash("Match scheduled");
  };
  const remove = async (id) => { const n = matches.filter((m) => m.id !== id); setMatches(n); flash("Removed"); };

  if (teams.length < 2) return <div className="fcard"><p className="amut" style={{ margin: 0, fontSize: 13.5 }}>Add at least two teams first (Teams & Squads tab).</p></div>;
  const sorted = [...matches].sort((a, b) => dateVal(a.date) - dateVal(b.date));
  return (
    <div>
      <div className="fcard">
        <div className="fct">Schedule a match</div>
        <div className="r2">
          <label className="fld"><span>Home team</span><select className="inp" value={home} onChange={(e) => setHome(e.target.value)}><option value="">Select…</option>{teams.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}</select></label>
          <label className="fld"><span>Away team</span><select className="inp" value={away} onChange={(e) => setAway(e.target.value)}><option value="">Select…</option>{teams.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}</select></label>
        </div>
        <div className="r3">
          <label className="fld"><span>Date</span><input className="inp" type="date" value={date} onChange={(e) => setDate(e.target.value)} /></label>
          <label className="fld"><span>Start time</span><input className="inp" type="time" value={time} onChange={(e) => setTime(e.target.value)} /></label>
          <label className="fld"><span>Venue (optional)</span><input className="inp" placeholder="e.g. Shivaji Stadium" value={venue} onChange={(e) => setVenue(e.target.value)} /></label>
        </div>
        <label className="fld"><span>Referees / umpires (optional)</span>
          <div className="r2">
            <input className="inp" placeholder="Referee name" value={refInput} onChange={(e) => setRefInput(e.target.value)} onKeyDown={(e) => e.key === "Enter" && (e.preventDefault(), addRef())} />
            <button className="btn ghost" onClick={addRef}>Add referee</button>
          </div>
        </label>
        {refs.length > 0 && (
          <div className="chip-row">
            {refs.map((r, i) => <span className="chip" key={i}>{r}<button onClick={() => removeRef(i)}>✕</button></span>)}
          </div>
        )}
        <button className="btn" style={{ marginTop: 4 }} onClick={add}>Add to schedule</button>
      </div>
      <div className="mlist">
        {sorted.length === 0 && <p className="amut" style={{ fontSize: 13 }}>No matches scheduled.</p>}
        {sorted.map((m) => (
          <div className="mrow" key={m.id}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div className="mt">{teamName(teams, m.home)} vs {teamName(teams, m.away)}</div>
              <div className="ms">{m.date ? fmtDate(m.date) : "TBD"}{m.time ? ` · ${fmtTime(m.time)}` : ""}{m.venue ? ` · ${m.venue}` : ""}{(m.referees && m.referees.length) ? ` · ${m.referees.length} ref${m.referees.length > 1 ? "s" : ""}` : ""} · {m.status === "played" ? "Result recorded" : "Scheduled"}</div>
            </div>
            <button className="ib danger" onClick={() => remove(m.id)}>{I.trash()}</button>
          </div>
        ))}
      </div>
    </div>
  );
}

function ResultsManager({ teams, matches, setMatches, players, flash }) {
  const [editing, setEditing] = useState(null);
  const sorted = [...matches].sort((a, b) => dateVal(a.date) - dateVal(b.date));
  if (matches.length === 0) return <div className="fcard"><p className="amut" style={{ margin: 0, fontSize: 13.5 }}>Schedule matches first (Schedule tab).</p></div>;
  if (editing) {
    const m = matches.find((x) => x.id === editing);
    if (m) return <ResultEditor m={m} teams={teams} matches={matches} setMatches={setMatches} players={players} flash={flash} onDone={() => setEditing(null)} />;
  }
  return (
    <div className="mlist">
      {sorted.map((m) => (
        <div className="mrow" key={m.id}>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div className="mt">{teamName(teams, m.home)} {m.status === "played" ? `${m.homeGoals}–${m.awayGoals}` : "vs"} {teamName(teams, m.away)}</div>
            <div className="ms">{m.date ? fmtDate(m.date) : "TBD"} · {m.status === "played" ? "Result recorded" : "Not played"}</div>
          </div>
          <button className="btn" style={{ padding: "8px 14px", fontSize: 13 }} onClick={() => setEditing(m.id)}>{m.status === "played" ? "Edit" : "Record"}</button>
        </div>
      ))}
    </div>
  );
}

function ResultEditor({ m, teams, matches, setMatches, players, flash, onDone }) {
  const homeT = teams.find((t) => t.id === m.home);
  const awayT = teams.find((t) => t.id === m.away);
  const homeSquad = homeT?.players || [];
  const awaySquad = awayT?.players || [];
  const hasSquad = homeSquad.length + awaySquad.length > 0;
  const teamOfPid = (pid) => homeSquad.includes(pid) ? m.home : (awaySquad.includes(pid) ? m.away : "");

  const [hg, setHg] = useState(m.homeGoals || 0);
  const [ag, setAg] = useState(m.awayGoals || 0);
  const [hpc, setHpc] = useState(m.homePC || 0);
  const [apc, setApc] = useState(m.awayPC || 0);
  const [hce, setHce] = useState(m.homeCE || 0);
  const [ace, setAce] = useState(m.awayCE || 0);
  const [pom, setPom] = useState(m.pom || "");
  const [events, setEvents] = useState(m.events || []);

  const [evType, setEvType] = useState("goal");
  const [evPlayer, setEvPlayer] = useState(""); // holds a Unique ID
  const [evMin, setEvMin] = useState("");

  const addEvent = () => {
    if (!evPlayer) return flash("Pick a player");
    const minNum = evMin === "" ? null : Math.max(1, Math.min(60, Number(evMin) || 0));
    const next = [...events, { id: uid(), type: evType, player: evPlayer, team: teamOfPid(evPlayer), min: minNum }];
    next.sort((a, b) => (a.min || 999) - (b.min || 999));
    setEvents(next);
    setEvPlayer(""); setEvMin("");
  };
  const removeEvent = (id) => setEvents(events.filter((e) => e.id !== id));

  const save = async (markPlayed) => {
    const next = matches.map((x) => x.id === m.id ? {
      ...x, homeGoals: Number(hg) || 0, awayGoals: Number(ag) || 0,
      homePC: Number(hpc) || 0, awayPC: Number(apc) || 0, homeCE: Number(hce) || 0, awayCE: Number(ace) || 0,
      pom, events, status: markPlayed ? "played" : x.status,
    } : x);
    setMatches(next);
    flash(markPlayed ? "Result saved & published" : "Saved");
    onDone();
  };

  const EV_LABELS = { goal: "Goal", assist: "Assist", green: "Green card", yellow: "Yellow card", red: "Red card" };
  const evLabel = (t) => EV_LABELS[t] || t;
  const squadOptions = (squad, label) => squad.length > 0 && (
    <optgroup label={label}>{squad.map((pid) => <option key={pid} value={pid}>{playerName(players, pid)} — {pid}</option>)}</optgroup>
  );

  return (
    <div>
      <button className="linkbtn" style={{ color: MUTE, marginBottom: 10 }} onClick={onDone}>← Back to matches</button>
      <div className="fcard">
        <div className="fct">{homeT?.name} vs {awayT?.name}</div>
        <div className="score-row">
          <div className="score-side"><span>{homeT?.name}</span><input className="inp num" type="number" min="0" value={hg} onChange={(e) => setHg(e.target.value)} /></div>
          <span className="score-dash">–</span>
          <div className="score-side"><input className="inp num" type="number" min="0" value={ag} onChange={(e) => setAg(e.target.value)} /><span>{awayT?.name}</span></div>
        </div>
      </div>

      <div className="fcard">
        <div className="fct">Penalty corners & circle entries</div>
        <div className="r2">
          <label className="fld"><span>{homeT?.name} — penalty corners</span><input className="inp" type="number" min="0" value={hpc} onChange={(e) => setHpc(e.target.value)} /></label>
          <label className="fld"><span>{awayT?.name} — penalty corners</span><input className="inp" type="number" min="0" value={apc} onChange={(e) => setApc(e.target.value)} /></label>
        </div>
        <div className="r2">
          <label className="fld"><span>{homeT?.name} — circle entries</span><input className="inp" type="number" min="0" value={hce} onChange={(e) => setHce(e.target.value)} /></label>
          <label className="fld"><span>{awayT?.name} — circle entries</span><input className="inp" type="number" min="0" value={ace} onChange={(e) => setAce(e.target.value)} /></label>
        </div>
      </div>

      <div className="fcard">
        <div className="fct">Player events (goals, assists, cards)</div>
        {!hasSquad ? (
          <p className="amut" style={{ fontSize: 13 }}>Build the squads (Teams & Squads tab) to record individual events. You can still save the scoreline above.</p>
        ) : (
          <>
            <div className="r4">
              <select className="inp" value={evType} onChange={(e) => setEvType(e.target.value)}>
                <option value="goal">Goal</option><option value="assist">Assist</option>
                <option value="green">Green card</option><option value="yellow">Yellow card</option><option value="red">Red card</option>
              </select>
              <select className="inp" value={evPlayer} onChange={(e) => setEvPlayer(e.target.value)}>
                <option value="">Select player…</option>
                {squadOptions(homeSquad, homeT.name)}
                {squadOptions(awaySquad, awayT.name)}
              </select>
              <input className="inp" type="number" min="1" max="60" placeholder="Min (1–60)" value={evMin} onChange={(e) => setEvMin(e.target.value)} />
              <button className="btn" onClick={addEvent}>Add</button>
            </div>
            {evMin !== "" && quarterOf(evMin) && <p className="hint" style={{ marginTop: 0 }}>Minute {evMin} → Quarter {quarterOf(evMin)}</p>}
            <div className="ev-list">
              {events.length === 0 && <p className="amut" style={{ fontSize: 12.5 }}>No events added.</p>}
              {events.map((e) => (
                <div className="ev-row" key={e.id}>
                  {e.min ? <span className="ev-min">{e.min}'</span> : <span className="ev-min muted2">—</span>}
                  <span className={"ev-tag " + e.type}>{evLabel(e.type)}</span>
                  <span style={{ flex: 1 }}>{playerName(players, e.player)} <small style={{ color: MUTE }}>({teamName(teams, e.team)}{e.min ? ` · Q${quarterOf(e.min)}` : ""})</small></span>
                  <button className="ib danger" onClick={() => removeEvent(e.id)}>{I.trash()}</button>
                </div>
              ))}
            </div>
          </>
        )}
      </div>

      <div className="fcard">
        <div className="fct">Player of the Match</div>
        {!hasSquad ? (
          <p className="amut" style={{ fontSize: 13 }}>Build the squads to pick a Player of the Match.</p>
        ) : (
          <select className="inp" value={pom} onChange={(e) => setPom(e.target.value)}>
            <option value="">Select…</option>
            {squadOptions(homeSquad, homeT.name)}
            {squadOptions(awaySquad, awayT.name)}
          </select>
        )}
      </div>

      <div style={{ display: "flex", gap: 10 }}>
        <button className="btn" onClick={() => save(true)}>Save & publish result</button>
        <button className="btn ghost" onClick={() => save(false)}>Save draft</button>
      </div>
    </div>
  );
}

function RegManager({ reg, setReg, flash }) {
  const [draft, setDraft] = useState(reg);
  useEffect(() => setDraft(reg), [reg]);
  if (!draft) return null;

  const setForm = (key, patch) => {
    setDraft({ ...draft, forms: draft.forms.map((f) => f.key === key ? { ...f, ...patch } : f) });
  };
  const uploadGuide = async (key, file) => {
    if (!file) return;
    const r = await readFileSafe(file);
    if (!r.ok) return flash(r.msg);
    setForm(key, { guide: r.dataUrl, guideName: file.name });
    flash("Guide attached — remember to Save");
  };
  const save = async () => { setReg(draft); await saveJSON("dh_reg", draft); flash("Registration saved"); };

  return (
    <div>
      <div className="fcard">
        <div className="fct">Section description</div>
        <textarea className="inp" rows={4} value={draft.desc} onChange={(e) => setDraft({ ...draft, desc: e.target.value })} />
      </div>

      {draft.forms.map((f) => (
        <div className="fcard" key={f.key}>
          <div className="fct">{f.label}</div>
          <label className="fld"><span>Status</span>
            <select className="inp" value={f.status} onChange={(e) => setForm(f.key, { status: e.target.value })}>
              <option value="live">Live (form is open)</option>
              <option value="soon">Coming soon</option>
            </select>
          </label>
          {f.status === "live" && (
            <>
              <label className="fld"><span>Form link (Google Form or any URL)</span>
                <input className="inp" placeholder="https://forms.gle/…" value={f.url || ""} onChange={(e) => setForm(f.key, { url: e.target.value })} /></label>
              <label className="fld"><span>Guide PDF (optional — shows a download button)</span>
                <input className="inp" type="file" accept=".pdf,.doc,.docx" onChange={(e) => uploadGuide(f.key, e.target.files[0])} /></label>
              <div className="ms">{f.guide ? `Guide attached: ${f.guideName || "file"}` : "No guide attached"}{f.guide && <button className="linkbtn" onClick={() => setForm(f.key, { guide: "", guideName: "" })}>remove</button>}</div>
            </>
          )}
        </div>
      ))}
      <button className="btn" onClick={save}>Save registration</button>
    </div>
  );
}

function FihManager({ items, setItems, flash }) {
  const [title, setTitle] = useState("");
  const [sub, setSub] = useState("");
  const [url, setUrl] = useState("");
  const [file, setFile] = useState(null);
  const [busy, setBusy] = useState(false);
  const fr = useRef(null);
  const add = async () => {
    if (!title.trim()) return flash("Add a title");
    if (!url.trim() && !file) return flash("Add a link or a file");
    setBusy(true);
    let dataUrl = "", fileName = "", kind = "File";
    if (file) {
      const r = await readFileSafe(file);
      if (!r.ok) { setBusy(false); return flash(r.msg); }
      dataUrl = r.dataUrl; fileName = file.name; kind = fileKind(file.name, file.type);
    }
    const next = [...items, { id: uid(), title: title.trim(), sub: sub.trim(), url: url.trim(), dataUrl, fileName, kind }];
    setItems(next); await saveJSON("dh_fih", next);
    setTitle(""); setSub(""); setUrl(""); setFile(null); if (fr.current) fr.current.value = ""; setBusy(false); flash("Added");
  };
  const remove = async (id) => { const n = items.filter((x) => x.id !== id); setItems(n); await saveJSON("dh_fih", n); flash("Deleted"); };
  const move = async (i, dir) => { const j = i + dir; if (j < 0 || j >= items.length) return; const n = [...items]; [n[i], n[j]] = [n[j], n[i]]; setItems(n); await saveJSON("dh_fih", n); };
  return (
    <div>
      <div className="fcard">
        <div className="fct">Add a rule / regulation</div>
        <input className="inp" placeholder="Title (e.g. FIH Rules of Hockey — Outdoor)" value={title} onChange={(e) => setTitle(e.target.value)} />
        <input className="inp" placeholder="Subtitle (optional, e.g. 2024 edition · fih.hockey)" value={sub} onChange={(e) => setSub(e.target.value)} />
        <input className="inp" placeholder="Link (https://…) — use this OR upload a file" value={url} onChange={(e) => setUrl(e.target.value)} />
        <label className="fld"><span>Or upload a file (PDF, Word) — overrides the link</span><input ref={fr} className="inp" type="file" accept=".pdf,.doc,.docx,.xls,.xlsx" onChange={(e) => setFile(e.target.files[0])} /></label>
        <button className="btn" onClick={add} disabled={busy}>{busy ? "Saving…" : "Add to page"}</button>
        <p className="hint">Visitors get an "Open" button for links and a "Download" button for files.</p>
      </div>
      <p className="amut" style={{ fontSize: 12.5 }}>Use the arrows to reorder — this is the order visitors see.</p>
      <div className="mlist">
        {items.length === 0 && <p className="amut" style={{ fontSize: 13 }}>No rules added yet.</p>}
        {items.map((d, i) => (
          <div className="mrow" key={d.id}>
            <div className="obtns"><button className="ib" onClick={() => move(i, -1)} disabled={i === 0}>{I.up()}</button><button className="ib" onClick={() => move(i, 1)} disabled={i === items.length - 1}>{I.down()}</button></div>
            <span className="ordn">{i + 1}</span>
            <div style={{ flex: 1, minWidth: 0 }}><div className="mt">{d.title}</div><div className="ms">{d.dataUrl ? `${d.kind} file` : (d.url || "No link")}</div></div>
            <button className="ib danger" onClick={() => remove(d.id)}>{I.trash()}</button>
          </div>
        ))}
      </div>
    </div>
  );
}

function ManageList({ items, onRemove }) {
  return (
    <div className="mlist">
      {items.length === 0 && <p className="amut" style={{ fontSize: 13 }}>Nothing uploaded yet.</p>}
      {items.map((d) => (
        <div className="mrow" key={d.id}>
          <div style={{ flex: 1, minWidth: 0 }}><div className="mt">{d.title}</div><div className="ms">{fmtDate(d.date)} · {d.kind}</div></div>
          <button className="ib danger" onClick={() => onRemove(d.id)}>{I.trash()}</button>
        </div>
      ))}
    </div>
  );
}

function Overlay({ children, onClose, wide }) {
  return <div className="ov" onClick={(e) => e.target.classList.contains("ov") && onClose()}><div className={"pnl" + (wide ? " wide" : "")}>{children}</div></div>;
}
function Toast({ children }) { return <div className="tst">{children}</div>; }

// ---------- defaults ----------
const DEFAULT_INTRO = "Delhi Hockey is the recognised state association affiliated to Hockey India and the Delhi Olympic Association. From grassroots development to elite high-performance programs, we are dedicated to bringing glory to the state and the nation. Key governance documents are available below.";
const DEFAULT_ABOUT = [
  { id: "a1", title: "The Selection Committee", date: "2024-04-01", dataUrl: "", fileName: "", kind: "File" },
  { id: "a2", title: "The Office Bearers", date: "2024-04-01", dataUrl: "", fileName: "", kind: "File" },
  { id: "a3", title: "Financial Audits 2024–25", date: "2024-04-01", dataUrl: "", fileName: "", kind: "File" },
];
const DEFAULT_FIH = [
  { id: "f1", title: "FIH Rules of Hockey (Outdoor)", sub: "Official FIH rulebook · fih.hockey", url: "https://www.fih.hockey/the-game/rules-of-hockey", dataUrl: "", fileName: "", kind: "File" },
  { id: "f2", title: "FIH Rules of Indoor Hockey", sub: "Official FIH rulebook · fih.hockey", url: "https://www.fih.hockey/the-game/rules-of-hockey", dataUrl: "", fileName: "", kind: "File" },
];
const DEFAULT_REG = {
  desc: "Delhi Hockey runs a separate registration system, in line with the Hockey India Registration portal. The purpose is to build our own database of all stakeholders for better understanding and mapping of the ecosystem. Registration on the Delhi Hockey system is mandatory if you wish to participate and be part of the ecosystem.",
  forms: [
    { key: "athlete", label: "Athlete Registration", status: "live", url: "https://forms.gle/Wcmm9vzrsmmH27dYA", guide: "", guideName: "" },
    { key: "coach", label: "Coach Registration", status: "soon", url: "", guide: "", guideName: "" },
    { key: "academy", label: "Academy / Club Registration", status: "soon", url: "", guide: "", guideName: "" },
    { key: "referee", label: "Referee Registration", status: "soon", url: "", guide: "", guideName: "" },
  ],
};

// ============================================================
// CSS
// ============================================================
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700;800;900&family=Inter:wght@400;500;600;700&display=swap');
*{box-sizing:border-box}
.wrap{max-width:1240px;margin:0 auto;padding:0 28px}

/* header */
.hdr{position:sticky;top:0;z-index:40;background:#fff;border-bottom:1px solid ${LINE}}
.hdr-in{display:flex;align-items:center;justify-content:space-between;padding:15px 28px}
.wordmark-sm{background:none;border:none;cursor:pointer;font-family:'Archivo',sans-serif;font-weight:800;font-size:19px;letter-spacing:-.3px;color:${ACCENT};padding:0}
.wordmark-sm span{color:${GOLD}}
.dnav{display:flex;gap:4px}
.nav{background:none;border:none;color:${MUTE};font:inherit;font-size:13px;font-weight:600;letter-spacing:.5px;padding:8px 12px;cursor:pointer;border-bottom:2px solid transparent;transition:.15s}
.nav:hover{color:${ACCENT}}
.nav.on{color:${ACCENT};border-bottom-color:${ORANGE}}
.burger{display:none;flex-direction:column;gap:5px;background:none;border:none;cursor:pointer;padding:6px}
.burger span{width:24px;height:2.5px;background:${ACCENT};border-radius:2px}
.mnav{display:flex;flex-direction:column;padding:6px 18px 16px;gap:2px;background:#fff;border-bottom:1px solid ${LINE}}
.mnav-l{text-align:left;background:none;border:none;color:${MUTE};font:inherit;font-size:14px;font-weight:600;padding:11px 8px;cursor:pointer}
.mnav-l.on{color:${ACCENT}}

/* hero */
.hero{position:relative;overflow:hidden;background:${BG}}
.hero-grid{display:grid;grid-template-columns:1.1fr .9fr;gap:48px;align-items:center;min-height:64vh;padding:64px 28px}
.mega{font-family:'Archivo',sans-serif;font-weight:800;font-size:clamp(56px,9vw,118px);line-height:.92;letter-spacing:-2px;margin:0;color:${ACCENT};text-transform:uppercase}
.mega::after{content:'';display:block;width:64px;height:5px;background:${ORANGE};border-radius:3px;margin-top:18px}
.hero-copy{font-size:17px;line-height:1.65;color:${TEXT};max-width:440px;margin:24px 0 26px}
.socs{display:flex;gap:12px}
.soc{width:40px;height:40px;border-radius:50%;display:grid;place-items:center;background:${BLUE_SOFT};color:${ACCENT};transition:.15s}
.soc:hover{background:${ACCENT};color:#fff}
.soc.sm{width:36px;height:36px}
.cta{margin-top:28px;background:${ACCENT};color:#fff;border:none;font-family:'Inter',sans-serif;font-weight:700;font-size:15px;padding:14px 30px;border-radius:8px;cursor:pointer;transition:.15s}
.cta:hover{background:${ACCENT2}}

/* live documents */
.live{background:#fff;border:1px solid ${LINE};border-radius:12px;overflow:hidden;align-self:stretch;max-height:560px;display:flex;flex-direction:column}
.live-head{display:flex;align-items:center;gap:10px;background:#fff;color:${ACCENT};border:none;border-bottom:2px solid ${GOLD};font-family:'Inter',sans-serif;font-weight:700;font-size:15px;letter-spacing:.3px;padding:16px 20px;cursor:pointer;width:100%;text-align:left}
.live-dot{width:8px;height:8px;border-radius:50%;background:${ORANGE}}
.badge-new{background:${ORANGE_SOFT};color:${ORANGE_DEEP};font-size:10px;font-weight:700;letter-spacing:.5px;padding:3px 8px;border-radius:20px;flex-shrink:0}
.pill-open{background:${ORANGE};color:#3A2400;font-size:11px;font-weight:700;padding:5px 11px;border-radius:20px;display:inline-block}
.live-all{margin-left:auto;opacity:.8}
.live-track{overflow-y:auto;flex:1}
.live-track::-webkit-scrollbar{width:5px}.live-track::-webkit-scrollbar-thumb{background:${LINE};border-radius:3px}
.live-row{display:flex;align-items:center;gap:12px;padding:15px 20px;border-bottom:1px solid ${LINE}}
.live-date{color:${GOLD_DEEP};font-weight:600;font-size:12px;letter-spacing:.2px;text-transform:uppercase}
.live-title{color:${TEXT};font-size:14.5px;font-weight:600;line-height:1.35;margin-top:4px}
.live-dl{color:${ACCENT};display:grid;place-items:center;width:30px;height:30px;border-radius:8px;flex-shrink:0;transition:.15s}
.live-dl:hover{color:#fff;background:${ACCENT}}
.live-empty{padding:30px 22px;color:${MUTE};font-size:14px;text-align:center}

/* pages */
.page{padding:56px 0 90px;background:#fff}
.ptitle{font-family:'Archivo',sans-serif;font-weight:800;font-size:clamp(30px,4vw,44px);letter-spacing:-1px;margin:0;color:${ACCENT}}
.psub{color:${MUTE};font-size:15.5px;max-width:640px;margin:12px 0 0;line-height:1.6}
.lead{font-size:16.5px;line-height:1.7;color:${TEXT};max-width:780px;margin:0 0 34px}

.dlist{display:flex;flex-direction:column;gap:10px}
.drow{display:flex;align-items:center;gap:18px;background:#fff;border:1px solid ${LINE};border-radius:10px;padding:18px 20px;text-decoration:none;transition:.15s}
.drow.link:hover{border-color:${ACCENT}}
.dnum{font-weight:700;color:${GOLD_DEEP};font-size:15px;width:26px}
.dtitle{font-weight:600;font-size:16px;color:${TEXT};line-height:1.35}
.dsub{font-size:12.5px;color:${MUTE};margin-top:5px;letter-spacing:.2px}
.pill{display:inline-flex;align-items:center;gap:7px;border:1.5px solid ${ACCENT};color:${ACCENT};background:#fff;border-radius:8px;padding:9px 16px;font-size:13px;font-weight:600;cursor:pointer;text-decoration:none;flex-shrink:0;transition:.15s}
.pill.solid{background:${ACCENT};border-color:${ACCENT};color:#fff}
.pill.solid:hover{background:${ACCENT2}}
.pill.ghost{opacity:.5;cursor:default;border-color:${LINE};color:${MUTE}}
.pill:hover{background:${ACCENT};color:#fff}
.pill.solid:hover{color:#fff}

.empty{background:#fff;border:1px dashed ${LINE};border-radius:10px;padding:40px;text-align:center;color:${MUTE};font-size:14.5px}

/* partners */
.pgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:16px}
.pcard{background:#fff;border:1px solid ${LINE};border-radius:10px;padding:26px 18px;text-align:center;text-decoration:none;transition:.15s}
.pcard:hover{border-color:${ACCENT}}
.pcard img{max-width:100%;max-height:74px;object-fit:contain;margin-bottom:14px}
.pfall{width:60px;height:60px;border-radius:10px;background:${ACCENT};color:#fff;font-family:'Archivo';font-weight:700;font-size:24px;display:grid;place-items:center;margin:0 auto 14px}
.pfall.sm{width:34px;height:34px;font-size:15px;border-radius:8px;margin:0}
.pname{font-weight:700;color:${TEXT};font-size:15px}
.plink{font-size:12.5px;color:${ACCENT};font-weight:600;margin-top:7px;display:inline-flex;gap:5px;align-items:center;justify-content:center}

/* contact */
.cgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:18px}
.ccard{background:#fff;border:1px solid ${LINE};border-radius:10px;padding:28px}
.clabel{color:${GOLD_DEEP};font-weight:800;font-size:12px;letter-spacing:1.5px;text-transform:uppercase}
.cbody{font-size:16px;line-height:1.7;margin:12px 0 0;color:${TEXT}}
.clink{color:${ACCENT};text-decoration:none;font-weight:600}.clink:hover{text-decoration:underline}

/* support / sponsorship */
.sup-vision{background:${BLUE_SOFT};border:1px solid ${LINE};border-left:4px solid ${ORANGE};border-radius:12px;padding:26px 28px;margin-bottom:8px}
.sup-vision-tag{color:${ORANGE_DEEP};font-weight:800;font-size:12px;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:10px}
.sup-vision-text{font-size:17px;line-height:1.7;color:${TEXT};margin:0}
.sup-h2{font-family:'Archivo',sans-serif;font-weight:800;font-size:22px;color:${ACCENT};margin:44px 0 18px;letter-spacing:-.3px}
.sup-intro{font-size:15px;line-height:1.6;color:${MUTE};margin:-8px 0 20px;max-width:620px}
.sup-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.sup-card{background:#fff;border:1px solid ${LINE};border-radius:12px;padding:22px;display:flex;flex-direction:column}
.sup-card-bar{width:38px;height:4px;border-radius:3px;background:${ORANGE};margin-bottom:14px}
.sup-card-title{font-weight:700;font-size:16px;color:${ACCENT};line-height:1.3;margin-bottom:8px}
.sup-card-body{font-size:14px;line-height:1.6;color:${MUTE}}
.sup-why{display:flex;flex-direction:column;gap:14px}
.sup-why-item{display:flex;gap:16px;align-items:flex-start;background:#fff;border:1px solid ${LINE};border-radius:10px;padding:18px 20px;font-size:15px;line-height:1.6;color:${TEXT}}
.sup-why-num{font-family:'Archivo';font-weight:800;font-size:20px;color:${GOLD};flex-shrink:0;line-height:1.3}
.sup-why-item b{color:${ACCENT}}
.sup-tiers{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.sup-tier{position:relative;background:#fff;border:1px solid ${LINE};border-radius:12px;padding:24px 22px}
.sup-tier.feat{border-color:${ORANGE};border-width:2px;box-shadow:0 6px 20px rgba(255,165,0,.10)}
.sup-tier-flag{position:absolute;top:-11px;left:22px;background:${ORANGE};color:#3A2400;font-size:10px;font-weight:800;letter-spacing:.5px;padding:4px 12px;border-radius:20px}
.sup-tier-name{font-family:'Archivo';font-weight:800;font-size:18px;color:${ACCENT}}
.sup-tier-scope{font-size:13px;color:${GOLD_DEEP};font-weight:600;margin:4px 0 14px}
.sup-tier-perks{list-style:none;padding:0;margin:0;display:flex;flex-direction:column;gap:10px}
.sup-tier-perks li{font-size:14px;line-height:1.5;color:${TEXT};padding-left:22px;position:relative}
.sup-tier-perks li::before{content:'';position:absolute;left:0;top:6px;width:10px;height:10px;border-radius:50%;background:${BLUE_SOFT};border:2px solid ${GOLD}}
.sup-cta{margin-top:44px;background:${ACCENT};border-radius:16px;padding:38px 34px;text-align:center;color:#fff}
.sup-cta-title{font-family:'Archivo';font-weight:800;font-size:24px;margin-bottom:12px}
.sup-cta-body{font-size:15.5px;line-height:1.7;color:rgba(255,255,255,.85);max-width:560px;margin:0 auto 24px}
.sup-cta-btn{margin-top:0;background:${ORANGE};color:#3A2400}
.sup-cta-btn:hover{background:#ffb733}
.sup-cta-note{margin-top:18px;font-size:13.5px;color:rgba(255,255,255,.75)}
.sup-cta-note .clink{color:${GOLD}}
@media(max-width:760px){.sup-grid{grid-template-columns:1fr}.sup-tiers{grid-template-columns:1fr}}

/* footer */
.ftr{background:#fff;border-top:2px solid ${GOLD};margin-top:0}
.ftr-in{display:flex;flex-wrap:wrap;gap:40px;justify-content:space-between;padding:50px 28px 30px}
.ftr-copy{color:${MUTE};font-size:13.5px;line-height:1.65;max-width:340px;margin:16px 0 0}
.ftr-nav{display:flex;flex-direction:column;gap:4px}
.ftr-l{background:none;border:none;color:${MUTE};font:inherit;font-size:13px;font-weight:600;letter-spacing:.4px;padding:5px 0;cursor:pointer;text-align:left}
.ftr-l:hover{color:${ACCENT}}
.ftr-bot{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;padding:18px 28px;border-top:1px solid ${LINE};font-size:12.5px;color:${MUTE}}
.admin-t{background:none;border:1px solid ${LINE};color:${MUTE};font:inherit;font-size:12px;padding:5px 14px;border-radius:40px;cursor:pointer}
.admin-t:hover{color:#fff;border-color:${ACCENT};background:${ACCENT}}

/* admin */
.ov{position:fixed;inset:0;background:rgba(9,22,51,.4);backdrop-filter:blur(4px);z-index:100;display:grid;place-items:center;padding:20px}
.pnl{background:${BG2};border:1px solid ${LINE};border-radius:18px;padding:28px;width:100%;max-width:430px;max-height:90vh;overflow-y:auto}
.pnl.wide{max-width:720px}
.ah{color:${ACCENT};font-family:'Archivo';font-weight:800;font-size:22px;margin:0 0 4px}
.amut{color:${MUTE};font-size:13.5px;margin:0 0 14px}
.arow{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px}
.x{background:#fff;border:1px solid ${LINE};width:34px;height:34px;border-radius:9px;cursor:pointer;font-size:15px;color:${ACCENT}}
.atabs{display:flex;gap:6px;flex-wrap:wrap}
.atab{background:#fff;border:1px solid ${LINE};color:${MUTE};font:inherit;font-size:13px;font-weight:600;padding:8px 15px;border-radius:40px;cursor:pointer}
.atab.on{background:${ACCENT};color:#fff;border-color:${ACCENT}}
.fcard{background:#fff;border:1px solid ${LINE};border-radius:10px;padding:18px;margin-bottom:16px}
.fct{font-weight:700;color:${ACCENT};margin-bottom:12px;font-size:14.5px}
.inp{width:100%;border:1.5px solid ${LINE};border-radius:10px;padding:11px 13px;font:inherit;font-size:14px;margin-bottom:10px;background:#fff;color:${TEXT}}
.inp:focus{outline:none;border-color:${ACCENT}}
.fld{display:block}.fld span{display:block;font-size:12px;color:${MUTE};font-weight:600;margin-bottom:5px}
.r2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.btn{background:${ACCENT};color:#fff;border:none;font:inherit;font-weight:700;font-size:14px;padding:11px 22px;border-radius:10px;cursor:pointer}
.btn:hover{background:${ACCENT2}}
.btn:disabled{opacity:.6;cursor:default}
.hint{font-size:12px;color:${MUTE};margin:8px 0 0}
.mlist{display:flex;flex-direction:column;gap:8px}
.mrow{display:flex;align-items:center;gap:12px;background:#fff;border:1px solid ${LINE};border-radius:11px;padding:12px 14px;color:${TEXT}}
.mt{font-weight:600;font-size:14px;color:${TEXT};overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ms{font-size:12px;color:${MUTE};margin-top:2px}
.ib{background:${BLUE_SOFT};border:none;width:32px;height:30px;border-radius:8px;cursor:pointer;color:${ACCENT};display:grid;place-items:center}
.ib:hover{background:#d8e3f7}.ib:disabled{opacity:.4;cursor:default}
.ib.danger{color:#C0392B}.ib.danger:hover{background:#fbe3e0}
.obtns{display:flex;flex-direction:column;gap:3px}
.ordn{font-family:'Archivo';font-weight:800;color:${GOLD_DEEP};width:20px;text-align:center}
.tst{position:fixed;bottom:28px;left:50%;transform:translateX(-50%);background:${ACCENT};color:#fff;padding:12px 22px;border-radius:40px;font-size:14px;font-weight:700;z-index:200}
.linkbtn{background:none;border:none;color:#C0392B;font:inherit;font-size:12px;cursor:pointer;margin-left:8px;text-decoration:underline;padding:0}

/* league public */
.ltabs{display:flex;gap:6px;flex-wrap:wrap;border-bottom:2px solid ${LINE};padding-bottom:0}
.sp-strip{margin-bottom:28px}
.sp-strip-h{font-size:11px;font-weight:800;letter-spacing:1.5px;text-transform:uppercase;color:${MUTE};text-align:center;margin-bottom:14px}
.sp-grid{display:flex;flex-wrap:wrap;gap:14px;justify-content:center}
.sp-card{display:flex;flex-direction:column;align-items:center;gap:6px;background:#fff;border:1px solid ${LINE};border-radius:10px;padding:18px 22px;min-width:140px;text-decoration:none;transition:.18s}
.sp-card:hover{border-color:${ACCENT}}
.sp-card img{max-height:54px;max-width:130px;object-fit:contain}
.sp-fallback{width:48px;height:48px;border-radius:12px;background:${ACCENT};color:#fff;font-family:'Archivo';font-weight:800;font-size:20px;display:grid;place-items:center}
.sp-name{font-weight:700;color:${TEXT};font-size:14px}
.sp-title{font-size:11.5px;color:${GOLD_DEEP};font-weight:600;letter-spacing:.3px}
.lg-selectors{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:24px}
.lg-sel{display:flex;flex-direction:column;gap:5px;flex:1;min-width:180px}
.lg-sel span{font-size:11px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:${MUTE}}
.lg-sel .inp{margin:0}
.ctx-bar{display:flex;gap:12px;flex-wrap:wrap;background:${BLUE_SOFT};border:1px solid ${ACCENT};border-radius:12px;padding:14px;margin-top:14px}
.ctx-bar .lg-sel{min-width:160px}
.season-add{display:grid;grid-template-columns:1fr auto;gap:10px}

.ltab{background:none;border:none;border-bottom:2px solid transparent;color:${MUTE};font:inherit;font-size:14px;font-weight:600;padding:10px 14px;cursor:pointer;margin-bottom:-2px}
.ltab:hover{color:${ACCENT}}
.ltab.on{color:${ACCENT};border-bottom-color:${GOLD}}
.tbl-wrap{overflow-x:auto;background:#fff;border:1px solid ${LINE};border-radius:10px}
.tbl{width:100%;border-collapse:collapse;min-width:540px}
.tbl th{font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:${ACCENT};font-weight:700;padding:14px 10px;text-align:center;background:${BLUE_SOFT};border-bottom:1px solid ${LINE}}
.tbl th.l,.tbl td.l{text-align:left;padding-left:18px}
.tbl th.pts{color:${GOLD_DEEP}}
.tbl td{padding:13px 10px;text-align:center;font-size:14px;color:${TEXT};border-bottom:1px solid ${LINE}}
.tbl tr:last-child td{border-bottom:none}
.tbl td.rank{color:${MUTE};font-weight:700;width:40px}
.tbl tbody tr.leader td:first-child{box-shadow:inset 3px 0 0 ${ORANGE}}
.tbl tbody tr.leader td.rank{color:${ORANGE_DEEP};font-weight:800}
.tbl td.team{font-weight:700;color:${ACCENT}}
.tbl td.pts{font-weight:800;color:${GOLD_DEEP};font-size:15px}
.tbl tbody tr:nth-child(even){background:${BG2}}
.tbl tbody tr:hover{background:${BLUE_SOFT}}

.fix-list{display:flex;flex-direction:column;gap:12px}
.fix{background:#fff;border:1px solid ${LINE};border-radius:10px;padding:18px 20px}
.fix-meta{font-size:12px;color:${MUTE};letter-spacing:.3px;margin-bottom:10px;text-transform:uppercase}
.fix-row{display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:14px}
.fix-team{font-weight:700;font-size:16px;color:${TEXT}}
.fix-team.h{text-align:right}.fix-team.a{text-align:left}
.fix-score{background:${ACCENT};color:#fff;font-weight:800;font-size:15px;padding:7px 16px;border-radius:10px;min-width:64px;text-align:center}
.fix-pom{margin-top:12px;color:${ORANGE_DEEP};font-size:13px;font-weight:600;text-align:center}
.fix-extras{margin-top:10px;display:flex;gap:18px;justify-content:center;flex-wrap:wrap;font-size:12.5px;color:${MUTE}}

.stat-cols{display:grid;grid-template-columns:1fr 1fr;gap:30px}
.stat-h{font-family:'Archivo';font-weight:800;font-size:16px;text-transform:uppercase;letter-spacing:.5px;color:${ACCENT};margin:0 0 12px}
.stat-list{display:flex;flex-direction:column;gap:8px}
.stat-row{display:flex;align-items:center;gap:12px;background:#fff;border:1px solid ${LINE};border-radius:11px;padding:11px 14px}
.stat-rank{font-family:'Archivo';font-weight:800;color:${GOLD_DEEP};width:22px;text-align:center}
.stat-name{font-weight:600;font-size:14px;color:${TEXT}}
.stat-team{font-size:11.5px;color:${MUTE};margin-top:2px}
.stat-val{font-family:'Archivo';font-weight:800;font-size:18px;color:${ACCENT}}
.stat-val small{font-size:11px;color:${MUTE};margin-left:2px;font-weight:600}
.cards{display:flex;gap:5px}
.card{width:22px;height:22px;border-radius:5px;display:grid;place-items:center;font-size:11px;font-weight:800;color:#fff}
.card.g{background:#2E9E5B}.card.y{background:#E8B500;color:#3A2E00}.card.r{background:#D0413A}
.muted2{color:${MUTE};font-size:13px}

.sq-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:16px}
.sq-card{background:#fff;border:1px solid ${LINE};border-radius:10px;padding:20px}
.sq-name{font-weight:800;font-size:16px;color:${ACCENT};margin-bottom:12px;font-family:'Archivo';text-transform:uppercase;letter-spacing:.3px}
.sq-list{margin:0;padding-left:20px;color:${TEXT};font-size:14px;line-height:1.9}

/* league admin */
.lsubs{display:flex;gap:6px;flex-wrap:wrap}
.lsub{background:#fff;border:1px solid ${LINE};color:${MUTE};font:inherit;font-size:12.5px;font-weight:600;padding:7px 13px;border-radius:40px;cursor:pointer}
.lsub.on{background:${ACCENT};color:#fff;border-color:${ACCENT}}
.team-block{background:#fff;border:1px solid ${LINE};border-radius:11px;padding:10px 14px}
.squad-edit{margin-top:8px;padding-top:10px;border-top:1px solid ${LINE}}
.squad-ol{margin:10px 0 0;padding-left:20px;color:${TEXT};font-size:13.5px;line-height:1.6}
.squad-ol li{display:flex;justify-content:space-between;align-items:center;padding-right:4px}
.squad-ol li span{flex:1}
.score-row{display:flex;align-items:center;justify-content:center;gap:16px}
.score-side{display:flex;align-items:center;gap:10px;font-weight:700;color:${TEXT};font-size:14px}
.inp.num{width:64px;text-align:center;margin:0;font-size:18px;font-weight:800;padding:8px}
.score-dash{font-size:24px;color:${MUTE}}
.r3{display:grid;grid-template-columns:1fr 1.4fr auto;gap:10px}
.r4{display:grid;grid-template-columns:1.1fr 1.5fr .8fr auto;gap:8px}
.ev-min{font-family:'Archivo';font-weight:800;font-size:13px;color:${GOLD_DEEP};min-width:30px;text-align:center}
.chip-row{display:flex;flex-wrap:wrap;gap:8px;margin:4px 0 10px}
.chip{display:inline-flex;align-items:center;gap:7px;background:${BLUE_SOFT};border:1px solid ${LINE};border-radius:40px;padding:5px 8px 5px 13px;font-size:13px;color:${ACCENT};font-weight:600}
.chip button{background:none;border:none;color:${MUTE};cursor:pointer;font-size:12px;padding:0;line-height:1}
.chip button:hover{color:#C0392B}

/* match report (public) */
.fix-report-toggle{margin-top:14px;text-align:center}
.report-btn{display:inline-flex;align-items:center;gap:7px;background:${BLUE_SOFT};border:1px solid ${LINE};color:${ACCENT};font:inherit;font-size:13px;font-weight:600;padding:8px 18px;border-radius:40px;cursor:pointer}
.report-btn:hover{border-color:${ACCENT};background:#dde8fb}
.report{margin-top:16px;padding-top:16px;border-top:1px solid ${LINE}}
.report-pom{text-align:center;color:${ORANGE_DEEP};font-size:14px;margin-bottom:16px;font-weight:600}
.report-stats{background:${BG2};border:1px solid ${LINE};border-radius:12px;padding:14px 18px;margin-bottom:16px}
.rs-row{display:grid;grid-template-columns:1fr 2fr 1fr;align-items:center;padding:6px 0}
.rs-v{font-family:'Archivo';font-weight:800;font-size:18px;color:${ACCENT};text-align:center}
.rs-l{font-size:12px;color:${MUTE};text-align:center;text-transform:uppercase;letter-spacing:.5px}
.rs-teams{display:flex;justify-content:space-between;border-top:1px solid ${LINE};margin-top:8px;padding-top:8px;font-size:12px;color:${MUTE};font-weight:600}
.report-timeline{display:flex;flex-direction:column;gap:14px}
.rq-head{font-family:'Archivo';font-weight:800;font-size:13px;text-transform:uppercase;letter-spacing:.5px;color:${ACCENT};margin-bottom:8px}
.rq-head span{color:${MUTE};font-weight:600;font-size:11px;letter-spacing:0}
.rq-ev{display:flex;align-items:center;gap:10px;padding:8px 12px;border-radius:9px;background:${BG2};border:1px solid ${LINE};margin-bottom:6px;font-size:13.5px}
.rq-ev.away{flex-direction:row-reverse;text-align:right}
.rq-min{font-family:'Archivo';font-weight:800;color:${GOLD_DEEP};min-width:30px}
.rq-name{flex:1;color:${TEXT};font-weight:600}
.rq-team{font-size:11.5px;color:${MUTE}}
.report-refs{margin-top:16px;display:flex;flex-direction:column;gap:4px;align-items:center;padding:12px;background:${BG2};border:1px solid ${LINE};border-radius:10px}
.rr-h{font-size:10.5px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:${MUTE}}
.report-refs>span:last-child{color:${TEXT};font-size:13.5px;font-weight:600}
.ev-list{margin-top:12px;display:flex;flex-direction:column;gap:6px}
.ev-row{display:flex;align-items:center;gap:10px;background:#fff;border:1px solid ${LINE};border-radius:9px;padding:8px 12px;font-size:13.5px;color:${TEXT}}
.ev-tag{font-size:11px;font-weight:700;padding:3px 9px;border-radius:20px;text-transform:uppercase;letter-spacing:.3px}
.ev-tag.goal{background:${BLUE_SOFT};color:${ACCENT}}
.ev-tag.assist{background:#e7f6ec;color:#2E9E5B}
.ev-tag.green{background:#e7f6ec;color:#2E9E5B}
.ev-tag.yellow{background:#fdf3d6;color:${GOLD_DEEP}}
.ev-tag.red{background:#fbe3e0;color:#C0392B}
.btn.ghost{background:#fff;border:1.5px solid ${ACCENT};color:${ACCENT}}
.btn.ghost:hover{background:${BLUE_SOFT}}
.mono{font-family:ui-monospace,'SF Mono',Menlo,monospace;font-size:11.5px;letter-spacing:.2px}
.sq-id{color:${MUTE};margin-left:auto;padding:0 8px}
.reg-head{display:flex;align-items:center;justify-content:space-between;margin:6px 0 10px}
.pick-list{background:#fff;border:1px solid ${LINE};border-radius:10px;margin-bottom:10px;overflow:hidden}
.pick-row{display:flex;align-items:center;justify-content:space-between;gap:10px;width:100%;background:none;border:none;border-bottom:1px solid ${LINE};color:${TEXT};font:inherit;font-size:13.5px;padding:10px 12px;cursor:pointer;text-align:left}
.pick-row:last-child{border-bottom:none}
.pick-row:hover{background:${BLUE_SOFT};color:${ACCENT}}
.pick-row .mono{color:${MUTE}}
.sq-list li{display:flex;align-items:center;gap:8px}
.sq-list .sq-id{font-family:ui-monospace,monospace;font-size:11px;color:${MUTE};margin-left:auto}
.sq-section{font-size:11px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:${GOLD_DEEP};margin:0 0 8px}
.sq-coaches{background:${BLUE_SOFT};border:1px solid ${LINE};border-radius:10px;padding:12px 14px;margin-bottom:14px}
.sq-coach-h{font-size:10.5px;font-weight:800;letter-spacing:1px;text-transform:uppercase;color:${MUTE};margin-bottom:6px}
.sq-coach{font-size:14px;font-weight:600;color:${TEXT};line-height:1.5}
.sq-coach .mono{color:${MUTE};font-weight:400}

/* registration */
.reg{padding:48px 0 80px;background:#fff}
.reg-h{font-family:'Archivo',sans-serif;font-weight:800;font-size:clamp(26px,3.5vw,36px);letter-spacing:-.5px;margin:0;color:${ACCENT}}
.reg-desc{color:${TEXT};font-size:15.5px;line-height:1.7;max-width:820px;margin:14px 0 32px}
.reg-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}
.reg-card{background:#fff;border:1px solid ${LINE};border-radius:10px;padding:24px 20px;display:flex;flex-direction:column;gap:14px;min-height:150px}
.reg-card.soon{opacity:.65}
.reg-name{font-weight:700;font-size:16px;color:${ACCENT};line-height:1.3}
.reg-btn{margin-top:auto;display:inline-flex;align-items:center;justify-content:center;gap:7px;background:${ACCENT};color:#fff;border:none;font-weight:700;font-size:13.5px;padding:11px 16px;border-radius:8px;text-decoration:none;transition:.15s}
.reg-btn:hover{background:${ACCENT2}}
.reg-guide{display:inline-flex;align-items:center;gap:6px;color:${ACCENT};font-size:12.5px;font-weight:600;text-decoration:none;justify-content:center}
.reg-guide:hover{text-decoration:underline}
.reg-soon{margin-top:auto;text-align:center;color:${MUTE};font-size:13px;font-weight:600;border:1px dashed ${LINE};border-radius:8px;padding:11px 16px}

@media(max-width:980px){
  .dnav{display:none}.burger{display:flex}
  .hero-grid{grid-template-columns:1fr;gap:32px;min-height:auto;padding:32px 28px 50px}
  .live{max-height:420px}
  .mega{font-size:clamp(56px,15vw,90px)}
  .reg-grid{grid-template-columns:1fr 1fr}
}
@media(max-width:560px){
  .r2{grid-template-columns:1fr}
  .pill-t{display:none}
  .mega{letter-spacing:-1.5px}
  .reg-grid{grid-template-columns:1fr}
  .stat-cols{grid-template-columns:1fr;gap:26px}
  .r3{grid-template-columns:1fr}
  .r4{grid-template-columns:1fr 1fr}
}
`;
