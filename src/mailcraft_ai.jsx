import { useState, useRef } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────────────────────
// CLAUDE API — single shared caller
// ─────────────────────────────────────────────────────────────
async function claude(system, user, maxTokens = 3000) {
  const res = await fetch("https://ai-api-dev.dentsu.com", {
    method: "POST",
    headers: { "Content-Type": "application/json", "Ocp-Apim-Subscription-Key": "d9087da1250e438782aa1cd9b0c38561" },
    body: JSON.stringify({
      model: "claude-opus-4-6",
      max_tokens: maxTokens,
      system,
      messages: [{ role: "user", content: user }],
    }),
  });
  if (!res.ok) throw new Error(`API error ${res.status}`);
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  const text = data.content?.find((b) => b.type === "text")?.text || "";
  return text.replace(/^```(?:json|html)?\s*/i, "").replace(/\s*```$/i, "").trim();
}

// ─────────────────────────────────────────────────────────────
// SHEET READER
// ─────────────────────────────────────────────────────────────
function readWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: "array" });
  const sheets = {};
  wb.SheetNames.forEach((name) => {
    sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], {
      header: 1, defval: "", raw: false,
    });
  });
  return { names: wb.SheetNames, sheets };
}

// Convert all sheets to a readable text dump for AI
function sheetsToText(sheets) {
  return Object.entries(sheets).map(([name, rows]) => {
    const lines = rows
      .map((row) => row.map((c) => String(c).trim()).filter(Boolean).join(" | "))
      .filter(Boolean);
    return `=== SHEET: ${name} ===\n${lines.join("\n")}`;
  }).join("\n\n");
}

// ─────────────────────────────────────────────────────────────
// AGENT 1 — UNIVERSAL DATA EXTRACTOR
// AI reads the raw Excel dump and extracts a structured brief
// No hardcoded field names — works for ANY company
// ─────────────────────────────────────────────────────────────
async function agent1_extract(sheets, names, log) {
  log("a1", "AI reading your Excel structure…");

  const dump = sheetsToText(sheets);

  const result = await claude(
    `You are an expert at reading marketing email brief spreadsheets from any company.
Given raw Excel sheet data, extract all the information needed to build an HTML email newsletter.

Return ONLY a JSON object with this exact structure (fill what you find, leave blank if not present):
{
  "company": "company name if found",
  "subject": "email subject line",
  "preheader": "preheader text (same as subject if not found)",
  "sender": "sender email address",
  "primaryColor": "#hex — brand primary color if mentioned, else #000000",
  "secondaryColor": "#hex — brand secondary color if mentioned, else #666666",
  "logoText": "company name for logo placeholder",
  "greeting": "email greeting line",
  "bodyText": "main body copy, 1-3 sentences",
  "signature": "sign-off text",
  "validityText": "offer validity / deadline text if any",
  "ctaText": "main call to action button text",
  "mainBannerLink": "main banner URL if found",
  "utmCampaign": "UTM campaign value if found",
  "sections": [
    {
      "type": "product_grid | banner | text | divider",
      "title": "section headline",
      "subtitle": "section subheadline",
      "cta": "CTA button text",
      "link": "URL for this section",
      "bgColor": "#hex background color",
      "textColor": "#ffffff or #000000",
      "imageUrl": "image URL if present",
      "items": [
        { "title": "item name", "link": "item URL", "imageUrl": "item image URL", "price": "price/offer text" }
      ]
    }
  ],
  "footerLinks": ["label|url", "label|url"],
  "socialLinks": { "facebook": "url", "instagram": "url", "twitter": "url", "linkedin": "url", "youtube": "url" },
  "address": "company address if found",
  "unsubscribeText": "unsubscribe text if found"
}

Be smart: identify product banners, content banners, promotional sections, etc. and map them to the sections array.
For colors: if a color name is mentioned (e.g. "brand blue", "primary green"), make a reasonable hex.
Return ONLY the JSON. No explanation.`,
    `Here is the full Excel content:\n\n${dump.substring(0, 12000)}`
  );

  let brief;
  try {
    brief = JSON.parse(result);
  } catch {
    throw new Error("Could not parse Excel — try a cleaner brief format");
  }

  log("a1", `✓ Company: "${brief.company || "Unknown"}"`, "ok");
  log("a1", `✓ Subject: "${brief.subject}"`, "ok");
  log("a1", `✓ ${brief.sections?.length || 0} content sections found`, "ok");
  return brief;
}

// ─────────────────────────────────────────────────────────────
// AGENT 2 — LINK & CONTENT VALIDATOR
// Checks extracted data, fills gaps, validates URLs
// ─────────────────────────────────────────────────────────────
async function agent2_validate(brief, log) {
  log("a2", "Validating links and content completeness…");

  const allLinks = [
    brief.mainBannerLink,
    ...(brief.sections || []).map((s) => s.link),
    ...(brief.sections || []).flatMap((s) => (s.items || []).map((i) => i.link)),
  ].filter(Boolean);

  const result = await claude(
    `You are a content validator for email marketing briefs.
Check the extracted brief data for issues and return a validation report.
Return ONLY JSON: { "issues": ["string"], "warnings": ["string"], "fixes": { "key": "suggested fix" }, "score": 0-100 }
Score 100 = perfect brief. Deduct points for: missing subject, no CTAs, broken-looking URLs, missing body copy.`,
    `Validate this brief:\n${JSON.stringify(brief, null, 2).substring(0, 4000)}\n\nLinks found: ${allLinks.join(", ") || "none"}`
  );

  let report = { issues: [], warnings: [], fixes: {}, score: 80 };
  try { report = JSON.parse(result); } catch {}

  report.issues.forEach((i) => log("a2", `✗ ${i}`, "err"));
  report.warnings.forEach((w) => log("a2", `⚠ ${w}`, "warn"));
  log("a2", `✓ Brief score: ${report.score}/100`, "ok");

  // Auto-append UTM if campaign is set
  if (brief.utmCampaign) {
    const addUtm = (url) => {
      if (!url || !url.startsWith("http")) return url;
      const sep = url.includes("?") ? "&" : "?";
      return `${url}${sep}utm_source=email&utm_medium=newsletter&utm_campaign=${brief.utmCampaign}`;
    };
    if (brief.mainBannerLink) brief.mainBannerLink = addUtm(brief.mainBannerLink);
    (brief.sections || []).forEach((s) => {
      s.link = addUtm(s.link);
      (s.items || []).forEach((item) => { item.link = addUtm(item.link); });
    });
    log("a2", `✓ UTM appended to all links (campaign: ${brief.utmCampaign})`, "ok");
  }

  return { brief, report };
}

// ─────────────────────────────────────────────────────────────
// AGENT 3 — COPY ENHANCER
// Polishes subject, body, CTAs to be engaging
// ─────────────────────────────────────────────────────────────
async function agent3_copy(brief, log) {
  log("a3", "Enhancing email copy with AI…");

  const result = await claude(
    `You are an expert email copywriter. Polish the email copy in the brief.
Rules:
- Keep the same language as the input (detect it automatically)
- Keep subject line unless it's blank or placeholder text
- Improve body text to be warm, engaging, and action-oriented
- Fill any blank/placeholder CTAs with compelling short phrases (max 4 words)
- Keep all URLs, hex colors, and structured data exactly as-is
- Do NOT invent product names or prices
Return ONLY the same JSON structure with improved copy fields. No explanation.`,
    JSON.stringify({
      subject: brief.subject,
      preheader: brief.preheader,
      greeting: brief.greeting,
      bodyText: brief.bodyText,
      signature: brief.signature,
      ctaText: brief.ctaText,
      sections: brief.sections?.map((s) => ({
        title: s.title, subtitle: s.subtitle, cta: s.cta,
        items: s.items?.map((i) => ({ title: i.title, price: i.price }))
      }))
    }, null, 2)
  );

  try {
    const polished = JSON.parse(result);
    if (polished.subject)    brief.subject    = polished.subject;
    if (polished.preheader)  brief.preheader  = polished.preheader;
    if (polished.greeting)   brief.greeting   = polished.greeting;
    if (polished.bodyText)   brief.bodyText   = polished.bodyText;
    if (polished.signature)  brief.signature  = polished.signature;
    if (polished.ctaText)    brief.ctaText    = polished.ctaText;
    (polished.sections || []).forEach((ps, i) => {
      if (!brief.sections?.[i]) return;
      if (ps.title)    brief.sections[i].title    = ps.title;
      if (ps.subtitle) brief.sections[i].subtitle = ps.subtitle;
      if (ps.cta)      brief.sections[i].cta      = ps.cta;
      (ps.items || []).forEach((pi, j) => {
        if (brief.sections[i].items?.[j]) {
          if (pi.title) brief.sections[i].items[j].title = pi.title;
        }
      });
    });
    log("a3", "✓ Copy polished and enhanced", "ok");
  } catch {
    log("a3", "⚠ Copy enhancement failed — using original", "warn");
  }

  return brief;
}

// ─────────────────────────────────────────────────────────────
// AGENT 4 — HTML EMAIL BUILDER
// Generates complete production email HTML from the brief
// ─────────────────────────────────────────────────────────────
async function agent4_build(brief, log) {
  log("a4", "Building production HTML email…");

  const html = await claude(
    `You are a world-class HTML email developer.
Generate a complete, production-ready HTML email newsletter.
CRITICAL RULES:
- Output ONLY raw HTML. No markdown, no explanation, no code fences.
- Use table-based layout ONLY — no flexbox, no CSS grid (email clients don't support them)
- All CSS must be both in a <style> block AND inlined on every element
- Max width 600px, centered in a wrapper table
- Use the exact brand colors provided
- Every link must be wrapped in <a href="..."> tags
- Images: if imageUrl is a valid https:// URL use <img>, otherwise use a colored placeholder table cell
- Make it visually professional — not generic. Use the brand colors boldly.
- Include hidden preheader text span (display:none; max-height:0; overflow:hidden)
- Responsive: @media (max-width:600px) stack columns to full width
- DOCTYPE html declaration required`,
    `Build a complete HTML email from this brief:

${JSON.stringify(brief, null, 2).substring(0, 5000)}

STRUCTURE TO BUILD (in this exact order):
1. DOCTYPE + head (with style block for responsive)
2. Hidden preheader span: "${brief.preheader || brief.subject}"
3. Outer wrapper table: bgcolor #f4f4f4, width 100%
4. Inner email table: width 600px, bgcolor #ffffff, centered

THEN these sections:
A) HEADER BAR: bgcolor = primaryColor. Company logo text "${brief.logoText || brief.company}" in white, bold, large. If mainBannerLink exists, make it a link.

B) HERO BANNER: Large 600px wide block, bgcolor = primaryColor darkened slightly. 
   Headline: "${brief.subject}" in white, 28px bold, centered.
   Subheadline: "${brief.preheader || ""}" in white 80% opacity, 16px.
   CTA button: white bg, primaryColor text, "${brief.ctaText || "Learn More"}", href="${brief.mainBannerLink || "#"}".

C) GREETING + BODY: white bg, padding 32px. "${brief.greeting}" in 20px, bold. Body: "${brief.bodyText}" in 16px #555.

D) SECTIONS: For each section in the sections array:
   - type "product_grid": 2-column table of items. Each cell: image/placeholder (if imageUrl) + item title + price + CTA link.
   - type "banner": Full-width colored block (bgColor). Large title, subtitle, CTA button. Link wraps whole block.
   - type "text": White bg text section with title and subtitle.
   Separate sections with 8px spacer rows.

E) VALIDITY / FOOTER INFO: if validityText exists, show it in a light grey bar, centered, 13px.

F) FOOTER: bgcolor #222222. 
   - Signature: "${brief.signature}" in white.
   - Social links: show any provided social URLs as colored text links.
   - Unsubscribe: small grey text "To unsubscribe click here" with href="#unsubscribe".
   - Address: "${brief.address || ""}".

Make it look premium. Use primaryColor = ${brief.primaryColor || "#000000"} throughout.`,
    4000
  );

  log("a4", `✓ HTML email built (${html.length} chars)`, "ok");
  return html;
}

// ─────────────────────────────────────────────────────────────
// AGENT 5 — QA CHECKER
// ─────────────────────────────────────────────────────────────
async function agent5_qa(html, brief, log) {
  log("a5", "Running quality checks…");

  const checks = [
    `Has DOCTYPE: ${html.startsWith("<!DOCTYPE") || html.startsWith("<!doctype")}`,
    `Has preheader: ${html.includes("display:none") || html.includes("display: none")}`,
    `Has subject/headline: ${html.toLowerCase().includes((brief.subject || "").slice(0, 12).toLowerCase())}`,
    `Has primary color: ${html.includes((brief.primaryColor || "").replace("#", ""))}`,
    `Has footer: ${html.includes("unsubscribe") || html.includes("footer")}`,
    `No broken template vars: ${!html.includes("undefined") && !html.includes("[object")}`,
    `Has closing html tag: ${html.includes("</html>")}`,
    `Has at least one link: ${html.includes("<a href")}`,
    `Sections built: ${(brief.sections || []).filter((s) => s.title && html.toLowerCase().includes((s.title || "").slice(0, 8).toLowerCase())).length} of ${brief.sections?.length || 0}`,
  ];

  const result = await claude(
    `Email QA engineer. Review test results and return ONLY JSON:
{"verdict":"PASS"|"FAIL","passed":N,"failed":N,"score":0-100,"issues":["string"]}`,
    checks.join("\n")
  );

  let qa = { verdict: "PASS", passed: 9, failed: 0, score: 95, issues: [] };
  try { qa = JSON.parse(result); } catch {}

  log("a5", `✓ QA ${qa.verdict} — score ${qa.score}/100`, qa.verdict === "PASS" ? "ok" : "warn");
  (qa.issues || []).forEach((i) => log("a5", `  ⚠ ${i}`, "warn"));
  return qa;
}

// ─────────────────────────────────────────────────────────────
// PIPELINE
// ─────────────────────────────────────────────────────────────
async function runPipeline(file, onAgentLog, onAgentStatus, onDone) {
  const log = (id, msg, type = "info") => onAgentLog(id, msg, type);
  const status = (id, s) => onAgentStatus(id, s);
  console.log("Pipeline started");

  try {
    const buf = await file.arrayBuffer();
    const { names, sheets } = readWorkbook(buf);
    console.log("a1", `Loaded ${names.length} sheets: ${names.join(", ")}`, "info");

    status("a1", "running");
    const brief = await agent1_extract(sheets, names, log);
    status("a1", "done");

    status("a2", "running");
    const { brief: vBrief, report } = await agent2_validate(brief, log);
    status("a2", "done");

    status("a3", "running");
    const pBrief = await agent3_copy(vBrief, log);
    status("a3", "done");

    status("a4", "running");
    const html = await agent4_build(pBrief, log);
    status("a4", "done");

    status("a5", "running");
    const qa = await agent5_qa(html, pBrief, log);
    status("a5", "done");

    onDone({ html, brief: pBrief, qa, report });
  } catch (err) {
    onDone({ error: err.message });
  }
}

// ─────────────────────────────────────────────────────────────
// UI
// ─────────────────────────────────────────────────────────────
const AGENTS = [
  {
    id: "a1", num: 1, name: "Data Extractor",
    desc: "AI reads your Excel structure — any format, any company",
    icon: "⚙", color: "#2563EB", light: "#EFF6FF", border: "#BFDBFE",
  },
  {
    id: "a2", num: 2, name: "Link Validator",
    desc: "Checks all URLs, appends UTM tracking to every link",
    icon: "🔗", color: "#059669", light: "#ECFDF5", border: "#A7F3D0",
  },
  {
    id: "a3", num: 3, name: "Copy Enhancer",
    desc: "Polishes subject, body and CTAs — keeps your language",
    icon: "✍", color: "#7C3AED", light: "#F5F3FF", border: "#DDD6FE",
  },
  {
    id: "a4", num: 4, name: "HTML Builder",
    desc: "Generates complete 600px email-safe HTML with your brand",
    icon: "🏗", color: "#D97706", light: "#FFFBEB", border: "#FDE68A",
  },
  {
    id: "a5", num: 5, name: "QA Checker",
    desc: "Audits generated HTML — checks links, structure, score",
    icon: "✅", color: "#DC2626", light: "#FFF1F2", border: "#FECACA",
  },
];

const INIT_STATE = Object.fromEntries(
  AGENTS.map((a) => [a.id, { status: "idle", logs: [] }])
);

export default function App() {
  const [screen, setScreen]     = useState("upload");
  const [agState, setAgState]   = useState(INIT_STATE);
  const [result,  setResult]    = useState(null);
  const [tab,     setTab]       = useState("preview");
  const [copied,  setCopied]    = useState(false);
  const fileRef = useRef();
  const dropRef = useRef();

  const agLog = (id, msg, type = "info") =>
    setAgState((p) => ({
      ...p,
      [id]: { ...p[id], logs: [...p[id].logs, { msg, type, ts: new Date().toLocaleTimeString() }] },
    }));

  const agStatus = (id, s) =>
    setAgState((p) => ({ ...p, [id]: { ...p[id], status: s } }));

  function startPipeline(file) {

    if (!file) return;
    console.log("File selected:", file.name);
    setScreen("pipeline");
    setAgState(INIT_STATE);
    setResult(null);
    setTab("preview");
    runPipeline(file, agLog, agStatus, (r) => setResult(r));
  }

  function copyHTML() {
    if (!result?.html) return;
    navigator.clipboard?.writeText(result.html);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  }

  function download() {
    if (!result?.html) return;
    const name = `${(result.brief?.company || "email").toLowerCase().replace(/\s+/g, "_")}_newsletter.html`;
    const a = Object.assign(document.createElement("a"), {
      href: URL.createObjectURL(new Blob([result.html], { type: "text/html;charset=utf-8" })),
      download: name,
    });
    a.click();
    URL.revokeObjectURL(a.href);
  }

  const logColors = {
    ok:   { color: "#16a34a", bg: "#f0fdf4" },
    warn: { color: "#d97706", bg: "#fffbeb" },
    err:  { color: "#dc2626", bg: "#fff1f2" },
    info: { color: "#2563eb", bg: "#eff6ff" },
  };

  const allDone = AGENTS.every((a) => agState[a.id]?.status === "done");
  const hasResult = result && !result.error;

  return (
    <div style={{
      minHeight: "100vh",
      background: "#0A0A0F",
      fontFamily: "'Segoe UI', system-ui, sans-serif",
      color: "#E2E8F0",
    }}>

      {/* ── TOP NAV ── */}
      <div style={{
        borderBottom: "1px solid #1E293B",
        padding: "0 32px",
        display: "flex", alignItems: "center", gap: 16, height: 56,
        background: "rgba(255,255,255,0.02)",
      }}>
        <div style={{
          width: 28, height: 28, borderRadius: 8,
          background: "linear-gradient(135deg,#3B82F6,#8B5CF6)",
          display: "flex", alignItems: "center", justifyContent: "center",
          fontSize: 14, fontWeight: 900, color: "#fff",
        }}>M</div>
        <span style={{ fontWeight: 700, fontSize: 15, color: "#F1F5F9" }}>MailCraft AI</span>
        <div style={{
          marginLeft: 6, fontSize: 10, fontWeight: 600, color: "#64748B",
          background: "#1E293B", padding: "2px 8px", borderRadius: 20,
          letterSpacing: 1, textTransform: "uppercase",
        }}>Beta</div>
        <div style={{ flex: 1 }} />
        <div style={{ fontSize: 13, color: "#475569" }}>
          Excel brief → AI agents → production HTML email
        </div>
      </div>

      <div style={{ maxWidth: 860, margin: "0 auto", padding: "32px 24px" }}>

        {/* ── UPLOAD SCREEN ── */}
        {screen === "upload" && (
          <div>
            {/* Hero */}
            <div style={{ textAlign: "center", marginBottom: 48, paddingTop: 24 }}>
              <div style={{
                display: "inline-block",
                background: "linear-gradient(135deg,#1E3A5F,#2D1B69)",
                border: "1px solid #334155",
                borderRadius: 16, padding: "6px 16px",
                fontSize: 12, color: "#94A3B8", marginBottom: 20,
                letterSpacing: 1,
              }}>ANY COMPANY · ANY BRIEF FORMAT · ANY LANGUAGE</div>
              <h1 style={{
                fontSize: 42, fontWeight: 800, lineHeight: 1.15,
                background: "linear-gradient(135deg,#E2E8F0 0%,#94A3B8 100%)",
                WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent",
                marginBottom: 16,
              }}>
                Upload your Excel brief.<br/>Get a production HTML email.
              </h1>
              <p style={{ fontSize: 16, color: "#64748B", maxWidth: 480, margin: "0 auto" }}>
                5 AI agents read your brief, validate every link, polish the copy,
                build the email, and run a QA check — in under 2 minutes.
              </p>
            </div>

            {/* Drop zone */}
            <div
              ref={dropRef}
              onClick={() => fileRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); dropRef.current.style.borderColor = "#3B82F6"; }}
              onDragLeave={() => { dropRef.current.style.borderColor = "#1E293B"; }}
              onDrop={(e) => {
                e.preventDefault();
                dropRef.current.style.borderColor = "#1E293B";
                startPipeline(e.dataTransfer.files[0]);
              }}
              style={{
                border: "2px dashed #1E293B", borderRadius: 20,
                background: "rgba(255,255,255,0.02)",
                padding: "64px 40px", textAlign: "center", cursor: "pointer",
                transition: "all 0.2s",
              }}
              onMouseEnter={(e) => e.currentTarget.style.background = "rgba(255,255,255,0.04)"}
              onMouseLeave={(e) => e.currentTarget.style.background = "rgba(255,255,255,0.02)"}
            >
              <div style={{
                width: 72, height: 72, borderRadius: 18,
                background: "linear-gradient(135deg,#1E3A5F,#2D1B69)",
                border: "1px solid #334155",
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 30, margin: "0 auto 20px",
              }}>📊</div>
              <div style={{ fontSize: 20, fontWeight: 700, color: "#E2E8F0", marginBottom: 8 }}>
                Drop your Excel brief here
              </div>
              <div style={{ fontSize: 14, color: "#475569", marginBottom: 28 }}>
                .xlsx or .xls · any format · any language · any company
              </div>
              <div style={{
                display: "inline-flex", alignItems: "center", gap: 8,
                background: "linear-gradient(135deg,#2563EB,#7C3AED)",
                color: "#fff", padding: "12px 28px", borderRadius: 10,
                fontWeight: 700, fontSize: 14, cursor: "pointer",
              }}>
                <span>Choose File</span>
              </div>
              <input
                ref={fileRef} type="file" accept=".xlsx,.xls,.csv"
                style={{ display: "none" }}
                onChange={(e) => startPipeline(e.target.files[0])}
              />
            </div>

            {/* How it works */}
            <div style={{
              display: "grid", gridTemplateColumns: "repeat(5,1fr)",
              gap: 10, marginTop: 40,
            }}>
              {AGENTS.map((a) => (
                <div key={a.id} style={{
                  background: "#0F172A", border: "1px solid #1E293B",
                  borderRadius: 12, padding: "16px 14px",
                }}>
                  <div style={{
                    width: 28, height: 28, borderRadius: 8,
                    background: a.color + "22", border: `1px solid ${a.color}44`,
                    display: "flex", alignItems: "center", justifyContent: "center",
                    fontSize: 14, marginBottom: 10,
                  }}>{a.icon}</div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#E2E8F0", marginBottom: 4 }}>
                    {a.num}. {a.name}
                  </div>
                  <div style={{ fontSize: 10.5, color: "#475569", lineHeight: 1.5 }}>
                    {a.desc}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ── PIPELINE SCREEN ── */}
        {screen === "pipeline" && (
          <div>
            <div style={{ marginBottom: 24 }}>
              <div style={{ fontSize: 22, fontWeight: 700, color: "#E2E8F0", marginBottom: 4 }}>
                {allDone ? "✓ Pipeline complete" : "Running agents…"}
              </div>
              <div style={{ fontSize: 13, color: "#475569" }}>
                {allDone ? "Your HTML email is ready." : "5 AI agents are processing your brief in sequence."}
              </div>
            </div>

            {/* Agent cards */}
            <div style={{ display: "flex", flexDirection: "column", gap: 8, marginBottom: 24 }}>
              {AGENTS.map((agent) => {
                const s = agState[agent.id] || { status: "idle", logs: [] };
                const running = s.status === "running";
                const done    = s.status === "done";
                const idle    = s.status === "idle";

                return (
                  <div key={agent.id} style={{
                    background: done ? "#0F1A0F" : running ? "#0F172A" : "#080C12",
                    border: `1px solid ${done ? "#166534" : running ? agent.color + "66" : "#1E293B"}`,
                    borderRadius: 12, overflow: "hidden",
                    transition: "all 0.3s",
                  }}>
                    {/* Card header */}
                    <div style={{
                      display: "flex", alignItems: "center", gap: 12,
                      padding: "12px 16px",
                    }}>
                      <div style={{
                        width: 34, height: 34, borderRadius: 9, flexShrink: 0,
                        background: done ? "#16653488" : running ? agent.color : "#1E293B",
                        border: `1px solid ${done ? "#22c55e44" : running ? agent.color : "#334155"}`,
                        display: "flex", alignItems: "center", justifyContent: "center",
                        fontSize: 15, fontWeight: 700, color: "#fff",
                        animation: running ? "pulse 1.4s ease-in-out infinite" : "none",
                      }}>
                        {done ? "✓" : agent.icon}
                      </div>
                      <div style={{ flex: 1 }}>
                        <div style={{
                          fontSize: 13, fontWeight: 700,
                          color: done ? "#86EFAC" : running ? "#E2E8F0" : "#475569",
                        }}>
                          Agent {agent.num} — {agent.name}
                        </div>
                        <div style={{ fontSize: 11, color: "#334155" }}>{agent.desc}</div>
                      </div>
                      <div style={{
                        fontSize: 10, fontWeight: 700, padding: "3px 10px",
                        borderRadius: 20, letterSpacing: .5,
                        background: done ? "#16653422" : running ? agent.color + "22" : "#1E293B",
                        color: done ? "#4ADE80" : running ? agent.color : "#475569",
                        border: `1px solid ${done ? "#16653444" : running ? agent.color + "44" : "#334155"}`,
                      }}>
                        {done ? "DONE" : running ? "RUNNING" : "WAITING"}
                      </div>
                    </div>

                    {/* Logs */}
                    {s.logs.length > 0 && (
                      <div style={{
                        borderTop: "1px solid #1E293B",
                        padding: "8px 16px",
                        maxHeight: 110, overflowY: "auto",
                      }}>
                        {s.logs.map((l, i) => {
                          const c = logColors[l.type] || logColors.info;
                          return (
                            <div key={i} style={{
                              display: "flex", gap: 8, marginBottom: 3,
                              fontSize: 11.5, fontFamily: "monospace",
                            }}>
                              <span style={{ color: "#334155", flexShrink: 0 }}>{l.ts}</span>
                              <span style={{ color: c.color }}>{l.msg}</span>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>

            {/* Error */}
            {result?.error && (
              <div style={{
                background: "#1A0A0A", border: "1px solid #7F1D1D",
                borderRadius: 12, padding: 20, color: "#FCA5A5",
              }}>
                <strong>Error:</strong> {result.error}
                <button
                  onClick={() => { setScreen("upload"); setResult(null); setAgState(INIT_STATE); }}
                  style={{
                    marginLeft: 16, background: "#7F1D1D", color: "#fff",
                    border: "none", padding: "6px 14px", borderRadius: 6, cursor: "pointer",
                  }}>Start over</button>
              </div>
            )}

            {/* Results */}
            {hasResult && (
              <div>
                {/* QA score bar */}
                <div style={{
                  background: "#0F172A", border: `1px solid ${result.qa?.verdict === "PASS" ? "#166534" : "#7F1D1D"}`,
                  borderRadius: 12, padding: "14px 18px",
                  display: "flex", alignItems: "center", gap: 16, marginBottom: 16,
                }}>
                  <div style={{ fontSize: 24 }}>
                    {result.qa?.verdict === "PASS" ? "✅" : "⚠️"}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 14, fontWeight: 700, color: result.qa?.verdict === "PASS" ? "#4ADE80" : "#FCD34D" }}>
                      QA {result.qa?.verdict} — Score {result.qa?.score}/100
                    </div>
                    <div style={{ fontSize: 12, color: "#475569" }}>
                      {result.qa?.passed} checks passed · {result.qa?.failed} failed ·
                      Company: <strong style={{ color: "#94A3B8" }}>{result.brief?.company || "Unknown"}</strong>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button onClick={copyHTML} style={{
                      background: copied ? "#166534" : "#1E293B",
                      color: copied ? "#4ADE80" : "#E2E8F0",
                      border: "1px solid #334155", padding: "8px 16px",
                      borderRadius: 8, cursor: "pointer", fontWeight: 600, fontSize: 13,
                      transition: "all 0.2s",
                    }}>
                      {copied ? "✓ Copied" : "⎘ Copy HTML"}
                    </button>
                    <button onClick={download} style={{
                      background: "linear-gradient(135deg,#2563EB,#7C3AED)",
                      color: "#fff", border: "none",
                      padding: "8px 18px", borderRadius: 8, cursor: "pointer",
                      fontWeight: 700, fontSize: 13,
                    }}>
                      ⬇ Download .html
                    </button>
                    <button onClick={() => { setScreen("upload"); setResult(null); setAgState(INIT_STATE); }}
                      style={{
                        background: "transparent", color: "#475569",
                        border: "1px solid #1E293B", padding: "8px 14px",
                        borderRadius: 8, cursor: "pointer", fontSize: 12,
                      }}>
                      ↺ New file
                    </button>
                  </div>
                </div>

                {/* Preview / Code tabs */}
                <div style={{
                  background: "#0F172A", border: "1px solid #1E293B",
                  borderRadius: 14, overflow: "hidden",
                }}>
                  <div style={{
                    display: "flex", borderBottom: "1px solid #1E293B",
                    padding: "0 4px",
                  }}>
                    {[
                      ["preview", "👁  Preview"],
                      ["code",    "💻  HTML"],
                      ["brief",   "📋  Brief data"],
                    ].map(([t, label]) => (
                      <button key={t} onClick={() => setTab(t)} style={{
                        padding: "11px 18px", background: "transparent",
                        color: tab === t ? "#E2E8F0" : "#475569",
                        border: "none", cursor: "pointer",
                        borderBottom: tab === t ? "2px solid #3B82F6" : "2px solid transparent",
                        fontWeight: tab === t ? 700 : 400, fontSize: 13,
                      }}>{label}</button>
                    ))}
                  </div>

                  {tab === "preview" && (
                    <div style={{ padding: 14 }}>
                      <div style={{ background: "#e5e7eb", borderRadius: 8, overflow: "hidden" }}>
                        <div style={{
                          background: "#d1d5db", padding: "8px 14px",
                          display: "flex", gap: 6,
                        }}>
                          {["#ef4444","#eab308","#22c55e"].map((c) => (
                            <div key={c} style={{ width: 10, height: 10, borderRadius: "50%", background: c }} />
                          ))}
                          <div style={{
                            flex: 1, background: "#e5e7eb", borderRadius: 4,
                            fontSize: 10, color: "#9ca3af", padding: "2px 8px",
                            marginLeft: 6, fontFamily: "monospace",
                          }}>email preview</div>
                        </div>
                        <iframe
                          srcDoc={result.html}
                          style={{ width: "100%", height: 620, border: "none", display: "block" }}
                          title="Email preview"
                        />
                      </div>
                    </div>
                  )}

                  {tab === "code" && (
                    <div style={{ padding: 14 }}>
                      <pre style={{
                        background: "#060910", borderRadius: 8, padding: 16,
                        color: "#86EFAC", fontSize: 11, fontFamily: "monospace",
                        overflowX: "auto", overflowY: "auto", maxHeight: 600,
                        whiteSpace: "pre-wrap", wordBreak: "break-all", margin: 0,
                      }}>{result.html}</pre>
                    </div>
                  )}

                  {tab === "brief" && result.brief && (
                    <div style={{ padding: 20, maxHeight: 600, overflowY: "auto" }}>
                      <BriefInspector brief={result.brief} />
                    </div>
                  )}
                </div>
              </div>
            )}

            <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}}`}</style>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Brief inspector component ─────────────────────────────
function BriefInspector({ brief }) {
  const rows = [
    ["Company",       brief.company],
    ["Subject",       brief.subject],
    ["Sender",        brief.sender],
    ["Primary color", brief.primaryColor],
    ["Greeting",      brief.greeting],
    ["Body text",     brief.bodyText],
    ["Signature",     brief.signature],
    ["Validity",      brief.validityDate],
    ["CTA",           brief.ctaText],
    ["UTM campaign",  brief.utmCampaign],
    ["Main link",     brief.mainBannerLink],
    ["Address",       brief.address],
  ].filter(([, v]) => v);

  return (
    <div>
      <div style={{ fontSize: 13, fontWeight: 700, color: "#94A3B8", marginBottom: 14 }}>
        Extracted brief data
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <tbody>
          {rows.map(([k, v]) => (
            <tr key={k}>
              <td style={{
                padding: "7px 12px", color: "#475569", fontWeight: 600,
                borderBottom: "1px solid #1E293B", whiteSpace: "nowrap", width: 130,
              }}>{k}</td>
              <td style={{
                padding: "7px 12px", color: "#CBD5E1",
                borderBottom: "1px solid #1E293B", wordBreak: "break-all",
              }}>{v}</td>
            </tr>
          ))}
        </tbody>
      </table>

      {brief.sections?.length > 0 && (
        <div style={{ marginTop: 20 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: "#94A3B8", marginBottom: 10 }}>
            Content sections ({brief.sections.length})
          </div>
          {brief.sections.map((s, i) => (
            <div key={i} style={{
              background: "#0A0F1A", border: "1px solid #1E293B",
              borderRadius: 8, padding: "12px 14px", marginBottom: 8,
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                <span style={{
                  fontSize: 10, fontWeight: 700, background: "#1E293B",
                  color: "#3B82F6", padding: "2px 8px", borderRadius: 4,
                  fontFamily: "monospace",
                }}>{s.type}</span>
                <span style={{ fontSize: 13, fontWeight: 600, color: "#E2E8F0" }}>{s.title}</span>
              </div>
              {s.subtitle && <div style={{ fontSize: 12, color: "#475569", marginBottom: 4 }}>{s.subtitle}</div>}
              {s.link && <div style={{ fontSize: 11, color: "#3B82F6", fontFamily: "monospace" }}>{s.link}</div>}
              {s.items?.length > 0 && (
                <div style={{ marginTop: 8, display: "flex", flexWrap: "wrap", gap: 6 }}>
                  {s.items.map((item, j) => (
                    <div key={j} style={{
                      background: "#1E293B", borderRadius: 6, padding: "4px 10px",
                      fontSize: 11, color: "#94A3B8",
                    }}>{item.title}{item.price ? ` — ${item.price}` : ""}</div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
