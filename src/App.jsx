import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx-js-style";

// ══════════════════════════════════════════════════════════════
//  CONFIG MONDAY API
// ══════════════════════════════════════════════════════════════
const MONDAY_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjI0NzU0NjgyNSwiYWFpIjoxMSwidWlkIjo0MDI1Nzg1MiwiaWFkIjoiMjAyMy0wMy0yOFQxNzo0MTozMC4wMDBaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTU1OTc4MTUsInJnbiI6ImV1YzEifQ.E20krB-L2O750U1V1vmj5HqW2AsZKz40DXU16OHopGU";
const BOARD_ID = "5096517828"; // CONTRAT_GLOBAL

// Colonnes Monday (IDs fixes standardisés)
const COL = {
  cnss:    "ncnss",
  cin:     "cin",
  projet:  "projetmd",
  societe: "societe_col",
  poste:   "poste",
  nom:     "name",
  dateDebut: "datedebutcontrat",
};

// ══════════════════════════════════════════════════════════════
//  UTILS
// ══════════════════════════════════════════════════════════════
function normStr(s) {
  return String(s || "").toLowerCase().replace(/[àáâã]/g,"a").replace(/[éèêë]/g,"e")
    .replace(/[îï]/g,"i").replace(/[ôõ]/g,"o").replace(/[ùûü]/g,"u").replace(/[ç]/g,"c").trim();
}

function findColIndex(headers, keywords) {
  const norm = headers.map(normStr);
  for (const kw of keywords) {
    const idx = norm.findIndex(h => h.includes(normStr(kw)));
    if (idx !== -1) return idx;
  }
  return -1;
}

function extractCNSS(value) {
  const s = String(value || "").replace(/\s/g, "");
  if (/^\d{7,12}$/.test(s)) return s;
  return null;
}

function parsePeriod(s) {
  if (!s) return "";
  const m = String(s).match(/(\d{2})[\/\-](\d{4})/);
  if (m) return `${m[1]}/${m[2]}`;
  return String(s).trim();
}

// ══════════════════════════════════════════════════════════════
//  MONDAY API
// ══════════════════════════════════════════════════════════════
async function fetchMondayAgents(boardId, token) {
  const query = `query {
    boards(ids: [${boardId}]) {
      items_page(limit: 500) {
        items {
          id
          name
          column_values { id text }
        }
      }
    }
  }`;
  const r = await fetch("https://api.monday.com/v2", {
    method: "POST",
    headers: { "Content-Type": "application/json", "Authorization": token, "API-Version": "2024-01" },
    body: JSON.stringify({ query })
  });
  const d = await r.json();
  if (d.errors) throw new Error(d.errors[0].message);
  const items = d?.data?.boards?.[0]?.items_page?.items || [];
  return items.map(item => {
    const c = {};
    for (const col of item.column_values) c[col.id] = col.text || "";
    return {
      id:       item.id,
      nom:      item.name || c[COL.nom] || "",
      cnss:     c[COL.cnss]    || "",
      cin:      c[COL.cin]     || "",
      projet:   c[COL.projet]  || "",
      societe:  c[COL.societe] || "",
      poste:    c[COL.poste]   || "",
      dateDebut: c[COL.dateDebut] || "",
    };
  }).filter(a => a.cnss);
}

// ══════════════════════════════════════════════════════════════
//  BDS PARSER
// ══════════════════════════════════════════════════════════════
function parseBDSFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const rows = [];
        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
          if (data.length < 2) continue;
          // Find header row
          let headerIdx = 0;
          for (let i = 0; i < Math.min(10, data.length); i++) {
            const row = data[i].map(normStr);
            if (row.some(c => c.includes("immatricul") || c.includes("matricul") || c.includes("cnss"))) {
              headerIdx = i; break;
            }
          }
          const headers = data[headerIdx].map(String);
          const cnssIdx = findColIndex(headers, ["immatriculation","matricul","cnss","n°"]);
          const nomIdx  = findColIndex(headers, ["nom","prenom","salarie","name"]);
          const joursIdx = findColIndex(headers, ["jours","nb jour","nombre"]);
          const salaireIdx = findColIndex(headers, ["salaire","brut","montant"]);
          if (cnssIdx === -1) {
            // Try to detect by scanning values for CNSS-like numbers
            for (let r = headerIdx + 1; r < data.length; r++) {
              for (let c = 0; c < data[r].length; c++) {
                const cnss = extractCNSS(data[r][c]);
                if (cnss) {
                  rows.push({ cnss, nom: data[r][nomIdx] || "", row: data[r], headers, sheet: sheetName });
                  break;
                }
              }
            }
          } else {
            for (let r = headerIdx + 1; r < data.length; r++) {
              const cnss = extractCNSS(data[r][cnssIdx]);
              if (cnss) {
                rows.push({
                  cnss,
                  nom:     String(data[r][nomIdx]    || ""),
                  jours:   data[r][joursIdx]  || "",
                  salaire: data[r][salaireIdx] || "",
                  row:     data[r],
                  headers,
                  sheet:   sheetName
                });
              }
            }
          }
        }
        resolve({ rows, fileName: file.name });
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ══════════════════════════════════════════════════════════════
//  EXCEL EXPORT
// ══════════════════════════════════════════════════════════════
function exportProjectExcel(projet, societe, periode, found, notFound, bdsHeaders) {
  const wb = XLSX.utils.book_new();

  // ── Styles xlsx-js-style (format correct) ──────────────────
  const styleHeader = {
    font:  { bold: true, color: { rgb: "FFFFFF" }, sz: 11 },
    fill:  { patternType: "solid", fgColor: { rgb: "1A3A6B" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: {
      top:    { style: "thin", color: { rgb: "AAAAAA" } },
      bottom: { style: "thin", color: { rgb: "AAAAAA" } },
      left:   { style: "thin", color: { rgb: "AAAAAA" } },
      right:  { style: "thin", color: { rgb: "AAAAAA" } },
    }
  };

  const styleGreen = {
    fill:  { patternType: "solid", fgColor: { rgb: "DCFCE7" } },
    font:  { color: { rgb: "166534" } },
    border: {
      top:    { style: "thin", color: { rgb: "BBBBBB" } },
      bottom: { style: "thin", color: { rgb: "BBBBBB" } },
      left:   { style: "thin", color: { rgb: "BBBBBB" } },
      right:  { style: "thin", color: { rgb: "BBBBBB" } },
    }
  };

  const styleRed = {
    fill:  { patternType: "solid", fgColor: { rgb: "FEE2E2" } },
    font:  { color: { rgb: "991B1B" } },
    border: {
      top:    { style: "thin", color: { rgb: "BBBBBB" } },
      bottom: { style: "thin", color: { rgb: "BBBBBB" } },
      left:   { style: "thin", color: { rgb: "BBBBBB" } },
      right:  { style: "thin", color: { rgb: "BBBBBB" } },
    }
  };

  const styleRecapLabel = {
    font:  { bold: true, color: { rgb: "1A3A6B" } },
    fill:  { patternType: "solid", fgColor: { rgb: "EBF5FB" } },
  };

  const styleRecapTitle = {
    font:  { bold: true, sz: 13, color: { rgb: "FFFFFF" } },
    fill:  { patternType: "solid", fgColor: { rgb: "1A3A6B" } },
    alignment: { horizontal: "center" },
  };

  // ── Helper : appliquer un style à toute une ligne ──────────
  function applyRowStyle(ws, rowIdx, numCols, style) {
    for (let c = 0; c < numCols; c++) {
      const ref = XLSX.utils.encode_cell({ r: rowIdx, c });
      if (!ws[ref]) ws[ref] = { v: "", t: "s" };
      ws[ref].s = style;
    }
  }

  // ── FEUILLE 1 : Agents Trouvés ─────────────────────────────
  const headers1 = bdsHeaders.length > 0
    ? [...bdsHeaders, "Statut"]
    : ["N° CNSS", "Nom et Prénom", "Nb Jours", "Salaire Brut", "Statut"];

  const foundRows = found.map(a => [
    ...(a.row || [a.cnss, a.nom, a.jours || "", a.salaire || ""]),
    "DÉCLARÉ ✓"
  ]);

  const ws1 = XLSX.utils.aoa_to_sheet([headers1, ...foundRows]);

  // Style header ligne 0
  applyRowStyle(ws1, 0, headers1.length, styleHeader);
  // Style vert lignes données
  for (let r = 1; r <= foundRows.length; r++) {
    applyRowStyle(ws1, r, headers1.length, styleGreen);
  }
  // Largeur colonnes
  ws1["!cols"] = headers1.map(() => ({ wch: 22 }));

  XLSX.utils.book_append_sheet(wb, ws1, "Agents Trouvés");

  // ── FEUILLE 2 : Agents Introuvables ───────────────────────
  const headers2 = ["N° CNSS", "Nom / Prénom", "Projet", "Société", "Poste", "Statut"];
  const notFoundRows = notFound.map(a => [
    a.cnss, a.nom, a.projet, a.societe, a.poste || "", "INTROUVABLE ✗"
  ]);

  const ws2 = XLSX.utils.aoa_to_sheet([headers2, ...notFoundRows]);

  applyRowStyle(ws2, 0, headers2.length, styleHeader);
  for (let r = 1; r <= notFoundRows.length; r++) {
    applyRowStyle(ws2, r, headers2.length, styleRed);
  }
  ws2["!cols"] = headers2.map(() => ({ wch: 22 }));

  XLSX.utils.book_append_sheet(wb, ws2, "Agents Introuvables");

  // ── FEUILLE 3 : Récapitulatif ──────────────────────────────
  const taux = Math.round(found.length / (found.length + notFound.length) * 100);
  const recapRows = [
    ["RÉCAPITULATIF"],
    ["Projet",                    projet],
    ["Société",                   societe],
    ["Période",                   periode],
    ["Total agents liste",        found.length + notFound.length],
    ["Agents déclarés (trouvés)", found.length],
    ["Agents introuvables",       notFound.length],
    ["Taux de couverture",        `${taux}%`],
  ];

  const ws3 = XLSX.utils.aoa_to_sheet(recapRows);

  // Titre en bleu marine
  if (ws3["A1"]) ws3["A1"].s = styleRecapTitle;
  // Labels en bleu clair
  for (let r = 1; r < recapRows.length; r++) {
    const refA = XLSX.utils.encode_cell({ r, c: 0 });
    if (ws3[refA]) ws3[refA].s = styleRecapLabel;
  }
  ws3["!cols"] = [{ wch: 30 }, { wch: 25 }];
  ws3["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }];

  XLSX.utils.book_append_sheet(wb, ws3, "Récapitulatif");

  // ── Export ─────────────────────────────────────────────────
  const fileName = `${projet.replace(/\s+/g,"_")}_${periode.replace("/","-")}.xlsx`;
  XLSX.writeFile(wb, fileName);
}

// ══════════════════════════════════════════════════════════════
//  MAIN APP COMPONENT
// ══════════════════════════════════════════════════════════════
export default function CNSSRapprochement() {
  const [bdsFiles, setBdsFiles] = useState([]);
  const [agents, setAgents] = useState([]);
  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("");
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("bds");
  const [manualAgents, setManualAgents] = useState("");
  const [personnelSource, setPersonnelSource] = useState("monday"); // "monday" | "excel" | "manual"
  const bdsInputRef = useRef();
  const personnelInputRef = useRef();

  // ── BDS File Upload ────────────────────────────────────────
  const handleBDSDrop = useCallback(async (files) => {
    setError("");
    const newFiles = [];
    for (const file of files) {
      try {
        setLoadingMsg(`Lecture ${file.name}...`);
        const parsed = await parseBDSFile(file);
        newFiles.push({
          ...parsed,
          id: Date.now() + Math.random(),
          societeTag: "",
          periodeTag: "",
        });
      } catch (e) {
        setError(`Erreur lecture ${file.name}: ${e.message}`);
      }
    }
    setBdsFiles(prev => [...prev, ...newFiles]);
    setLoadingMsg("");
  }, []);

  const handleBDSInput = (e) => {
    if (e.target.files?.length) handleBDSDrop(Array.from(e.target.files));
  };

  const updateBDSTag = (id, field, value) => {
    setBdsFiles(prev => prev.map(f => f.id === id ? { ...f, [field]: value } : f));
  };

  const removeBDS = (id) => setBdsFiles(prev => prev.filter(f => f.id !== id));

  // ── Personnel: fetch from Monday ───────────────────────────
  const fetchFromMonday = async () => {
    setLoading(true);
    setLoadingMsg("Récupération agents Monday.com...");
    setError("");
    try {
      const data = await fetchMondayAgents(BOARD_ID, MONDAY_TOKEN);
      setAgents(data);
      setLoadingMsg(`✅ ${data.length} agents chargés depuis Monday`);
      setTimeout(() => setLoadingMsg(""), 3000);
    } catch (e) {
      setError(`Erreur Monday API: ${e.message}`);
      setLoadingMsg("");
    }
    setLoading(false);
  };

  // ── Personnel: import from Excel ───────────────────────────
  const handlePersonnelExcel = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (data.length < 2) return;
        const headers = data[0].map(normStr);
        const cnssIdx    = findColIndex(headers, ["cnss","immatricul","matricul"]);
        const nomIdx     = findColIndex(headers, ["nom","prenom","name"]);
        const projetIdx  = findColIndex(headers, ["projet","chantier","marche"]);
        const societeIdx = findColIndex(headers, ["societe","entreprise","societe"]);
        const dateIdx    = findColIndex(headers, ["date"]);
        const parsed = [];
        for (let r = 1; r < data.length; r++) {
          const row = data[r];
          const cnss = extractCNSS(row[cnssIdx]);
          if (!cnss) continue;
          parsed.push({
            cnss,
            nom:      String(row[nomIdx]     || ""),
            projet:   String(row[projetIdx]  || ""),
            societe:  String(row[societeIdx] || ""),
            dateDebut: parsePeriod(String(row[dateIdx] || "")),
            poste: "",
          });
        }
        setAgents(parsed);
        setLoadingMsg(`✅ ${parsed.length} agents importés`);
        setTimeout(() => setLoadingMsg(""), 3000);
      } catch (e) {
        setError("Erreur import Excel personnel: " + e.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // ── ANALYSE ────────────────────────────────────────────────
  const runAnalysis = () => {
    setError("");
    if (bdsFiles.length === 0) { setError("Veuillez uploader au moins un fichier BDS."); return; }
    if (agents.length === 0)   { setError("Veuillez charger la liste du personnel."); return; }

    // Grouper agents par projet + société + période
    const groups = {};
    for (const agent of agents) {
      const key = `${agent.projet}||${agent.societe}||${parsePeriod(agent.dateDebut)}`;
      if (!groups[key]) groups[key] = { projet: agent.projet, societe: agent.societe, periode: parsePeriod(agent.dateDebut), agents: [] };
      groups[key].agents.push(agent);
    }

    const res = [];
    for (const group of Object.values(groups)) {
      if (!group.projet) continue;

      // Trouver le(s) fichier(s) BDS correspondant
      const matchingBDS = bdsFiles.filter(f => {
        const socMatch = !f.societeTag || normStr(f.societeTag).includes(normStr(group.societe)) || normStr(group.societe).includes(normStr(f.societeTag));
        const perMatch = !f.periodeTag || f.periodeTag === group.periode || !group.periode;
        return socMatch && perMatch;
      });

      const usedBDS = matchingBDS.length > 0 ? matchingBDS : bdsFiles;
      const bdsHeaders = usedBDS[0]?.rows?.[0]?.headers || [];

      // Construire set CNSS du BDS
      const cnssSet = new Set();
      const cnssRowMap = {};
      for (const bds of usedBDS) {
        for (const row of bds.rows) {
          cnssSet.add(row.cnss);
          cnssRowMap[row.cnss] = row;
        }
      }

      const found = [];
      const notFound = [];
      for (const agent of group.agents) {
        const normalizedCNSS = agent.cnss.replace(/\s/g, "");
        if (cnssSet.has(normalizedCNSS)) {
          found.push({ ...agent, ...cnssRowMap[normalizedCNSS] });
        } else {
          notFound.push(agent);
        }
      }

      res.push({
        id: Date.now() + Math.random(),
        projet:  group.projet,
        societe: group.societe,
        periode: group.periode,
        total:   group.agents.length,
        found,
        notFound,
        bdsHeaders,
        taux: Math.round(found.length / group.agents.length * 100),
      });
    }

    setResults(res);
    setActiveTab("results");
  };

  // ══════════════════════════════════════════════════════════
  //  RENDER
  // ══════════════════════════════════════════════════════════
  const totalAgents   = results.reduce((s, r) => s + r.total, 0);
  const totalFound    = results.reduce((s, r) => s + r.found.length, 0);
  const totalNotFound = results.reduce((s, r) => s + r.notFound.length, 0);
  const tauxGlobal    = totalAgents > 0 ? Math.round(totalFound / totalAgents * 100) : 0;

  return (
    <div style={{ fontFamily: "'Segoe UI', system-ui, sans-serif", background: "#0f172a", minHeight: "100vh", color: "#e2e8f0" }}>

      {/* HEADER */}
      <div style={{ background: "linear-gradient(135deg, #1e293b 0%, #0f172a 100%)", borderBottom: "1px solid #334155", padding: "20px 32px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <div style={{ width: 40, height: 40, borderRadius: 10, background: "linear-gradient(135deg, #3b82f6, #1d4ed8)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 20 }}>⚖️</div>
            <div>
              <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: "#f8fafc" }}>Rapprochement CNSS</h1>
              <div style={{ fontSize: 12, color: "#64748b" }}>GROUPE KIRKOS · Outil de vérification des déclarations</div>
            </div>
          </div>
        </div>
        {loadingMsg && <div style={{ background: "#1e3a5f", border: "1px solid #3b82f6", borderRadius: 8, padding: "8px 16px", fontSize: 13, color: "#93c5fd" }}>{loadingMsg}</div>}
      </div>

      {/* TABS */}
      <div style={{ display: "flex", gap: 0, borderBottom: "1px solid #1e293b", background: "#0f172a", padding: "0 32px" }}>
        {[
          { id: "bds",        label: "1 — Fichiers BDS", icon: "📁" },
          { id: "personnel",  label: "2 — Personnel",    icon: "👥" },
          { id: "results",    label: "3 — Résultats",    icon: "📊", badge: results.length },
        ].map(tab => (
          <button key={tab.id} onClick={() => setActiveTab(tab.id)} style={{
            background: "none", border: "none", padding: "14px 20px", cursor: "pointer",
            color: activeTab === tab.id ? "#3b82f6" : "#64748b",
            borderBottom: activeTab === tab.id ? "2px solid #3b82f6" : "2px solid transparent",
            fontSize: 14, fontWeight: activeTab === tab.id ? 600 : 400,
            display: "flex", alignItems: "center", gap: 8, transition: "all 0.2s",
          }}>
            {tab.icon} {tab.label}
            {tab.badge > 0 && <span style={{ background: "#3b82f6", color: "white", borderRadius: 999, fontSize: 11, padding: "1px 7px" }}>{tab.badge}</span>}
          </button>
        ))}
      </div>

      <div style={{ padding: "28px 32px", maxWidth: 1100, margin: "0 auto" }}>

        {error && (
          <div style={{ background: "#450a0a", border: "1px solid #ef4444", borderRadius: 10, padding: "12px 16px", marginBottom: 20, color: "#fca5a5", fontSize: 14, display: "flex", gap: 10 }}>
            ❌ {error}
          </div>
        )}

        {/* ════ ONGLET BDS ════ */}
        {activeTab === "bds" && (
          <div>
            <h2 style={{ color: "#f1f5f9", marginBottom: 6, fontSize: 18 }}>Fichiers BDS / État 45 CNSS</h2>
            <p style={{ color: "#64748b", fontSize: 14, marginBottom: 20 }}>Uploadez un ou plusieurs fichiers Excel exportés depuis Damancom. Taguez chaque fichier avec sa société et sa période.</p>

            {/* Drop zone */}
            <div
              onClick={() => bdsInputRef.current?.click()}
              onDragOver={e => e.preventDefault()}
              onDrop={e => { e.preventDefault(); handleBDSDrop(Array.from(e.dataTransfer.files)); }}
              style={{
                border: "2px dashed #334155", borderRadius: 14, padding: "40px 20px",
                textAlign: "center", cursor: "pointer", marginBottom: 20,
                background: "#1e293b", transition: "border-color 0.2s",
              }}
            >
              <div style={{ fontSize: 40, marginBottom: 10 }}>📂</div>
              <div style={{ color: "#94a3b8", fontSize: 15 }}>Glisser-déposer ou <span style={{ color: "#3b82f6", textDecoration: "underline" }}>parcourir</span></div>
              <div style={{ color: "#475569", fontSize: 12, marginTop: 6 }}>Formats acceptés : .xlsx, .xls, .csv</div>
              <input ref={bdsInputRef} type="file" accept=".xlsx,.xls,.csv" multiple style={{ display: "none" }} onChange={handleBDSInput} />
            </div>

            {/* Files list */}
            {bdsFiles.map(f => (
              <div key={f.id} style={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 12, padding: "16px 20px", marginBottom: 12, display: "flex", alignItems: "center", gap: 16 }}>
                <div style={{ fontSize: 28 }}>📊</div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 600, color: "#e2e8f0", fontSize: 14 }}>{f.fileName}</div>
                  <div style={{ color: "#22c55e", fontSize: 13 }}>{f.rows.length} agents détectés</div>
                </div>
                <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                  <input
                    value={f.societeTag}
                    onChange={e => updateBDSTag(f.id, "societeTag", e.target.value)}
                    placeholder="Société (tag)"
                    style={{ background: "#0f172a", border: "1px solid #475569", borderRadius: 8, padding: "6px 12px", color: "#e2e8f0", fontSize: 13, width: 160 }}
                  />
                  <input
                    value={f.periodeTag}
                    onChange={e => updateBDSTag(f.id, "periodeTag", e.target.value)}
                    placeholder="Période MM/AAAA"
                    style={{ background: "#0f172a", border: "1px solid #475569", borderRadius: 8, padding: "6px 12px", color: "#e2e8f0", fontSize: 13, width: 130 }}
                  />
                  <button onClick={() => removeBDS(f.id)} style={{ background: "#450a0a", border: "none", borderRadius: 8, padding: "6px 12px", color: "#f87171", cursor: "pointer", fontSize: 16 }}>✕</button>
                </div>
              </div>
            ))}

            {bdsFiles.length === 0 && (
              <div style={{ textAlign: "center", color: "#475569", padding: "20px 0", fontSize: 14 }}>
                Aucun fichier BDS chargé
              </div>
            )}
          </div>
        )}

        {/* ════ ONGLET PERSONNEL ════ */}
        {activeTab === "personnel" && (
          <div>
            <h2 style={{ color: "#f1f5f9", marginBottom: 6, fontSize: 18 }}>Liste du Personnel</h2>
            <p style={{ color: "#64748b", fontSize: 14, marginBottom: 20 }}>Source des agents à rapprocher avec le BDS.</p>

            {/* Source selector */}
            <div style={{ display: "flex", gap: 10, marginBottom: 24 }}>
              {[
                { id: "monday", label: "Monday.com", icon: "🔗", desc: "CONTRAT_GLOBAL" },
                { id: "excel",  label: "Import Excel", icon: "📋", desc: "Fichier .xlsx" },
              ].map(src => (
                <button key={src.id} onClick={() => setPersonnelSource(src.id)} style={{
                  flex: 1, background: personnelSource === src.id ? "#1e3a5f" : "#1e293b",
                  border: `2px solid ${personnelSource === src.id ? "#3b82f6" : "#334155"}`,
                  borderRadius: 12, padding: "16px 20px", cursor: "pointer", textAlign: "left", transition: "all 0.2s",
                }}>
                  <div style={{ fontSize: 22, marginBottom: 6 }}>{src.icon}</div>
                  <div style={{ fontWeight: 600, color: "#e2e8f0", fontSize: 14 }}>{src.label}</div>
                  <div style={{ color: "#64748b", fontSize: 12 }}>{src.desc}</div>
                </button>
              ))}
            </div>

            {personnelSource === "monday" && (
              <div style={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 12, padding: 24 }}>
                <div style={{ marginBottom: 16, color: "#94a3b8", fontSize: 14 }}>
                  Tableau : <span style={{ color: "#3b82f6" }}>CONTRAT_GLOBAL</span> (ID: {BOARD_ID})<br/>
                  Colonnes utilisées : <code style={{ color: "#a78bfa" }}>ncnss, projetmd, societe_col, poste</code>
                </div>
                <button onClick={fetchFromMonday} disabled={loading} style={{
                  background: "linear-gradient(135deg, #3b82f6, #1d4ed8)", border: "none",
                  borderRadius: 10, padding: "12px 28px", color: "white", fontSize: 14,
                  fontWeight: 600, cursor: loading ? "not-allowed" : "pointer", opacity: loading ? 0.7 : 1,
                }}>
                  {loading ? "⏳ Chargement..." : "🔄 Charger depuis Monday"}
                </button>
                {agents.length > 0 && (
                  <div style={{ marginTop: 16, color: "#22c55e", fontSize: 14 }}>
                    ✅ {agents.length} agents chargés — {[...new Set(agents.map(a => a.projet).filter(Boolean))].length} projets
                  </div>
                )}
              </div>
            )}

            {personnelSource === "excel" && (
              <div style={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 12, padding: 24 }}>
                <div style={{ marginBottom: 16, color: "#94a3b8", fontSize: 14 }}>
                  Colonnes attendues : <span style={{ color: "#fbbf24" }}>CNSS · Nom · Projet · Société · Date début</span>
                </div>
                <button onClick={() => personnelInputRef.current?.click()} style={{
                  background: "#1e3a5f", border: "1px solid #3b82f6", borderRadius: 10,
                  padding: "12px 28px", color: "#93c5fd", fontSize: 14, fontWeight: 600, cursor: "pointer",
                }}>
                  📁 Importer Excel personnel
                </button>
                <input ref={personnelInputRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={handlePersonnelExcel} />
                {agents.length > 0 && (
                  <div style={{ marginTop: 12, color: "#22c55e", fontSize: 14 }}>
                    ✅ {agents.length} agents importés
                  </div>
                )}
              </div>
            )}

            {/* Preview table */}
            {agents.length > 0 && (
              <div style={{ marginTop: 24 }}>
                <h3 style={{ color: "#94a3b8", fontSize: 14, marginBottom: 12 }}>Aperçu — {agents.length} agents</h3>
                <div style={{ overflowX: "auto", borderRadius: 10, border: "1px solid #334155" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
                    <thead>
                      <tr style={{ background: "#1e293b" }}>
                        {["N° CNSS","Nom / Prénom","Projet","Société","Date Début"].map(h => (
                          <th key={h} style={{ padding: "10px 14px", textAlign: "left", color: "#64748b", fontWeight: 600, borderBottom: "1px solid #334155" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {agents.slice(0, 8).map((a, i) => (
                        <tr key={i} style={{ borderBottom: "1px solid #1e293b" }}>
                          <td style={{ padding: "9px 14px", color: "#3b82f6", fontFamily: "monospace" }}>{a.cnss}</td>
                          <td style={{ padding: "9px 14px", color: "#e2e8f0" }}>{a.nom}</td>
                          <td style={{ padding: "9px 14px", color: "#a78bfa" }}>{a.projet}</td>
                          <td style={{ padding: "9px 14px", color: "#94a3b8" }}>{a.societe}</td>
                          <td style={{ padding: "9px 14px", color: "#64748b" }}>{a.dateDebut}</td>
                        </tr>
                      ))}
                      {agents.length > 8 && (
                        <tr>
                          <td colSpan={5} style={{ padding: "9px 14px", color: "#475569", textAlign: "center", fontSize: 12 }}>
                            ... et {agents.length - 8} autres agents
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Launch analysis */}
            <div style={{ marginTop: 28, display: "flex", justifyContent: "flex-end" }}>
              <button onClick={runAnalysis} disabled={bdsFiles.length === 0 || agents.length === 0} style={{
                background: "linear-gradient(135deg, #059669, #047857)", border: "none",
                borderRadius: 12, padding: "14px 36px", color: "white", fontSize: 15,
                fontWeight: 700, cursor: bdsFiles.length === 0 || agents.length === 0 ? "not-allowed" : "pointer",
                opacity: bdsFiles.length === 0 || agents.length === 0 ? 0.5 : 1,
                boxShadow: "0 4px 20px rgba(5,150,105,0.3)",
              }}>
                🔍 Lancer l'analyse
              </button>
            </div>
          </div>
        )}

        {/* ════ ONGLET RÉSULTATS ════ */}
        {activeTab === "results" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 24 }}>
              <div>
                <h2 style={{ color: "#f1f5f9", marginBottom: 4, fontSize: 18 }}>Résultats du Rapprochement</h2>
                <p style={{ color: "#64748b", fontSize: 14 }}>{results.length} projet(s) analysé(s)</p>
              </div>
              {results.length > 0 && (
                <button onClick={runAnalysis} style={{ background: "#1e293b", border: "1px solid #475569", borderRadius: 10, padding: "9px 18px", color: "#94a3b8", fontSize: 13, cursor: "pointer" }}>
                  🔄 Relancer
                </button>
              )}
            </div>

            {/* KPIs */}
            {results.length > 0 && (
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 28 }}>
                {[
                  { label: "Total vérifiés",   value: totalAgents,   color: "#3b82f6", bg: "#1e3a5f",  icon: "👤" },
                  { label: "Déclarés",          value: totalFound,    color: "#22c55e", bg: "#14532d",  icon: "✅" },
                  { label: "Introuvables",      value: totalNotFound, color: "#ef4444", bg: "#450a0a",  icon: "❌" },
                  { label: "Taux couverture",   value: `${tauxGlobal}%`, color: tauxGlobal >= 95 ? "#22c55e" : "#f59e0b", bg: tauxGlobal >= 95 ? "#14532d" : "#451a03", icon: tauxGlobal >= 95 ? "🟢" : "🟡" },
                ].map(kpi => (
                  <div key={kpi.label} style={{ background: kpi.bg, border: `1px solid ${kpi.color}33`, borderRadius: 12, padding: "18px 20px" }}>
                    <div style={{ fontSize: 24, marginBottom: 6 }}>{kpi.icon}</div>
                    <div style={{ fontSize: 26, fontWeight: 700, color: kpi.color }}>{kpi.value}</div>
                    <div style={{ fontSize: 12, color: "#64748b", marginTop: 4 }}>{kpi.label}</div>
                  </div>
                ))}
              </div>
            )}

            {results.length === 0 ? (
              <div style={{ textAlign: "center", padding: "60px 20px", color: "#475569" }}>
                <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
                <div style={{ fontSize: 16 }}>Lancez l'analyse depuis l'onglet Personnel</div>
              </div>
            ) : (
              results.map(r => (
                <div key={r.id} style={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 14, marginBottom: 16, overflow: "hidden" }}>
                  {/* Project header */}
                  <div style={{ padding: "16px 20px", background: "#162032", borderBottom: "1px solid #334155", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                      <div style={{ fontSize: 24 }}>🏗️</div>
                      <div>
                        <div style={{ fontWeight: 700, color: "#f1f5f9", fontSize: 15 }}>{r.projet}</div>
                        <div style={{ color: "#64748b", fontSize: 12 }}>{r.societe} · {r.periode}</div>
                      </div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
                      <div style={{ textAlign: "center" }}>
                        <div style={{ fontWeight: 700, color: "#22c55e", fontSize: 18 }}>{r.found.length}/{r.total}</div>
                        <div style={{ color: "#64748b", fontSize: 11 }}>Trouvés</div>
                      </div>
                      <div style={{ textAlign: "center" }}>
                        <div style={{ fontWeight: 700, color: r.taux >= 95 ? "#22c55e" : "#f59e0b", fontSize: 18 }}>{r.taux}%</div>
                        <div style={{ color: "#64748b", fontSize: 11 }}>Couverture</div>
                      </div>
                      <button
                        onClick={() => exportProjectExcel(r.projet, r.societe, r.periode, r.found, r.notFound, r.bdsHeaders)}
                        style={{ background: "linear-gradient(135deg, #059669, #047857)", border: "none", borderRadius: 10, padding: "10px 18px", color: "white", fontSize: 13, fontWeight: 600, cursor: "pointer" }}
                      >
                        ⬇️ Télécharger BDS filtré
                      </button>
                    </div>
                  </div>

                  {/* Not found agents */}
                  {r.notFound.length > 0 && (
                    <div style={{ padding: "14px 20px" }}>
                      <div style={{ color: "#f87171", fontSize: 13, fontWeight: 600, marginBottom: 8 }}>
                        ⚠️ {r.notFound.length} agent(s) non déclaré(s) dans le BDS :
                      </div>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                        {r.notFound.map((a, i) => (
                          <span key={i} style={{ background: "#450a0a", border: "1px solid #ef4444", borderRadius: 6, padding: "4px 10px", fontSize: 12, color: "#fca5a5" }}>
                            {a.cnss} — {a.nom}
                          </span>
                        ))}
                      </div>
                    </div>
                  )}
                  {r.notFound.length === 0 && (
                    <div style={{ padding: "12px 20px", color: "#22c55e", fontSize: 13 }}>
                      ✅ Tous les agents sont déclarés dans le BDS
                    </div>
                  )}
                </div>
              ))
            )}
          </div>
        )}
      </div>
    </div>
  );
}
