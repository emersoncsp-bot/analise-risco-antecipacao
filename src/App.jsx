import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ *
 *  Utilitários
 * ------------------------------------------------------------------ */

// Normaliza textos para comparação (sem acento, minúsculo, só alfanumérico)
const norm = (s) =>
  String(s ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const trim = (v) => String(v ?? "").trim();

// Procura o índice de uma coluna dado um conjunto de candidatos (ordem = prioridade)
function findCol(headers, candidates) {
  const nh = headers.map(norm);
  for (const c of candidates) {
    const i = nh.indexOf(c);
    if (i !== -1) return i;
  }
  for (const c of candidates) {
    const i = nh.findIndex((h) => h && h.startsWith(c));
    if (i !== -1) return i;
  }
  for (const c of candidates) {
    const i = nh.findIndex((h) => h && h.includes(c));
    if (i !== -1) return i;
  }
  return -1;
}

// Lê o primeiro sheet como matriz de linhas (array de arrays)
function readSheetMatrix(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false, blankrows: false });
}

// Detecta a linha de cabeçalho (a que melhor casa com os candidatos esperados)
function detectHeaderRow(rows, candidateSets) {
  let best = -1;
  let bestRow = 0;
  const scan = Math.min(rows.length, 30);
  for (let r = 0; r < scan; r++) {
    const nh = (rows[r] || []).map(norm);
    let score = 0;
    for (const cands of candidateSets) {
      if (cands.some((c) => nh.some((h) => h && (h === c || h.startsWith(c) || h.includes(c))))) score++;
    }
    if (score > best) {
      best = score;
      bestRow = r;
    }
  }
  return bestRow;
}

/* ------------------------------------------------------------------ *
 *  Definição das colunas
 * ------------------------------------------------------------------ */

const PROD_COLS = [
  { key: "pedido", label: "Pedido/item", cands: ["pedidoitem", "pedido"] },
  { key: "ordem", label: "Ordem", cands: ["ordem"] },
  { key: "halb", label: "Halb", cands: ["halb"] },
  { key: "inicio", label: "Início (Enfornam.)", cands: ["inicioenfornam", "inicioenforn", "enfornamento", "enfornam", "inicio"] },
  { key: "cliente", label: "Cliente externo", cands: ["clienteexterno", "cliente"] },
  { key: "emi", label: "Teste EMI", cands: ["testeemi", "emi"] },
  { key: "descricao", label: "Descrição produto", cands: ["descricaoproduto", "descricaodoproduto", "descricao", "produto"] },
];

const AVAIL_COLS = [
  { key: "tuboUt", label: "Tubo Padrão UT", cands: ["tubopadraout", "tubopadraoultrassom", "tubout"] },
  { key: "tuboEmi", label: "Tubo Padrão EMI", cands: ["tubopadraoemi", "tuboemi"] },
  { key: "drift", label: "Drift", cands: ["drift"] },
  { key: "sapataUt", label: "Sapata UT", cands: ["sapataut", "sapata"] },
  { key: "abendi", label: "Qualificação ABENDI (para NDT)", cands: ["qualificacaoabendiparandt", "qualificacaoabendi", "abendi"] },
];

const AVAIL_KEY_CANDS = ["halbq", "halb"];

/* ------------------------------------------------------------------ *
 *  Parsers
 * ------------------------------------------------------------------ */

function parseProduction(arrayBuffer) {
  const rows = readSheetMatrix(arrayBuffer);
  const headerRow = detectHeaderRow(rows, PROD_COLS.map((c) => c.cands));
  const headers = rows[headerRow] || [];
  const colIdx = {};
  PROD_COLS.forEach((c) => (colIdx[c.key] = findCol(headers, c.cands)));

  const items = [];
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const obj = {};
    let hasContent = false;
    PROD_COLS.forEach((c) => {
      const v = colIdx[c.key] >= 0 ? trim(row[colIdx[c.key]]) : "";
      obj[c.key] = v;
      if (v) hasContent = true;
    });
    if (hasContent) items.push(obj);
  }
  const missing = PROD_COLS.filter((c) => colIdx[c.key] < 0).map((c) => c.label);
  return { items, missing };
}

function parseAvailability(arrayBuffer) {
  const rows = readSheetMatrix(arrayBuffer);
  const headerRow = detectHeaderRow(rows, [AVAIL_KEY_CANDS, ...AVAIL_COLS.map((c) => c.cands)]);
  const headers = rows[headerRow] || [];
  const keyIdx = findCol(headers, AVAIL_KEY_CANDS);
  const colIdx = {};
  AVAIL_COLS.forEach((c) => (colIdx[c.key] = findCol(headers, c.cands)));

  const map = new Map();
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const keyRaw = keyIdx >= 0 ? trim(row[keyIdx]) : "";
    if (!keyRaw) continue;
    const key = norm(keyRaw);
    const obj = {};
    AVAIL_COLS.forEach((c) => {
      obj[c.key] = colIdx[c.key] >= 0 ? trim(row[colIdx[c.key]]) : "";
    });
    if (!map.has(key)) map.set(key, obj);
  }
  const missing = [];
  if (keyIdx < 0) missing.push("HALBQ (chave)");
  AVAIL_COLS.filter((c) => colIdx[c.key] < 0).forEach((c) => missing.push(c.label));
  return { map, count: map.size, missing };
}

// Uma célula é "ruim" se está em branco ou marcada como indisponível
const isBad = (v) => {
  const t = trim(v);
  if (!t) return true;
  return norm(t).includes("indisponivel");
};

/* ------------------------------------------------------------------ *
 *  Componente de Upload
 * ------------------------------------------------------------------ */

function UploadZone({ title, subtitle, file, accentColor, status, onFile }) {
  const [drag, setDrag] = useState(false);
  const [hover, setHover] = useState(false);
  const inputId = useMemo(() => "f_" + Math.random().toString(36).slice(2), []);

  const handle = (f) => f && onFile(f);
  const active = drag || hover; // borda azul / realce ao passar o mouse ou arrastar

  return (
    <label
      htmlFor={inputId}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      onDragOver={(e) => {
        e.preventDefault();
        setDrag(true);
      }}
      onDragLeave={() => setDrag(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDrag(false);
        handle(e.dataTransfer.files?.[0]);
      }}
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 8,
        flex: 1,
        minWidth: 260,
        cursor: "pointer",
        background: active ? "rgba(37,99,235,0.04)" : "#fff",
        border: `1.5px dashed ${active ? "#2563EB" : "#CBD5E1"}`,
        borderRadius: 14,
        padding: "20px 22px",
        transition: "all .18s ease",
        boxShadow: active
          ? "0 6px 20px rgba(37,99,235,0.16), 0 0 0 3px rgba(37,99,235,0.10)"
          : "0 1px 2px rgba(15,23,42,.04)",
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <span
          style={{
            width: 34,
            height: 34,
            borderRadius: 9,
            background: `${accentColor}15`,
            color: accentColor,
            display: "grid",
            placeItems: "center",
            flexShrink: 0,
          }}
        >
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
            <path d="M14 2v6h6" />
            <path d="M12 18v-6M9 15l3-3 3 3" />
          </svg>
        </span>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14.5, color: "#0F172A", lineHeight: 1.2 }}>{title}</div>
          <div style={{ fontSize: 12, color: "#64748B" }}>{subtitle}</div>
        </div>
      </div>

      <div style={{ fontSize: 12.5, color: file ? "#2563EB" : "#94A3B8", marginTop: 2 }}>
        {file ? (
          <span style={{ color: "#2563EB", fontWeight: 600 }}>✓ {file}{status ? ` — ${status}` : ""}</span>
        ) : (
          "Arraste o arquivo aqui ou clique para selecionar (.xlsx, .xls, .csv)"
        )}
      </div>

      <input
        id={inputId}
        type="file"
        accept=".xlsx,.xls,.csv,.tsv,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv"
        style={{ display: "none" }}
        onChange={(e) => handle(e.target.files?.[0])}
      />
    </label>
  );
}

/* ------------------------------------------------------------------ *
 *  App principal
 * ------------------------------------------------------------------ */

export default function App() {
  const [prod, setProd] = useState(null); // { items, missing, fileName }
  const [avail, setAvail] = useState(null); // { map, count, missing, fileName }
  const [error, setError] = useState("");
  const [view, setView] = useState("import"); // "import" | "summary"

  const onProd = useCallback(async (file) => {
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseProduction(buf);
      setProd({ ...parsed, fileName: file.name });
    } catch (e) {
      setError("Não foi possível ler a planilha de Sequência de Produção. Verifique o formato.");
    }
  }, []);

  const onAvail = useCallback(async (file) => {
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const parsed = parseAvailability(buf);
      setAvail({ ...parsed, fileName: file.name });
    } catch (e) {
      setError("Não foi possível ler a planilha de Disponibilidade de Instrumentos. Verifique o formato.");
    }
  }, []);

  const rows = useMemo(() => {
    if (!prod) return [];
    return prod.items.map((it) => {
      const match = avail ? avail.map.get(norm(it.halb)) : null;
      const a = {};
      AVAIL_COLS.forEach((c) => (a[c.key] = match ? match[c.key] : ""));
      // Linha vermelha apenas após importar disponibilidade
      let flagged = false;
      if (avail) {
        flagged = AVAIL_COLS.some((c) => isBad(a[c.key]));
      }
      return { ...it, ...a, _flagged: flagged, _matched: !!match };
    });
  }, [prod, avail]);

  const flaggedCount = rows.filter((r) => r._flagged).length;
  const allCols = [...PROD_COLS, ...AVAIL_COLS];
  const hasData = rows.length > 0;

  const reset = () => {
    setProd(null);
    setAvail(null);
    setError("");
    setView("import");
  };

  return (
    <div style={{ minHeight: "100vh", background: "#F7F8FA", color: "#0F172A", fontFamily: "Manrope, system-ui, sans-serif" }}>
      <style>{CSS}</style>

      {/* Cabeçalho */}
      <header className="no-print" style={{ borderBottom: "1px solid #E5E9F0", background: "#FFFFFF" }}>
        <div style={{ maxWidth: 1400, margin: "0 auto", padding: "20px 28px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 16, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            <div style={{ width: 6, height: 46, borderRadius: 4, background: "linear-gradient(180deg,#2563EB,#1E3A8A)" }} />
            <div>
              <div style={{ fontFamily: "Sora, sans-serif", fontWeight: 700, fontSize: 20, letterSpacing: ".02em", color: "#1E3A8A", lineHeight: 1.1 }}>
                CONTROLE DA QUALIDADE
              </div>
              <div style={{ fontSize: 13.5, color: "#64748B", fontWeight: 500 }}>
                Análise de risco · Reunião de antecipação
              </div>
            </div>
          </div>

          <div style={{ display: "flex", gap: 10 }}>
            {view === "summary" && (
              <button onClick={() => setView("import")} className="btn-ghost">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: 6 }}>
                  <path d="M19 12H5M12 19l-7-7 7-7" />
                </svg>
                Importações
              </button>
            )}
            {hasData && (
              <button onClick={reset} className="btn-ghost">
                Limpar
              </button>
            )}
            {view === "summary" && (
              <button onClick={() => window.print()} disabled={!hasData} className="btn-primary" style={{ opacity: hasData ? 1 : 0.45, cursor: hasData ? "pointer" : "not-allowed" }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" />
                  <path d="M6 14h12v8H6z" />
                </svg>
                Gerar impressão / PDF
              </button>
            )}
          </div>
        </div>
      </header>

      <main style={{ maxWidth: 1400, margin: "0 auto", padding: "24px 28px 48px" }}>
        {/* ===================== TELA DE IMPORTAÇÃO ===================== */}
        {view === "import" && (
        <section className="no-print" style={cardStyle}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
            <h2 style={{ margin: 0, fontSize: 15, fontWeight: 700, color: "#0F172A" }}>Importação de planilhas</h2>
            <span style={{ fontSize: 12, color: "#94A3B8" }}>Formatos: .xlsx · .xls · .csv</span>
          </div>
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
            <UploadZone
              title="Sequência de Produção"
              subtitle="Pedido, Ordem, Halb, Enfornamento…"
              accentColor="#2563EB"
              file={prod?.fileName}
              status={prod ? `${prod.items.length} item(ns)` : ""}
              onFile={onProd}
            />
            <UploadZone
              title="Disponibilidade de Instrumentos"
              subtitle="Correlação por HALBQ"
              accentColor="#475569"
              file={avail?.fileName}
              status={avail ? `${avail.count} registro(s)` : ""}
              onFile={onAvail}
            />
          </div>

          {(prod?.missing?.length || avail?.missing?.length) ? (
            <div style={{ marginTop: 14, fontSize: 12.5, color: "#B45309", background: "#FFFBEB", border: "1px solid #FDE68A", borderRadius: 10, padding: "8px 12px" }}>
              ⚠ Colunas não localizadas automaticamente:{" "}
              {[...(prod?.missing || []), ...(avail?.missing || [])].join(", ")}. Verifique os títulos na planilha.
            </div>
          ) : null}

          {error && (
            <div style={{ marginTop: 14, fontSize: 12.5, color: "#B91C1C", background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: 10, padding: "8px 12px" }}>
              {error}
            </div>
          )}

          <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 20, borderTop: "1px solid #EEF1F6", paddingTop: 18 }}>
            <button
              onClick={() => setView("summary")}
              disabled={!hasData}
              className="btn-primary"
              style={{ opacity: hasData ? 1 : 0.45, cursor: hasData ? "pointer" : "not-allowed" }}
            >
              Ver resumo
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M5 12h14M12 5l7 7-7 7" />
              </svg>
            </button>
          </div>
          {!hasData && (
            <div style={{ fontSize: 12, color: "#94A3B8", textAlign: "right", marginTop: 6 }}>
              Importe ao menos a planilha de Sequência de Produção para habilitar o resumo.
            </div>
          )}
        </section>
        )}

        {/* ===================== TELA DE RESUMO ===================== */}
        {view === "summary" && (
        <>
        {/* Resumo / legenda */}
        {hasData && (
          <div className="no-print" style={{ display: "flex", alignItems: "center", gap: 18, margin: "0 2px 12px", flexWrap: "wrap" }}>
            <Stat label="Itens" value={rows.length} color="#2563EB" />
            <Stat label="Correlacionados" value={rows.filter((r) => r._matched).length} color="#0F766E" />
            <Stat label="Em risco" value={flaggedCount} color="#DC2626" />
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginLeft: "auto", fontSize: 12.5, color: "#64748B" }}>
              <span style={{ width: 14, height: 14, borderRadius: 4, background: "#FDE3E3", border: "1px solid #F4B4B4" }} />
              Indisponível ou em branco
            </div>
          </div>
        )}

        {/* Tabela resumo */}
        <section className="print-area">
          {hasData ? (
            <div className="table-scroll" style={{ border: "1px solid #E5E9F0", borderRadius: 14, overflow: "auto", maxHeight: "74vh", background: "#fff", boxShadow: "0 1px 3px rgba(15,23,42,.05)" }}>
              <table style={{ borderCollapse: "separate", borderSpacing: 0, width: "100%", fontSize: 10 }}>
                <thead>
                  {/* Cabeçalho de impressão — repete em todas as páginas do PDF */}
                  <tr className="print-row">
                    <th className="print-title-cell" colSpan={allCols.length}>
                      <div className="pt-title">
                        <strong>CONTROLE DA QUALIDADE</strong>
                        <span>Análise de risco · Reunião de antecipação</span>
                      </div>
                    </th>
                  </tr>
                  <tr className="print-row">
                    <th className="print-sum-cell" colSpan={allCols.length}>
                      <span><b>{rows.length}</b> Itens</span>
                      <span><b>{rows.filter((r) => r._matched).length}</b> Correlacionados</span>
                      <span className="ps-risk"><b>{flaggedCount}</b> Em risco</span>
                      <span className="ps-note">Linhas em vermelho · indisponível ou em branco</span>
                    </th>
                  </tr>
                  <tr className="group-row">
                    <th className="group-cell grp-prod" colSpan={PROD_COLS.length}>
                      Sequência de Produção
                    </th>
                    <th className="group-cell grp-inst" colSpan={AVAIL_COLS.length}>
                      Disponibilidade de Instrumentos
                    </th>
                  </tr>
                  <tr className="head-row">
                    {allCols.map((c, i) => (
                      <th
                        key={c.key}
                        className="head-cell"
                        style={{
                          borderLeft: i === PROD_COLS.length ? "2px solid #CBD5E1" : undefined,
                        }}
                      >
                        {c.label}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {rows.map((r, ri) => (
                    <tr key={ri} className={r._flagged ? "row-bad" : ri % 2 ? "row-alt" : ""}>
                      {allCols.map((c, i) => {
                        const v = r[c.key];
                        const instr = i >= PROD_COLS.length;
                        const cellBad = r._flagged && instr && isBad(v);
                        return (
                          <td
                            key={c.key}
                            className="cell"
                            style={{
                              borderLeft: i === PROD_COLS.length ? "2px solid #CBD5E1" : undefined,
                              fontWeight: cellBad ? 700 : 400,
                              color: cellBad ? "#B91C1C" : "#1E293B",
                              whiteSpace: "nowrap",
                            }}
                          >
                            {trim(v) || (instr && r._flagged ? "Indisponível" : "—")}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div className="no-print" style={{ textAlign: "center", color: "#94A3B8", padding: "60px 20px", border: "1.5px dashed #D8DEE9", borderRadius: 14, background: "#FCFDFE" }}>
              <div style={{ fontSize: 15, fontWeight: 600, color: "#64748B" }}>Nenhum dado para exibir</div>
              <div style={{ fontSize: 13, marginTop: 4 }}>Importe ao menos a planilha de Sequência de Produção para gerar a tabela resumo.</div>
            </div>
          )}
        </section>
        </>
        )}
      </main>
    </div>
  );
}

function Stat({ label, value, color }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <span style={{ fontFamily: "Sora, sans-serif", fontWeight: 700, fontSize: 22, color }}>{value}</span>
      <span style={{ fontSize: 12.5, color: "#64748B", fontWeight: 500 }}>{label}</span>
    </div>
  );
}

const cardStyle = {
  background: "#FFFFFF",
  border: "1px solid #E5E9F0",
  borderRadius: 16,
  padding: "22px 24px",
  boxShadow: "0 1px 3px rgba(15,23,42,.05)",
};

/* ------------------------------------------------------------------ *
 *  CSS (injetado) — inclui regras de impressão A4 paisagem
 * ------------------------------------------------------------------ */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700&family=Manrope:wght@400;500;600;700&display=swap');

* { box-sizing: border-box; }
body { margin: 0; }

.btn-primary {
  display: inline-flex; align-items: center; gap: 8px;
  background: #2563EB; color: #fff; border: none;
  font-family: 'Manrope', sans-serif; font-weight: 600; font-size: 13.5px;
  padding: 10px 16px; border-radius: 10px;
  box-shadow: 0 1px 2px rgba(37,99,235,.35);
}
.btn-primary:hover:not(:disabled) { background: #1D4FD7; }
.btn-ghost {
  display: inline-flex; align-items: center;
  background: #fff; color: #475569; border: 1px solid #D8DEE9;
  font-family: 'Manrope', sans-serif; font-weight: 600; font-size: 13.5px;
  padding: 10px 16px; border-radius: 10px; cursor: pointer;
}
.btn-ghost:hover { background: #F1F5F9; }

.table-scroll::-webkit-scrollbar { width: 11px; height: 11px; }
.table-scroll::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 8px; border: 3px solid #fff; }
.table-scroll::-webkit-scrollbar-track { background: #F1F5F9; }

.group-cell {
  position: sticky; top: 0; z-index: 6;
  height: 30px; padding: 0 11px;
  font-family: 'Sora', sans-serif; font-weight: 700; font-size: 10.5px;
  letter-spacing: .05em; text-transform: uppercase; text-align: left;
  color: #fff;
}
.grp-prod { background: #1E3A8A; }
.grp-inst { background: #475569; border-left: 2px solid #fff; }

.head-cell {
  position: sticky; top: 30px; z-index: 5;
  background: #EEF2F8; color: #3B4658;
  font-weight: 700; font-size: 9.5px; letter-spacing: .02em;
  text-transform: uppercase; text-align: left;
  padding: 7px 11px; white-space: nowrap;
  border-bottom: 1px solid #DCE3EC;
}

.cell {
  padding: 5px 11px; border-bottom: 1px solid #EEF1F6;
  vertical-align: middle; font-size: 9.5px;
}
tbody tr.row-alt { background: #FAFBFD; }
tbody tr:hover { background: #EFF4FF; }
tbody tr.row-bad { background: #FDE6E6 !important; }
tbody tr.row-bad:hover { background: #FBD5D5 !important; }

/* Linhas exclusivas do PDF (ficam no <thead>, repetem em todas as páginas) */
.print-row { display: none; }
.print-only { display: none; }

@media print {
  @page { size: A4 landscape; margin: 7mm; }
  html, body { background: #fff !important; }
  .no-print { display: none !important; }
  .print-only { display: block; }
  .print-row { display: table-row; }

  .print-area { margin: 0 !important; }
  .table-scroll {
    max-height: none !important; overflow: visible !important;
    border: none !important; border-radius: 0 !important; box-shadow: none !important;
  }
  table { font-size: 7.5px !important; width: 100% !important; }
  thead { display: table-header-group; }
  tr { page-break-inside: avoid; }

  /* Título — repetido no topo de todas as páginas */
  .print-title-cell {
    padding: 0 0 4px; border-bottom: 2px solid #1E3A8A; text-align: left;
    background: #fff !important;
  }
  .pt-title strong { font-family: 'Sora', sans-serif; display: block; font-size: 13px; letter-spacing: .03em; color: #1E3A8A; }
  .pt-title span { font-size: 9px; color: #475569; }

  /* Resumo de quantidades — topo de cada página */
  .print-sum-cell {
    padding: 5px 0 7px; text-align: left; background: #fff !important;
    font-weight: 500; color: #334155; font-size: 8.5px;
  }
  .print-sum-cell span { margin-right: 16px; }
  .print-sum-cell b { font-family: 'Sora', sans-serif; font-size: 10px; color: #1E3A8A; margin-right: 3px; }
  .print-sum-cell .ps-risk b { color: #C0392B; }
  .print-sum-cell .ps-note { color: #94A3B8; font-style: italic; }

  .group-cell { position: static; height: auto; padding: 3px 5px; font-size: 7.5px;
    -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .head-cell { position: static; padding: 3px 5px; font-size: 7.5px; letter-spacing: 0;
    background: #EEF2F8 !important; color: #334155 !important;
    -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .cell { padding: 2.5px 5px; font-size: 7.5px; }
  tbody tr.row-bad, .grp-prod, .grp-inst, tbody tr.row-alt {
    -webkit-print-color-adjust: exact; print-color-adjust: exact;
  }
}
`;
