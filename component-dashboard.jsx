import React, { useState, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, Legend,
} from "recharts";
import { Upload, Cpu, RefreshCw, AlertTriangle, LayoutGrid, ChevronUp, ChevronDown } from "lucide-react";

const PALETTE = ["#E8874A", "#5FC9C0", "#D97BA8", "#9AD16A", "#B79CF0", "#F0C15C", "#6FA8DC", "#E86A6A"];

const COLORS = {
  bg: "#0F2036",
  panel: "#16304E",
  panelAlt: "#1C3B5E",
  border: "#2C4C70",
  ink: "#EAF1F8",
  inkMuted: "#8AA0BA",
  accent: "#E8874A",
  accent2: "#5FC9C0",
  danger: "#E86A6A",
};

const FIELD_MATCHERS = {
  id: [/^id$/i, /^sku$/i, /item\s*id/i],
  category: [/categ/i, /type/i, /group/i],
  name: [/component/i, /product/i, /item\s*name/i, /^name$/i],
  brand: [/brand/i, /manufacturer/i, /vendor/i],
  model: [/model/i],
  spec: [/spec/i, /detail/i, /description/i],
  price: [/price/i, /cost/i, /unit\s*cost/i],
  stock: [/stock/i, /qty/i, /quantity/i, /units?\s*on\s*hand/i],
  value: [/total\s*value/i, /^value$/i, /inventory\s*value/i],
  warranty: [/warrant/i],
};

function mapRow(rawRow) {
  const keys = Object.keys(rawRow);
  const used = new Set();
  const picked = {};
  for (const field of Object.keys(FIELD_MATCHERS)) {
    let found = null;
    for (const pattern of FIELD_MATCHERS[field]) {
      found = keys.find((k) => !used.has(k) && pattern.test(k));
      if (found) break;
    }
    if (found) {
      used.add(found);
      picked[field] = rawRow[found];
    }
  }
  return picked;
}

function toNumber(v) {
  if (v === null || v === undefined || v === "") return 0;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function formatCurrency(n) {
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(n);
}

function formatNumber(n) {
  return new Intl.NumberFormat("en-US").format(Math.round(n));
}

function CornerCard({ children, style, className = "" }) {
  return (
    <div
      className={className}
      style={{
        position: "relative",
        background: COLORS.panel,
        border: `1px solid ${COLORS.border}`,
        padding: "18px 20px",
        ...style,
      }}
    >
      <span style={{ position: "absolute", top: -1, left: -1, width: 10, height: 10, borderTop: `2px solid ${COLORS.accent}`, borderLeft: `2px solid ${COLORS.accent}` }} />
      <span style={{ position: "absolute", bottom: -1, right: -1, width: 10, height: 10, borderBottom: `2px solid ${COLORS.accent}`, borderRight: `2px solid ${COLORS.accent}` }} />
      {children}
    </div>
  );
}

function StatCard({ label, value, sub }) {
  return (
    <CornerCard style={{ flex: "1 1 180px", minWidth: 160 }}>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.inkMuted, letterSpacing: 0.2 }}>{label}</div>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 28, fontWeight: 600, color: COLORS.ink, marginTop: 6 }}>{value}</div>
      {sub ? <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: COLORS.accent2, marginTop: 4 }}>{sub}</div> : null}
    </CornerCard>
  );
}

function CustomTooltip({ active, payload, label, money }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div style={{ background: COLORS.panelAlt, border: `1px solid ${COLORS.border}`, padding: "8px 12px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.ink }}>
      <div style={{ color: COLORS.inkMuted, marginBottom: 4 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color || COLORS.ink }}>
          {p.name}: {money ? formatCurrency(p.value) : formatNumber(p.value)}
        </div>
      ))}
    </div>
  );
}

export default function ComponentDashboard() {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [activeCategory, setActiveCategory] = useState("all");
  const [sortKey, setSortKey] = useState("value");
  const [sortDir, setSortDir] = useState("desc");
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  const handleFile = useCallback((file) => {
    if (!file) return;
    setLoading(true);
    setError("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const sheetName = wb.SheetNames[0];
        const sheet = wb.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: null });
        if (!json.length) {
          setError("That sheet doesn't have any rows to read.");
          setRows(null);
          setLoading(false);
          return;
        }
        const mapped = json.map((r, i) => {
          const m = mapRow(r);
          const price = toNumber(m.price);
          const stock = toNumber(m.stock);
          const value = m.value !== undefined ? toNumber(m.value) : price * stock;
          return {
            id: m.id ?? `row-${i}`,
            category: (m.category ?? "Uncategorized").toString().trim() || "Uncategorized",
            name: m.name ?? "Unnamed component",
            brand: m.brand ?? "—",
            model: m.model ?? "—",
            spec: m.spec ?? "",
            price,
            stock,
            value,
            warranty: m.warranty ?? null,
          };
        });
        if (!mapped.some((r) => r.price || r.stock)) {
          setError("Couldn't find price or stock columns in this file — check the headers and try again.");
          setRows(null);
          setLoading(false);
          return;
        }
        setRows(mapped);
        setFileName(file.name);
        setActiveCategory("all");
      } catch (err) {
        setError("Couldn't read that file. Make sure it's a valid .xlsx or .csv file.");
        setRows(null);
      }
      setLoading(false);
    };
    reader.onerror = () => {
      setError("Something went wrong reading the file.");
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onInputChange = (e) => {
    const file = e.target.files && e.target.files[0];
    handleFile(file);
  };

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    handleFile(file);
  };

  const categories = useMemo(() => {
    if (!rows) return [];
    return Array.from(new Set(rows.map((r) => r.category))).sort();
  }, [rows]);

  const filteredRows = useMemo(() => {
    if (!rows) return [];
    return activeCategory === "all" ? rows : rows.filter((r) => r.category === activeCategory);
  }, [rows, activeCategory]);

  const stats = useMemo(() => {
    if (!rows) return null;
    const totalValue = rows.reduce((s, r) => s + r.value, 0);
    const totalStock = rows.reduce((s, r) => s + r.stock, 0);
    const avgPrice = rows.length ? rows.reduce((s, r) => s + r.price, 0) / rows.length : 0;
    return {
      count: rows.length,
      totalValue,
      totalStock,
      categories: categories.length,
      avgPrice,
    };
  }, [rows, categories]);

  const categoryAgg = useMemo(() => {
    if (!rows) return [];
    const map = new Map();
    for (const r of rows) {
      const cur = map.get(r.category) || { category: r.category, value: 0, stock: 0, count: 0 };
      cur.value += r.value;
      cur.stock += r.stock;
      cur.count += 1;
      map.set(r.category, cur);
    }
    return Array.from(map.values()).sort((a, b) => b.value - a.value);
  }, [rows]);

  const sortedTable = useMemo(() => {
    const copy = [...filteredRows];
    copy.sort((a, b) => {
      const av = a[sortKey];
      const bv = b[sortKey];
      if (typeof av === "string" || typeof bv === "string") {
        return sortDir === "asc"
          ? String(av).localeCompare(String(bv))
          : String(bv).localeCompare(String(av));
      }
      return sortDir === "asc" ? av - bv : bv - av;
    });
    return copy;
  }, [filteredRows, sortKey, sortDir]);

  const toggleSort = (key) => {
    if (sortKey === key) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortKey(key);
      setSortDir("desc");
    }
  };

  const reset = () => {
    setRows(null);
    setFileName("");
    setError("");
    setActiveCategory("all");
  };

  return (
    <div
      style={{
        fontFamily: "'Space Grotesk', sans-serif",
        background: COLORS.bg,
        color: COLORS.ink,
        minHeight: 560,
        padding: 24,
        backgroundImage:
          "linear-gradient(rgba(95,201,192,0.06) 1px, transparent 1px), linear-gradient(90deg, rgba(95,201,192,0.06) 1px, transparent 1px)",
        backgroundSize: "28px 28px",
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');
        .cd-th { cursor: pointer; user-select: none; }
        .cd-th:hover { color: ${COLORS.accent2}; }
        .cd-chip { transition: background 0.15s ease, color 0.15s ease; }
        .cd-row:hover { background: ${COLORS.panelAlt}; }
        .cd-upload-btn:hover { background: ${COLORS.accent}; color: ${COLORS.bg}; }
      `}</style>

      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 16, marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 34, height: 34, border: `1px solid ${COLORS.accent}`, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <Cpu size={18} color={COLORS.accent} />
          </div>
          <div>
            <div style={{ fontSize: 18, fontWeight: 600 }}>Component inventory dashboard</div>
            <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 11, color: COLORS.inkMuted }}>
              {fileName ? fileName : "No file loaded"}
            </div>
          </div>
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" onChange={onInputChange} style={{ display: "none" }} />
          <button
            className="cd-upload-btn"
            onClick={() => inputRef.current && inputRef.current.click()}
            style={{
              display: "flex", alignItems: "center", gap: 8,
              background: "transparent", border: `1px solid ${COLORS.accent}`, color: COLORS.accent,
              padding: "9px 16px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12,
              cursor: "pointer",
            }}
          >
            <Upload size={14} />
            {rows ? "Replace file" : "Import file"}
          </button>
          {rows ? (
            <button
              onClick={reset}
              style={{
                display: "flex", alignItems: "center", gap: 8,
                background: "transparent", border: `1px solid ${COLORS.border}`, color: COLORS.inkMuted,
                padding: "9px 16px", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, cursor: "pointer",
              }}
            >
              <RefreshCw size={14} />
              Clear
            </button>
          ) : null}
        </div>
      </div>

      {error ? (
        <div style={{ display: "flex", gap: 8, alignItems: "center", background: "rgba(232,106,106,0.1)", border: `1px solid ${COLORS.danger}`, color: COLORS.danger, padding: "10px 14px", marginBottom: 20, fontFamily: "'IBM Plex Mono', monospace", fontSize: 12 }}>
          <AlertTriangle size={16} />
          {error}
        </div>
      ) : null}

      {!rows ? (
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={() => inputRef.current && inputRef.current.click()}
          style={{
            border: `1.5px dashed ${dragOver ? COLORS.accent2 : COLORS.border}`,
            background: dragOver ? "rgba(95,201,192,0.05)" : "transparent",
            minHeight: 320,
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            justifyContent: "center",
            gap: 12,
            cursor: "pointer",
            textAlign: "center",
            padding: 24,
          }}
        >
          <LayoutGrid size={30} color={dragOver ? COLORS.accent2 : COLORS.inkMuted} />
          <div style={{ fontSize: 16, fontWeight: 500 }}>{loading ? "Reading file…" : "Drop your spreadsheet here"}</div>
          <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.inkMuted, maxWidth: 420 }}>
            .xlsx, .xls or .csv — expects columns like category, component name, brand, price and stock quantity. Nothing leaves your browser.
          </div>
        </div>
      ) : (
        <>
          {/* Stat cards */}
          <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 20 }}>
            <StatCard label="Components" value={formatNumber(stats.count)} />
            <StatCard label="Inventory value" value={formatCurrency(stats.totalValue)} sub={activeCategory !== "all" ? `filtered: ${activeCategory}` : null} />
            <StatCard label="Units in stock" value={formatNumber(stats.totalStock)} />
            <StatCard label="Categories" value={formatNumber(stats.categories)} />
            <StatCard label="Avg. price" value={formatCurrency(stats.avgPrice)} />
          </div>

          {/* Category filter chips */}
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 20 }}>
            <button
              className="cd-chip"
              onClick={() => setActiveCategory("all")}
              style={{
                fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, padding: "6px 12px",
                border: `1px solid ${activeCategory === "all" ? COLORS.accent : COLORS.border}`,
                background: activeCategory === "all" ? COLORS.accent : "transparent",
                color: activeCategory === "all" ? COLORS.bg : COLORS.inkMuted,
                cursor: "pointer",
              }}
            >
              all ({rows.length})
            </button>
            {categories.map((c) => (
              <button
                key={c}
                className="cd-chip"
                onClick={() => setActiveCategory(c)}
                style={{
                  fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, padding: "6px 12px",
                  border: `1px solid ${activeCategory === c ? COLORS.accent : COLORS.border}`,
                  background: activeCategory === c ? COLORS.accent : "transparent",
                  color: activeCategory === c ? COLORS.bg : COLORS.inkMuted,
                  cursor: "pointer",
                }}
              >
                {c} ({rows.filter((r) => r.category === c).length})
              </button>
            ))}
          </div>

          {/* Charts */}
          <div style={{ display: "grid", gridTemplateColumns: "1.3fr 1fr", gap: 16, marginBottom: 20 }}>
            <CornerCard>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.inkMuted, marginBottom: 8 }}>
                Inventory value by category
              </div>
              <div style={{ height: 260 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={categoryAgg} margin={{ top: 4, right: 8, left: 0, bottom: 4 }}>
                    <CartesianGrid stroke={COLORS.border} strokeDasharray="3 3" vertical={false} />
                    <XAxis dataKey="category" tick={{ fill: COLORS.inkMuted, fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                    <YAxis tick={{ fill: COLORS.inkMuted, fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${Math.round(v / 1000)}k`} />
                    <Tooltip content={<CustomTooltip money />} cursor={{ fill: "rgba(255,255,255,0.04)" }} />
                    <Bar dataKey="value" name="Value" radius={[2, 2, 0, 0]}>
                      {categoryAgg.map((entry, i) => (
                        <Cell key={entry.category} fill={PALETTE[i % PALETTE.length]} />
                      ))}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </CornerCard>

            <CornerCard>
              <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 12, color: COLORS.inkMuted, marginBottom: 8 }}>
                Stock distribution
              </div>
              <div style={{ height: 260 }}>
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie
                      data={categoryAgg}
                      dataKey="stock"
                      nameKey="category"
                      innerRadius={50}
                      outerRadius={85}
                      paddingAngle={2}
                    >
                      {categoryAgg.map((entry, i) => (
                        <Cell key={entry.category} fill={PALETTE[i % PALETTE.length]} stroke={COLORS.panel} strokeWidth={2} />
                      ))}
                    </Pie>
                    <Tooltip content={<CustomTooltip />} />
                    <Legend
                      layout="vertical"
                      align="right"
                      verticalAlign="middle"
                      formatter={(value) => <span style={{ color: COLORS.inkMuted, fontFamily: "'IBM Plex Mono', monospace", fontSize: 11 }}>{value}</span>}
                    />
                  </PieChart>
                </ResponsiveContainer>
              </div>
            </CornerCard>
          </div>

          {/* Table */}
          <CornerCard style={{ padding: 0, overflow: "hidden" }}>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontFamily: "'IBM Plex Mono', monospace", fontSize: 12.5 }}>
                <thead>
                  <tr style={{ background: COLORS.panelAlt, borderBottom: `1px solid ${COLORS.border}` }}>
                    {[
                      ["category", "Category"],
                      ["name", "Component"],
                      ["brand", "Brand"],
                      ["price", "Price"],
                      ["stock", "Stock"],
                      ["value", "Value"],
                    ].map(([key, label]) => (
                      <th
                        key={key}
                        className="cd-th"
                        onClick={() => toggleSort(key)}
                        style={{ textAlign: key === "price" || key === "stock" || key === "value" ? "right" : "left", padding: "10px 14px", color: COLORS.inkMuted, fontWeight: 500 }}
                      >
                        <span style={{ display: "inline-flex", alignItems: "center", gap: 4 }}>
                          {label}
                          {sortKey === key ? (sortDir === "asc" ? <ChevronUp size={12} /> : <ChevronDown size={12} />) : null}
                        </span>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sortedTable.map((r, i) => (
                    <tr key={r.id + i} className="cd-row" style={{ borderBottom: `1px solid ${COLORS.border}` }}>
                      <td style={{ padding: "9px 14px", color: COLORS.accent2 }}>{r.category}</td>
                      <td style={{ padding: "9px 14px", color: COLORS.ink }}>{r.name}</td>
                      <td style={{ padding: "9px 14px", color: COLORS.inkMuted }}>{r.brand}</td>
                      <td style={{ padding: "9px 14px", color: COLORS.ink, textAlign: "right" }}>{formatCurrency(r.price)}</td>
                      <td style={{ padding: "9px 14px", color: COLORS.ink, textAlign: "right" }}>{formatNumber(r.stock)}</td>
                      <td style={{ padding: "9px 14px", color: COLORS.accent, textAlign: "right", fontWeight: 500 }}>{formatCurrency(r.value)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </CornerCard>
        </>
      )}
    </div>
  );
}
