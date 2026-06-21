import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── FONT ───────────────────────────────────────────────────────────────────
const fontLink = document.createElement("link");
fontLink.rel = "stylesheet";
fontLink.href = "https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap";
document.head.appendChild(fontLink);

// ─── THEME ──────────────────────────────────────────────────────────────────
const T = {
  sage: "#5A7A5C", sageDark: "#3E5840", sageLight: "#8FAF91",
  cream: "#F5F0E8", creamDark: "#EDE5D8",
  brown: "#C4A882", brownDark: "#9E7D56",
  gray: "#8B8680", grayLight: "#D8D3CC",
  white: "#FAFAF7", text: "#2C2A28", textSoft: "#6B6560",
  red: "#C0392B", amber: "#D4840A",
};

// ─── STORAGE ────────────────────────────────────────────────────────────────
const STORE = {
  get: (k) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : null; } catch { return null; } },
  set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
};

const initMaster = () => STORE.get("master") || { customers: [], models: [], items: [], processes: [] };
const initProduction = () => STORE.get("production") || [];
const getToday = () => new Date().toISOString().split("T")[0];
const getCurrentShift = () => { const h = new Date().getHours(); return h >= 6 && h < 14 ? "1" : h >= 14 && h < 22 ? "2" : "3"; };

// ─── ICONS ──────────────────────────────────────────────────────────────────
const icons = {
  plus: "M12 5v14M5 12h14",
  trash: "M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6",
  x: "M18 6L6 18M6 6l12 12",
  chevronDown: "M6 9l6 6 6-6",
  settings: "M12 15a3 3 0 100-6 3 3 0 000 6z",
  clipboard: "M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2",
  chart: "M18 20V10M12 20V4M6 20v-6",
  check: "M20 6L9 17l-5-5",
  download: "M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4M7 10l5 5 5-5M12 15V3",
  lock: "M19 11H5a2 2 0 00-2 2v7a2 2 0 002 2h14a2 2 0 002-2v-7a2 2 0 00-2-2zM7 11V7a5 5 0 0110 0v4",
  unlock: "M19 11H5a2 2 0 00-2 2v7a2 2 0 002 2h14a2 2 0 002-2v-7a2 2 0 00-2-2zM7 11V7a5 5 0 019.9-1",
  sun: "M12 7a5 5 0 100 10A5 5 0 0012 7zM12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42",
  moon: "M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z",
  factory: "M2 20h20M4 20V10l6-6v6l6-6v16",
};

const Icon = ({ name, size = 20, color = "currentColor" }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none"
    stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    {(icons[name] || "").split("M").filter(Boolean).map((d, i) => <path key={i} d={"M" + d} />)}
  </svg>
);

// ─── LOGO ───────────────────────────────────────────────────────────────────
const Logo = ({ size = 38 }) => (
  <svg width={size} height={size} viewBox="0 0 48 48" fill="none">
    <rect width="48" height="48" rx="12" fill={T.white + "25"} />
    {/* Factory silhouette */}
    <rect x="6" y="28" width="36" height="14" rx="2" fill={T.white + "60"} />
    <rect x="10" y="20" width="8" height="8" fill={T.sageLight + "90"} />
    <rect x="22" y="14" width="8" height="14" fill={T.sageLight + "90"} />
    <rect x="34" y="20" width="8" height="8" fill={T.sageLight + "90"} />
    {/* Chimney */}
    <rect x="12" y="10" width="3" height="10" fill={T.white + "80"} />
    <rect x="25" y="6" width="3" height="8" fill={T.white + "80"} />
    {/* Windows */}
    <rect x="8" y="32" width="4" height="5" rx="1" fill={T.amber + "90"} />
    <rect x="15" y="32" width="4" height="5" rx="1" fill={T.amber + "90"} />
    <rect x="22" y="32" width="4" height="5" rx="1" fill={T.amber + "90"} />
    <rect x="29" y="32" width="4" height="5" rx="1" fill={T.amber + "90"} />
    <rect x="36" y="32" width="4" height="5" rx="1" fill={T.amber + "90"} />
    {/* Gear accent */}
    <circle cx="38" cy="12" r="5" fill="none" stroke={T.white + "70"} strokeWidth="2" />
    <circle cx="38" cy="12" r="2" fill={T.white + "70"} />
  </svg>
);

// ─── UI COMPONENTS ──────────────────────────────────────────────────────────
const Btn = ({ children, onClick, variant = "primary", small, style, disabled }) => {
  const base = {
    display: "inline-flex", alignItems: "center", justifyContent: "center", gap: 5,
    fontFamily: "Montserrat,sans-serif", fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer",
    border: "none", borderRadius: 9, transition: "all .18s", opacity: disabled ? .45 : 1,
    padding: small ? "6px 12px" : "10px 18px", fontSize: small ? 11 : 13,
  };
  const vars = {
    primary: { background: T.sage, color: T.white },
    danger:  { background: T.red,  color: T.white },
    outline: { background: "transparent", color: T.sage, border: `2px solid ${T.sage}` },
    ghost:   { background: T.creamDark, color: T.text },
    amber:   { background: T.amber, color: T.white },
    green:   { background: "#2E7D32", color: T.white },
  };
  return <button onClick={disabled ? undefined : onClick} style={{ ...base, ...vars[variant], ...style }}>{children}</button>;
};

const Field = ({ label, children }) => (
  <div style={{ marginBottom: 12 }}>
    {label && <div style={{ fontSize: 11, fontWeight: 600, color: T.textSoft, marginBottom: 4, textTransform: "uppercase", letterSpacing: .5 }}>{label}</div>}
    {children}
  </div>
);

const inputStyle = {
  width: "100%", padding: "10px 13px", borderRadius: 9, border: `1.5px solid ${T.grayLight}`,
  fontFamily: "Montserrat,sans-serif", fontSize: 14, color: T.text, background: T.white,
  outline: "none", boxSizing: "border-box",
};

const Input = ({ label, value, onChange, placeholder, type = "text" }) => (
  <Field label={label}>
    <input type={type} value={value} onChange={e => onChange(e.target.value)}
      placeholder={placeholder} style={inputStyle} />
  </Field>
);

const Select = ({ label, value, onChange, options, placeholder }) => (
  <Field label={label}>
    <div style={{ position: "relative" }}>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ ...inputStyle, paddingRight: 34, appearance: "none", color: value ? T.text : T.gray }}>
        <option value="">{placeholder || "Pilih..."}</option>
        {options.map(o => <option key={o.value ?? o} value={o.value ?? o}>{o.label ?? o}</option>)}
      </select>
      <div style={{ position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)", pointerEvents: "none", color: T.gray }}>
        <Icon name="chevronDown" size={15} />
      </div>
    </div>
  </Field>
);

const Card = ({ children, style }) => (
  <div style={{ background: T.white, borderRadius: 13, padding: 14, boxShadow: "0 2px 10px rgba(0,0,0,.07)", marginBottom: 12, ...style }}>{children}</div>
);

const Tag = ({ label, onRemove, color }) => (
  <span style={{ display: "inline-flex", alignItems: "center", gap: 4, padding: "3px 8px", background: color || T.creamDark, borderRadius: 20, fontSize: 11, fontWeight: 600, color: T.text, margin: "2px" }}>
    {label}
    {onRemove && <span onClick={onRemove} style={{ cursor: "pointer", lineHeight: 1 }}><Icon name="x" size={11} color={T.red} /></span>}
  </span>
);

// ─── MULTI-PROCESS INPUT ─────────────────────────────────────────────────────
const MultiProcessInput = ({ label, value = [], onChange, availableProcesses = [] }) => {
  const [inp, setInp] = useState("");
  const add = (v) => { const val = (v || inp).trim(); if (val && !value.includes(val)) { onChange([...value, val]); setInp(""); } };
  const remove = (p) => onChange(value.filter(x => x !== p));
  return (
    <Field label={label}>
      <div style={{ border: `1.5px solid ${T.grayLight}`, borderRadius: 9, padding: "7px 9px", background: T.white }}>
        <div style={{ display: "flex", flexWrap: "wrap", marginBottom: value.length ? 6 : 0 }}>
          {value.map(p => <Tag key={p} label={p} onRemove={() => remove(p)} color={T.sageLight + "40"} />)}
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          <input value={inp} onChange={e => setInp(e.target.value)}
            onKeyDown={e => e.key === "Enter" && add()}
            placeholder="Ketik + Enter…"
            style={{ flex: 1, border: "none", outline: "none", fontFamily: "Montserrat", fontSize: 12, background: "transparent", color: T.text }} />
          <Btn small onClick={() => add()}>+ Tambah</Btn>
        </div>
        {availableProcesses.filter(p => !value.includes(p)).length > 0 && (
          <div style={{ marginTop: 6, borderTop: `1px solid ${T.creamDark}`, paddingTop: 6, display: "flex", flexWrap: "wrap" }}>
            {availableProcesses.filter(p => !value.includes(p)).map(p => (
              <span key={p} onClick={() => add(p)} style={{ cursor: "pointer", padding: "2px 8px", background: T.brown + "30", borderRadius: 20, fontSize: 11, margin: "2px", fontWeight: 500, color: T.brownDark }}>+ {p}</span>
            ))}
          </div>
        )}
      </div>
    </Field>
  );
};

// ═══════════════════════════════════════════════════════════════════════════
// PAGE: MASTER DATA  (no confirmation)
// ═══════════════════════════════════════════════════════════════════════════
const MasterData = ({ master, setMaster }) => {
  const [tab, setTab] = useState("customer");
  const [form, setForm] = useState({ name: "", processes: [] });

  const tabs = [{ id: "customer", label: "Customer" }, { id: "model", label: "Model" }, { id: "item", label: "Item" }, { id: "proses", label: "Proses" }];

  const handleAdd = () => {
    if (!form.name.trim()) return;
    const u = { ...master };
    if (tab === "customer") u.customers = [...(u.customers || []), form.name.trim()];
    if (tab === "model")    u.models    = [...(u.models    || []), form.name.trim()];
    if (tab === "item")     u.items     = [...(u.items     || []), { name: form.name.trim(), processes: form.processes }];
    if (tab === "proses")   u.processes = [...(u.processes || []), form.name.trim()];
    setMaster(u); STORE.set("master", u);
    setForm({ name: "", processes: [] });
  };

  const handleDel = (type, idx) => {
    const u = { ...master };
    if (type === "customer") u.customers = u.customers.filter((_, i) => i !== idx);
    if (type === "model")    u.models    = u.models.filter((_, i) => i !== idx);
    if (type === "item")     u.items     = u.items.filter((_, i) => i !== idx);
    if (type === "proses")   u.processes = u.processes.filter((_, i) => i !== idx);
    setMaster(u); STORE.set("master", u);
  };

  const listMap = { customer: { key: "customers", type: "customer" }, model: { key: "models", type: "model" }, item: { key: "items", type: "item" }, proses: { key: "processes", type: "proses" } };
  const { key, type } = listMap[tab];
  const list = master[key] || [];

  return (
    <div style={{ paddingBottom: 90 }}>
      {/* Tab Bar */}
      <div style={{ display: "flex", background: T.creamDark, borderRadius: 11, padding: 3, marginBottom: 16, gap: 2 }}>
        {tabs.map(t => (
          <button key={t.id} onClick={() => { setTab(t.id); setForm({ name: "", processes: [] }); }}
            style={{ flex: 1, padding: "8px 0", borderRadius: 9, border: "none", cursor: "pointer", fontFamily: "Montserrat,sans-serif", fontWeight: 600, fontSize: 11, background: tab === t.id ? T.sage : "transparent", color: tab === t.id ? T.white : T.textSoft, transition: "all .2s" }}>
            {t.label}
          </button>
        ))}
      </div>

      {/* Form */}
      <Card>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.sage, marginBottom: 12, textTransform: "uppercase", letterSpacing: .5 }}>
          ＋ Tambah {tabs.find(t => t.id === tab)?.label}
        </div>
        <Input label="Nama" value={form.name} onChange={v => setForm({ ...form, name: v })} placeholder={`Masukkan nama ${tab}...`} />
        {tab === "item" && (
          <MultiProcessInput label="Proses untuk item ini" value={form.processes}
            onChange={v => setForm({ ...form, processes: v })} availableProcesses={master.processes || []} />
        )}
        <Btn onClick={handleAdd} disabled={!form.name.trim()} style={{ width: "100%" }}>
          <Icon name="plus" size={15} color={T.white} /> Simpan
        </Btn>
      </Card>

      {/* List */}
      <Card>
        <div style={{ fontSize: 11, fontWeight: 700, color: T.textSoft, marginBottom: 8, textTransform: "uppercase", letterSpacing: .5 }}>
          Daftar — {list.length} data
        </div>
        {list.length === 0
          ? <div style={{ textAlign: "center", padding: "18px 0", color: T.grayLight, fontSize: 13 }}>Belum ada data</div>
          : list.map((item, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 0", borderBottom: i < list.length - 1 ? `1px solid ${T.creamDark}` : "none" }}>
              <div>
                <div style={{ fontWeight: 600, fontSize: 13 }}>{typeof item === "string" ? item : item.name}</div>
                {typeof item === "object" && item.processes?.length > 0 && (
                  <div style={{ marginTop: 3 }}>{item.processes.map(p => <Tag key={p} label={p} color={T.sageLight + "30"} />)}</div>
                )}
              </div>
              <button onClick={() => handleDel(type, i)} style={{ background: "none", border: "none", cursor: "pointer", padding: 6 }}>
                <Icon name="trash" size={15} color={T.red} />
              </button>
            </div>
          ))
        }
      </Card>
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════════════════
// PAGE: INPUT PRODUKSI
// ═══════════════════════════════════════════════════════════════════════════
const InputProduksi = ({ master, production, setProduction }) => {
  const empty = { date: getToday(), shift: getCurrentShift(), proses: "", customer: "", model: "", item: "", proses_detail: [], ok: "", ng: "", next_process: [] };
  const [form, setForm] = useState(empty);

  const itemProcesses = () => {
    const it = (master.items || []).find(x => (typeof x === "string" ? x : x.name) === form.item);
    return typeof it === "object" ? (it.processes || []) : [];
  };

  const handleSave = () => {
    if (!form.customer || !form.item || form.ok === "") return;
    const entry = { ...form, id: Date.now() };
    const updated = [...production, entry];
    setProduction(updated); STORE.set("production", updated);
    setForm({ ...empty, date: form.date, shift: form.shift });
  };

  const todayShiftEntries = production.filter(p => p.date === form.date && p.shift === form.shift);

  return (
    <div style={{ paddingBottom: 90 }}>
      {/* Date & Shift strip */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <div style={{ flex: 1, background: T.sage, borderRadius: 11, padding: "10px 13px", color: T.white }}>
          <div style={{ fontSize: 9, fontWeight: 600, opacity: .75, textTransform: "uppercase", letterSpacing: .5 }}>Tanggal</div>
          <input type="date" value={form.date} onChange={e => setForm({ ...form, date: e.target.value })}
            style={{ background: "transparent", border: "none", color: T.white, fontFamily: "Montserrat", fontSize: 13, fontWeight: 700, outline: "none", width: "100%", marginTop: 1 }} />
        </div>
        <div style={{ background: T.brown, borderRadius: 11, padding: "10px 13px", color: T.white }}>
          <div style={{ fontSize: 9, fontWeight: 600, opacity: .75, textTransform: "uppercase", letterSpacing: .5 }}>Shift</div>
          <select value={form.shift} onChange={e => setForm({ ...form, shift: e.target.value })}
            style={{ background: "transparent", border: "none", color: T.white, fontFamily: "Montserrat", fontSize: 14, fontWeight: 800, outline: "none", marginTop: 1 }}>
            <option value="1" style={{ color: T.text }}>1</option>
            <option value="2" style={{ color: T.text }}>2</option>
            <option value="3" style={{ color: T.text }}>3</option>
          </select>
        </div>
      </div>

      <Card>
        <div style={{ fontSize: 12, fontWeight: 700, color: T.sage, marginBottom: 12, textTransform: "uppercase", letterSpacing: .5 }}>Form Input Produksi</div>
        <Select label="Proses Awal" value={form.proses} onChange={v => setForm({ ...form, proses: v })} options={(master.processes || []).map(p => ({ value: p, label: p }))} placeholder="Pilih proses..." />
        <Select label="Customer"    value={form.customer} onChange={v => setForm({ ...form, customer: v })} options={(master.customers || []).map(c => ({ value: c, label: c }))} placeholder="Pilih customer..." />
        <Select label="Model"       value={form.model}    onChange={v => setForm({ ...form, model: v })}    options={(master.models    || []).map(m => ({ value: m, label: m }))} placeholder="Pilih model..." />
        <Select label="Item"        value={form.item}     onChange={v => setForm({ ...form, item: v, proses_detail: [], next_process: [] })} options={(master.items || []).map(it => { const n = typeof it === "string" ? it : it.name; return { value: n, label: n }; })} placeholder="Pilih item..." />
        {form.item && <MultiProcessInput label="Detail Proses" value={form.proses_detail} onChange={v => setForm({ ...form, proses_detail: v })} availableProcesses={itemProcesses()} />}

        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: T.textSoft, marginBottom: 4, textTransform: "uppercase", letterSpacing: .5 }}>Hasil OK</div>
            <input type="number" min="0" value={form.ok} onChange={e => setForm({ ...form, ok: e.target.value })} placeholder="0"
              style={{ ...inputStyle, fontSize: 20, fontWeight: 800, color: T.sage, textAlign: "center", border: `2px solid ${T.sage}50`, background: T.sage + "0A" }} />
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, fontWeight: 600, color: T.textSoft, marginBottom: 4, textTransform: "uppercase", letterSpacing: .5 }}>Hasil NG</div>
            <input type="number" min="0" value={form.ng} onChange={e => setForm({ ...form, ng: e.target.value })} placeholder="0"
              style={{ ...inputStyle, fontSize: 20, fontWeight: 800, color: T.red, textAlign: "center", border: `2px solid ${T.red}50`, background: T.red + "0A" }} />
          </div>
        </div>

        <MultiProcessInput label="Next Process" value={form.next_process} onChange={v => setForm({ ...form, next_process: v })} availableProcesses={master.processes || []} />

        <Btn onClick={handleSave} disabled={!form.customer || !form.item || form.ok === ""} style={{ width: "100%", marginTop: 4 }}>
          <Icon name="check" size={15} color={T.white} /> Simpan Produksi
        </Btn>
      </Card>

      {/* Compact recent list */}
      {todayShiftEntries.length > 0 && (
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "10px 13px", borderBottom: `1px solid ${T.creamDark}`, fontSize: 11, fontWeight: 700, color: T.textSoft, textTransform: "uppercase", letterSpacing: .5 }}>
            Input Shift {form.shift} Hari Ini — {todayShiftEntries.length} data
          </div>
          {todayShiftEntries.slice().reverse().map(p => (
            <div key={p.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 13px", borderBottom: `1px solid ${T.creamDark}` }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: T.text }}>{p.item} <span style={{ color: T.textSoft, fontWeight: 400 }}>· {p.customer}</span></div>
              <div style={{ display: "flex", gap: 10 }}>
                <span style={{ fontSize: 13, fontWeight: 800, color: T.sage }}>{p.ok || 0}<span style={{ fontSize: 9, fontWeight: 500, color: T.textSoft }}> OK</span></span>
                <span style={{ fontSize: 13, fontWeight: 800, color: T.red }}>{p.ng || 0}<span style={{ fontSize: 9, fontWeight: 500, color: T.textSoft }}> NG</span></span>
              </div>
            </div>
          ))}
        </Card>
      )}
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════════════════
// EXCEL EXPORT
// ═══════════════════════════════════════════════════════════════════════════
const exportToExcel = (entries, shift, date) => {
  if (!entries.length) return;
  const rows = entries.map((p, i) => ({
    "No": i + 1,
    "Tanggal": p.date,
    "Shift": `Shift ${p.shift}`,
    "Proses": p.proses || "-",
    "Customer": p.customer || "-",
    "Model": p.model || "-",
    "Item": p.item || "-",
    "Detail Proses": (p.proses_detail || []).join(", ") || "-",
    "Hasil OK": parseInt(p.ok) || 0,
    "Hasil NG": parseInt(p.ng) || 0,
    "Total": (parseInt(p.ok) || 0) + (parseInt(p.ng) || 0),
    "Next Process": (p.next_process || []).join(" → ") || "-",
  }));

  const totalOK = entries.reduce((s, p) => s + (parseInt(p.ok) || 0), 0);
  const totalNG = entries.reduce((s, p) => s + (parseInt(p.ng) || 0), 0);
  rows.push({ "No": "", "Tanggal": "", "Shift": "TOTAL", "Proses": "", "Customer": "", "Model": "", "Item": "", "Detail Proses": "", "Hasil OK": totalOK, "Hasil NG": totalNG, "Total": totalOK + totalNG, "Next Process": "" });

  const ws = XLSX.utils.json_to_sheet(rows);

  // Column widths
  ws["!cols"] = [{ wch: 4 }, { wch: 12 }, { wch: 8 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 20 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Shift ${shift}`);
  XLSX.writeFile(wb, `Monitoring_Produksi_${date}_Shift${shift}.xlsx`);
};

// ═══════════════════════════════════════════════════════════════════════════
// PAGE: DASHBOARD — compact table view + logo
// ═══════════════════════════════════════════════════════════════════════════
const Dashboard = ({ production }) => {
  const [selectedDate, setSelectedDate] = useState(getToday());
  const [activeShift, setActiveShift] = useState(getCurrentShift());
  const [closedShifts, setClosedShifts] = useState(() => STORE.get("closedShifts") || {});

  const shiftConfig = [
    { id: "1", label: "Shift 1", time: "06–14", color: T.amber, icon: "sun" },
    { id: "2", label: "Shift 2", time: "14–22", color: T.sage, icon: "sun" },
    { id: "3", label: "Shift 3", time: "22–06", color: T.sageDark, icon: "moon" },
  ];

  const shiftKey = (shift) => `${selectedDate}_shift${shift}`;
  const isClosed = (shift) => !!closedShifts[shiftKey(shift)];

  const closeShift = (shift) => {
    const u = { ...closedShifts, [shiftKey(shift)]: { closedAt: new Date().toISOString() } };
    setClosedShifts(u); STORE.set("closedShifts", u);
  };
  const openShift = (shift) => {
    const { [shiftKey(shift)]: _, ...rest } = closedShifts;
    setClosedShifts(rest); STORE.set("closedShifts", rest);
  };

  const getEntries = (shift) => production.filter(p => p.date === selectedDate && p.shift === shift);
  const sum = (entries) => ({
    ok: entries.reduce((s, p) => s + (parseInt(p.ok) || 0), 0),
    ng: entries.reduce((s, p) => s + (parseInt(p.ng) || 0), 0),
    n: entries.length,
  });

  const allEntries = production.filter(p => p.date === selectedDate);
  const grandSum = sum(allEntries);
  const activeShiftColor = shiftConfig.find(s => s.id === activeShift)?.color || T.sage;

  return (
    <div style={{ paddingBottom: 90 }}>

      {/* ── HEADER CARD with LOGO ── */}
      <div style={{ background: `linear-gradient(135deg, ${T.sageDark} 0%, ${T.sage} 60%, ${T.sageLight} 100%)`, borderRadius: 16, padding: "16px", marginBottom: 14, position: "relative", overflow: "hidden" }}>
        {/* Decorative circle */}
        <div style={{ position: "absolute", right: -20, top: -20, width: 120, height: 120, borderRadius: "50%", background: T.white + "10" }} />
        <div style={{ position: "absolute", right: 20, bottom: -30, width: 80, height: 80, borderRadius: "50%", background: T.white + "08" }} />

        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 14 }}>
          <Logo size={44} />
          <div>
            <div style={{ fontSize: 9, fontWeight: 600, color: T.sageLight, textTransform: "uppercase", letterSpacing: 1.5 }}>Monitoring Produksi</div>
            <div style={{ fontSize: 18, fontWeight: 800, color: T.white, letterSpacing: -.3, lineHeight: 1.2 }}>Dashboard</div>
          </div>
          <div style={{ marginLeft: "auto" }}>
            <input type="date" value={selectedDate} onChange={e => setSelectedDate(e.target.value)}
              style={{ background: T.white + "20", border: `1px solid ${T.white}40`, borderRadius: 8, color: T.white, fontFamily: "Montserrat", fontSize: 11, fontWeight: 700, outline: "none", padding: "6px 8px" }} />
          </div>
        </div>

        {/* Grand Total Chips */}
        <div style={{ display: "flex", gap: 8 }}>
          {[
            { label: "Total Input", val: grandSum.n, bg: T.white + "20" },
            { label: "Total OK", val: grandSum.ok, bg: T.sage + "50" },
            { label: "Total NG", val: grandSum.ng, bg: T.red + "60" },
          ].map(m => (
            <div key={m.label} style={{ flex: 1, background: m.bg, borderRadius: 9, padding: "8px 6px", textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 800, color: T.white }}>{m.val}</div>
              <div style={{ fontSize: 9, color: T.white + "CC", fontWeight: 600, textTransform: "uppercase" }}>{m.label}</div>
            </div>
          ))}
        </div>
      </div>

      {/* ── SHIFT TABS ── */}
      <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
        {shiftConfig.map(s => {
          const e = getEntries(s.id);
          const sm = sum(e);
          const active = activeShift === s.id;
          return (
            <button key={s.id} onClick={() => setActiveShift(s.id)}
              style={{ flex: 1, padding: "10px 4px", borderRadius: 11, border: active ? `2px solid ${s.color}` : "2px solid transparent", cursor: "pointer", background: active ? s.color : T.white, fontFamily: "Montserrat", transition: "all .2s", boxShadow: active ? `0 4px 14px ${s.color}40` : "0 1px 4px rgba(0,0,0,.07)" }}>
              <div style={{ fontSize: 12, fontWeight: 800, color: active ? T.white : T.text }}>{s.label}</div>
              <div style={{ fontSize: 9, color: active ? T.white + "CC" : T.textSoft, marginTop: 1 }}>{s.time}</div>
              <div style={{ marginTop: 4, fontSize: 11, fontWeight: 700, color: active ? T.white : T.sage }}>{sm.ok}<span style={{ fontWeight: 400, fontSize: 9 }}>ok</span> <span style={{ color: active ? T.white + "BB" : T.red }}>{sm.ng}<span style={{ fontWeight: 400, fontSize: 9 }}>ng</span></span></div>
              {isClosed(s.id) && <div style={{ fontSize: 8, color: active ? T.white + "CC" : T.amber, fontWeight: 700, marginTop: 2 }}>✓ SELESAI</div>}
            </button>
          );
        })}
      </div>

      {/* ── ACTIVE SHIFT DETAIL ── */}
      {shiftConfig.map(s => {
        if (s.id !== activeShift) return null;
        const entries = getEntries(s.id);
        const sm = sum(entries);
        const closed = isClosed(s.id);
        const closedAt = closedShifts[shiftKey(s.id)]?.closedAt;

        return (
          <div key={s.id}>
            {/* Summary row + actions */}
            <Card style={{ padding: "10px 12px" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                <div style={{ display: "flex", gap: 8, flex: 1 }}>
                  {[{ l: "Input", v: sm.n, c: T.brown }, { l: "OK", v: sm.ok, c: T.sage }, { l: "NG", v: sm.ng, c: T.red }].map(m => (
                    <div key={m.l} style={{ textAlign: "center", flex: 1 }}>
                      <div style={{ fontSize: 18, fontWeight: 800, color: m.c }}>{m.v}</div>
                      <div style={{ fontSize: 9, color: T.textSoft, fontWeight: 600, textTransform: "uppercase" }}>{m.l}</div>
                    </div>
                  ))}
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 5, alignItems: "flex-end" }}>
                  {/* Download Excel */}
                  <Btn small variant="green" onClick={() => exportToExcel(entries, s.id, selectedDate)} disabled={entries.length === 0}>
                    <Icon name="download" size={12} color={T.white} /> Excel
                  </Btn>
                  {/* Close/Open Shift */}
                  {closed
                    ? <Btn small variant="ghost" onClick={() => openShift(s.id)}><Icon name="unlock" size={11} /> Buka</Btn>
                    : <Btn small variant="amber" onClick={() => closeShift(s.id)}><Icon name="lock" size={11} color={T.white} /> Selesai</Btn>
                  }
                </div>
              </div>
              {closed && closedAt && (
                <div style={{ marginTop: 7, padding: "5px 8px", background: T.sage + "15", borderRadius: 7, fontSize: 10, color: T.sage, fontWeight: 600 }}>
                  ✅ Shift selesai · Ditutup {new Date(closedAt).toLocaleTimeString("id-ID")}
                </div>
              )}
            </Card>

            {/* Compact Table — fits in 1 screen */}
            {entries.length === 0 ? (
              <div style={{ textAlign: "center", padding: "36px 0", color: T.grayLight, fontSize: 13 }}>
                <Icon name="clipboard" size={36} color={T.grayLight} />
                <div style={{ marginTop: 8 }}>Belum ada input shift ini</div>
              </div>
            ) : (
              <Card style={{ padding: 0, overflow: "hidden" }}>
                {/* Table header */}
                <div style={{ display: "grid", gridTemplateColumns: "24px 1fr 60px 42px 42px", gap: 0, background: activeShiftColor, padding: "7px 10px" }}>
                  {["#", "Item / Customer", "Proses", "OK", "NG"].map(h => (
                    <div key={h} style={{ fontSize: 10, fontWeight: 700, color: T.white, textAlign: h === "OK" || h === "NG" ? "center" : "left", textTransform: "uppercase", letterSpacing: .4 }}>{h}</div>
                  ))}
                </div>
                {/* Rows */}
                {entries.map((p, i) => (
                  <div key={p.id || i} style={{ display: "grid", gridTemplateColumns: "24px 1fr 60px 42px 42px", gap: 0, padding: "7px 10px", background: i % 2 === 0 ? T.white : T.cream, borderBottom: `1px solid ${T.creamDark}` }}>
                    <div style={{ fontSize: 10, color: T.textSoft, fontWeight: 600 }}>{i + 1}</div>
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: T.text, lineHeight: 1.2 }}>{p.item}</div>
                      <div style={{ fontSize: 10, color: T.textSoft }}>{p.customer}{p.model ? ` · ${p.model}` : ""}</div>
                      {p.next_process?.length > 0 && (
                        <div style={{ fontSize: 9, color: T.brown, fontWeight: 600, marginTop: 2 }}>→ {p.next_process.join(" · ")}</div>
                      )}
                    </div>
                    <div style={{ fontSize: 10, color: T.brownDark, fontWeight: 600, wordBreak: "break-word", lineHeight: 1.3 }}>{p.proses || "-"}</div>
                    <div style={{ fontSize: 14, fontWeight: 800, color: T.sage, textAlign: "center" }}>{p.ok || 0}</div>
                    <div style={{ fontSize: 14, fontWeight: 800, color: T.red, textAlign: "center" }}>{p.ng || 0}</div>
                  </div>
                ))}
                {/* Total row */}
                <div style={{ display: "grid", gridTemplateColumns: "24px 1fr 60px 42px 42px", gap: 0, padding: "7px 10px", background: activeShiftColor + "20", borderTop: `2px solid ${activeShiftColor}` }}>
                  <div />
                  <div style={{ fontSize: 11, fontWeight: 800, color: T.text }}>TOTAL</div>
                  <div />
                  <div style={{ fontSize: 14, fontWeight: 800, color: T.sage, textAlign: "center" }}>{sm.ok}</div>
                  <div style={{ fontSize: 14, fontWeight: 800, color: T.red, textAlign: "center" }}>{sm.ng}</div>
                </div>
              </Card>
            )}
          </div>
        );
      })}
    </div>
  );
};

// ═══════════════════════════════════════════════════════════════════════════
// ROOT
// ═══════════════════════════════════════════════════════════════════════════
export default function App() {
  const [page, setPage] = useState("dashboard");
  const [master, setMaster] = useState(initMaster);
  const [production, setProduction] = useState(initProduction);

  const nav = [
    { id: "master",   label: "Master",   icon: "settings" },
    { id: "produksi", label: "Produksi", icon: "clipboard" },
    { id: "dashboard",label: "Dashboard",icon: "chart" },
  ];

  const pageTitles = { master: "Master Data", produksi: "Input Produksi", dashboard: "Dashboard" };

  return (
    <div style={{ fontFamily: "Montserrat,sans-serif", background: T.cream, minHeight: "100vh", maxWidth: 480, margin: "0 auto" }}>
      {/* Header — hidden on dashboard (dashboard has its own logo card) */}
      {page !== "dashboard" && (
        <div style={{ background: `linear-gradient(135deg, ${T.sageDark} 0%, ${T.sage} 100%)`, padding: "16px 18px 13px", position: "sticky", top: 0, zIndex: 100, boxShadow: "0 2px 16px rgba(0,0,0,.14)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <Logo size={34} />
            <div>
              <div style={{ fontSize: 9, fontWeight: 600, color: T.sageLight, textTransform: "uppercase", letterSpacing: 1.5 }}>Monitoring Produksi</div>
              <div style={{ fontSize: 17, fontWeight: 800, color: T.white, letterSpacing: -.2 }}>{pageTitles[page]}</div>
            </div>
            <div style={{ marginLeft: "auto", fontSize: 10, background: T.brown + "70", color: T.white, padding: "3px 9px", borderRadius: 20, fontWeight: 600 }}>
              Shift {getCurrentShift()}
            </div>
          </div>
        </div>
      )}

      {/* Body */}
      <div style={{ padding: page === "dashboard" ? "14px 14px 0" : "16px 14px 0" }}>
        {page === "master"    && <MasterData master={master} setMaster={setMaster} />}
        {page === "produksi"  && <InputProduksi master={master} production={production} setProduction={setProduction} />}
        {page === "dashboard" && <Dashboard production={production} />}
      </div>

      {/* Bottom Nav */}
      <div style={{ position: "fixed", bottom: 0, left: "50%", transform: "translateX(-50%)", width: "100%", maxWidth: 480, background: T.white, borderTop: `1px solid ${T.creamDark}`, display: "flex", zIndex: 200, boxShadow: "0 -3px 16px rgba(0,0,0,.09)" }}>
        {nav.map(n => (
          <button key={n.id} onClick={() => setPage(n.id)}
            style={{ flex: 1, padding: "11px 0 9px", border: "none", background: "transparent", cursor: "pointer", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, transition: "all .2s" }}>
            <div style={{ padding: page === n.id ? "5px 14px" : "5px 10px", borderRadius: 18, background: page === n.id ? T.sage + "20" : "transparent", transition: "all .2s" }}>
              <Icon name={n.icon} size={21} color={page === n.id ? T.sage : T.gray} />
            </div>
            <span style={{ fontSize: 9, fontWeight: 700, color: page === n.id ? T.sage : T.gray, letterSpacing: .3, textTransform: "uppercase" }}>{n.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
