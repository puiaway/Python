import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import OrderedDict, defaultdict
import csv
import json

"""
ATM CMD‑V4 Trace Viewer + CNG_DISP_STANDARD Reporter
----------------------------------------------------
This GUI loads ATM dispenser trace XML (EVENT/COMMAND/ENTRY), parses values,
shows meanings (glossary), and can export grouped reports for CNG_DISP_STANDARD.

• Browse, filter, sort, inspect rows
• Meanings panel with glossary (built‑in + load/save JSON)
• File → Export → CNG_DISP_STANDARD Report…
    - Produces two CSVs: items (per line) and transaction summary (per element)

Run:  python atm_cmdv4_viewer_with_cng_report.py
"""

TAGS_TO_SHOW = ("EVENT", "COMMAND", "ENTRY")
TRUNCATE_AT = 48

# --- Built-in minimal glossary (extend or override via Help → Load Glossary JSON…) ---
BUILTIN_GLOSSARY = {
    # Common fields
    "LEN": {"label": "Length", "meaning": "Length of payload/status dataset (bytes)", "manual_ref": "Header format"},
    "RSTA": {"label": "Result Status", "meaning": "Device/job status (A/OK …)", "manual_ref": "Appendix: Status ranges"},
    "RACT": {"label": "Result Action", "meaning": "Last action code executed", "manual_ref": "Appendix: RACT"},
    "RRET": {"label": "Return Code", "meaning": "00=OK, else error code", "manual_ref": "Appendix: Return codes"},
    # Cassette n=1..n
    "nSTA": {"label": "Cassette n Status", "meaning": "Cassette ready/missing/low", "manual_ref": "Cassette status"},
    "nNUM": {"label": "Cassette n Notes", "meaning": "Notes counted in cassette", "manual_ref": "Cassette counters"},
    "nVAL": {"label": "Cassette n Denomination", "meaning": "Value in minor currency units", "manual_ref": "Note params (DFP/DFD)"},
    "nREJ": {"label": "Cassette n Rejects", "meaning": "Reject counter", "manual_ref": "Cassette counters"},
    # STDV examples
    "S_SW": {"label": "Service Switch", "meaning": "Service switch state", "manual_ref": "Device status"},
    "LCMD": {"label": "Last Command", "meaning": "Code of last executed command", "manual_ref": "Cross reference"},
    "LSTA": {"label": "Last Status", "meaning": "Result of last command", "manual_ref": "Device status"},
    "C_OUT": {"label": "Cash Output", "meaning": "Cash output path/lock flag", "manual_ref": "Device status"},
    # CNG_DISP_STANDARD line meanings (positional)
    "CNG_DISP_STANDARD": {"label": "Dispense Mix", "meaning": "F1=count, F2=cassette, F3=amount(minor), F4=flag", "manual_ref": "DBS – Dispensing"},
}

FLAG_MEANINGS = {
    "O": "OK / Offered",
    "R": "Rejected",
    "E": "Error",
    "C": "Cancelled",
}

# ----------------------------- parsing helpers -----------------------------

def strip_ns(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def smart_split_payload(text: str):
    if text is None:
        return []
    return [s.strip() for s in text.split(";") if s.strip()]


def parse_segment(segment: str) -> OrderedDict:
    out = OrderedDict()
    segment = (segment or "").strip()
    if not segment:
        return out
    parts = [p.strip() for p in segment.split(",") if p.strip() != ""]
    if any("=" in p for p in parts):
        # key=value mode
        seen = defaultdict(int)
        for p in parts:
            if "=" in p:
                k, v = p.split("=", 1)
                k, v = k.strip(), v.strip()
            else:
                seen["F"] += 1
                k, v = f"F{seen['F']}", p
            seen[k] += 1
            key = k if seen[k] == 1 else f"{k}#{seen[k]}"
            out[key] = v
    else:
        # positional fields
        for i, v in enumerate(parts, 1):
            out[f"F{i}"] = v
    return out

# ----------------------------- data model -----------------------------
class Row:
    __slots__ = ("tag", "type", "date", "time", "segment", "fields", "attrib", "raw")
    def __init__(self, tag, etype, date, time, segment, fields, attrib, raw):
        self.tag = tag
        self.type = etype
        self.date = date
        self.time = time
        self.segment = segment
        self.fields = fields
        self.attrib = attrib
        self.raw = raw

# ----------------------------- reporter logic -----------------------------

def detect_cng_disp(elem_type: str, payload: str) -> bool:
    if not payload and not elem_type:
        return False
    t = (elem_type or "").upper()
    p = (payload or "").upper()
    return t == "CNG_DISP_STANDARD" or "CNG_DISP_STANDARD" in p


def parse_cng_line(line: str):
    parts = [p.strip() for p in (line or "").split(",") if p.strip()]
    if not parts:
        return None
    if parts[0].upper() == "CNG_DISP_STANDARD":
        parts = parts[1:]
    if len(parts) < 4:
        return None
    f1, f2, f3, f4 = parts[:4]
    def to_int(x):
        try:
            return int(x)
        except Exception:
            return None
    return {
        "F1_NOTE_COUNT": to_int(f1) if to_int(f1) is not None else f1,
        "F2_CASSETTE_ID": to_int(f2) if to_int(f2) is not None else f2,
        "F3_AMOUNT_MINOR": to_int(f3) if to_int(f3) is not None else f3,
        "F4_LINE_RESULT": f4,
        "RAW_FIELDS": parts,
    }


def export_cng_reports(xml_path: Path, save_dir: Path):
    items = []
    txn_summaries = defaultdict(lambda: {"date": "", "time": "", "total_notes": 0, "total_amount_minor": 0, "by_cassette": defaultdict(lambda: {"notes": 0, "amount_minor": 0})})
    txn_counter = 0

    for _, elem in ET.iterparse(xml_path, events=("end",)):
        tag = strip_ns(elem.tag)
        if tag not in TAGS_TO_SHOW:
            elem.clear(); continue
        etype = elem.attrib.get("type", tag)
        payload = (elem.text or "").strip()
        if not detect_cng_disp(etype, payload):
            elem.clear(); continue

        date = elem.attrib.get("date", "")
        time = elem.attrib.get("time", "")
        txn_counter += 1
        txn_id = f"{date}T{time}#{txn_counter}"

        for seg in smart_split_payload(payload):
            parsed = parse_cng_line(seg)
            if not parsed:
                continue
            count = parsed["F1_NOTE_COUNT"]
            cass = parsed["F2_CASSETTE_ID"]
            amt = parsed["F3_AMOUNT_MINOR"]
            flag = parsed["F4_LINE_RESULT"]
            flag_mean = FLAG_MEANINGS.get(str(flag).upper(), "—")
            items.append({
                "txn_id": txn_id,
                "date": date,
                "time": time,
                "element_tag": tag,
                "element_type": etype,
                "F1_NOTE_COUNT": count,
                "F1_meaning": "Number of notes on this line",
                "F2_CASSETTE_ID": cass,
                "F2_meaning": "Cassette source (1..n)",
                "F3_AMOUNT_MINOR": amt,
                "F3_meaning": "Amount in minor currency units",
                "F4_LINE_RESULT": flag,
                "F4_meaning": flag_mean,
            })
            if isinstance(count, int):
                txn_summaries[txn_id]["total_notes"] += count
                txn_summaries[txn_id]["by_cassette"][cass]["notes"] += count
            if isinstance(amt, int):
                txn_summaries[txn_id]["total_amount_minor"] += amt
                txn_summaries[txn_id]["by_cassette"][cass]["amount_minor"] += amt
            txn_summaries[txn_id]["date"] = date
            txn_summaries[txn_id]["time"] = time
        elem.clear()

    # Write CSVs
    items_path = save_dir / "cng_disp_items.csv"
    sum_path = save_dir / "cng_disp_txn_summary.csv"

    if items:
        keys = [
            "txn_id","date","time","element_tag","element_type",
            "F1_NOTE_COUNT","F1_meaning","F2_CASSETTE_ID","F2_meaning",
            "F3_AMOUNT_MINOR","F3_meaning","F4_LINE_RESULT","F4_meaning"
        ]
        with open(items_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader(); w.writerows(items)
    if txn_summaries:
        rows = []
        for txn_id, s in txn_summaries.items():
            parts = []
            for c_id, cs in sorted(s["by_cassette"].items(), key=lambda kv: str(kv[0])):
                parts.append(f"C{c_id}: notes={cs['notes']}, amt_minor={cs['amount_minor']}")
            rows.append({
                "txn_id": txn_id,
                "date": s["date"],
                "time": s["time"],
                "total_notes": s["total_notes"],
                "total_amount_minor": s["total_amount_minor"],
                "by_cassette": " | "+" | ".join(parts) if parts else "",
            })
        keys2 = ["txn_id","date","time","total_notes","total_amount_minor","by_cassette"]
        with open(sum_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys2)
            w.writeheader(); w.writerows(rows)

    return items_path if items else None, sum_path if txn_summaries else None

# ----------------------------- GUI -----------------------------
class App(tk.Tk):
    def __init__(self, xml_path: str | None = None):
        super().__init__()
        self.title("ATM CMD‑V4 Trace Viewer + CNG Reporter")
        self.geometry("1450x860")

        self.all_rows: list[Row] = []
        self.glossary = dict(BUILTIN_GLOSSARY)
        self.current_xml: Path | None = Path(xml_path) if xml_path else None
        self._sort_state = {"col": None, "asc": True}

        self._build_menu()
        self._build_toolbar()
        self._build_body()
        self._build_status()

        if self.current_xml and self.current_xml.exists():
            try:
                self.load_xml(self.current_xml)
            except Exception as e:
                messagebox.showerror("Load error", str(e))

    # ---- UI builders ----
    def _build_menu(self):
        m = tk.Menu(self)
        mf = tk.Menu(m, tearoff=0)
        mf.add_command(label="Open XML…", command=self.on_open)
        # Export submenu
        mx = tk.Menu(mf, tearoff=0)
        mx.add_command(label="CNG_DISP_STANDARD Report…", command=self.on_export_cng)
        mf.add_cascade(label="Export", menu=mx)
        mf.add_separator()
        mf.add_command(label="Exit", command=self.destroy)
        m.add_cascade(label="File", menu=mf)

        mh = tk.Menu(m, tearoff=0)
        mh.add_command(label="Load Glossary JSON…", command=self.on_load_glossary)
        mh.add_command(label="Save Current Glossary as JSON…", command=self.on_save_glossary)
        mh.add_separator()
        mh.add_command(label="About meanings…", command=lambda: messagebox.showinfo("About", "Meanings are based on the CMD‑V4 manual. Extend/override via JSON."))
        m.add_cascade(label="Help", menu=mh)
        self.config(menu=m)

    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(8,6))
        bar.pack(fill="x")
        ttk.Label(bar, text="Filter:").pack(side="left")
        self.var_filter = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.var_filter, width=50)
        e.pack(side="left", padx=(6, 12))
        self.var_filter.trace_add("write", lambda *_: self.refresh_view())

        ttk.Label(bar, text="Tag:").pack(side="left")
        self.var_tag = tk.StringVar(value="All")
        cb_tag = ttk.Combobox(bar, textvariable=self.var_tag, state="readonly", width=10)
        cb_tag["values"] = ("All",) + TAGS_TO_SHOW
        cb_tag.pack(side="left", padx=(6, 12))
        cb_tag.bind("<<ComboboxSelected>>", lambda _e: self.on_tag_changed())

        ttk.Label(bar, text="Type:").pack(side="left")
        self.var_type = tk.StringVar(value="All")
        self.cb_type = ttk.Combobox(bar, textvariable=self.var_type, state="readonly", width=20)
        self.cb_type.pack(side="left", padx=(6, 12))
        self.cb_type.bind("<<ComboboxSelected>>", lambda _e: self.refresh_view())

        ttk.Button(bar, text="Open XML…", command=self.on_open).pack(side="right")
        ttk.Button(bar, text="Export CNG Report…", command=self.on_export_cng).pack(side="right", padx=(0,8))

    def _build_body(self):
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)

        # Left table
        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="nsew")
        self.tree = ttk.Treeview(left, columns=("date","time","tag","type","segment"), show="headings")
        for col, w in ("date",120), ("time",110), ("tag",90), ("type",180), ("segment",90):
            self.tree.heading(col, text=col.upper(), command=lambda c=col: self.sort_by(c))
            self.tree.column(col, width=w, anchor="w")
        vsb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        self.tree.bind("<<TreeviewSelect>>", self.on_selection)
        self.tree.bind("<Double-1>", self.on_double_click)

        # Right panel
        right = ttk.Frame(body, padding=(10,8))
        right.grid(row=0, column=1, sticky="nsew")
        ttk.Label(right, text="Parsed Fields", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.kv = tk.Text(right, height=18, wrap="none"); self.kv.configure(state="disabled"); self.kv.pack(fill="x", pady=(2,8))
        ttk.Label(right, text="Parameter Glossary", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.mean = tk.Text(right, wrap="word"); self.mean.configure(state="disabled"); self.mean.pack(fill="both", expand=True)

        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

    def _build_status(self):
        self.var_status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.var_status, anchor="w").pack(fill="x")

    # ---- file/glossary actions ----
    def on_open(self):
        path = filedialog.askopenfilename(title="Open Trace XML", filetypes=[("XML files","*.xml *.XML"),("All files","*.*")])
        if path:
            self.load_xml(path)

    def on_export_cng(self):
        if not self.current_xml or not Path(self.current_xml).exists():
            messagebox.showwarning("CNG Report", "Open an XML file first.")
            return
        directory = filedialog.askdirectory(title="Select folder to save CNG reports")
        if not directory:
            return
        items_path, sum_path = export_cng_reports(Path(self.current_xml), Path(directory))
        if items_path or sum_path:
            msg = "Exported:"
            if items_path: msg += f"\n- {items_path}"
            if sum_path: msg += f"\n- {sum_path}"
            messagebox.showinfo("CNG Report", msg)
        else:
            messagebox.showinfo("CNG Report", "No CNG_DISP_STANDARD transactions found.")

    def on_load_glossary(self):
        path = filedialog.askopenfilename(title="Load Glossary JSON", filetypes=[("JSON","*.json")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                raise ValueError("Glossary root must be an object.")
            self.glossary.update(data)
            messagebox.showinfo("Glossary", f"Loaded {len(data)} entries.")
            self.on_selection()
        except Exception as e:
            messagebox.showerror("Glossary", str(e))

    def on_save_glossary(self):
        path = filedialog.asksaveasfilename(title="Save Glossary JSON", defaultextension=".json", filetypes=[("JSON","*.json")], initialfile="cmdv4_glossary.json")
        if not path: return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.glossary, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Glossary", f"Saved to {path}")
        except Exception as e:
            messagebox.showerror("Glossary", str(e))

    # ---- loading & view ----
    def load_xml(self, path: str | Path):
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(str(p))
        self.current_xml = p
        self.title(f"ATM CMD‑V4 Trace Viewer + CNG Reporter — {p.name}")
        self.var_status.set(f"Loading {p} …")
        self.all_rows.clear(); self.tree.delete(*self.tree.get_children())
        self._set_type_choices([])

        try:
            for _, elem in ET.iterparse(p, events=("end",)):
                tag = strip_ns(elem.tag)
                if tag not in TAGS_TO_SHOW:
                    elem.clear(); continue
                date = elem.attrib.get("date", "")
                time = elem.attrib.get("time", "")
                etype = elem.attrib.get("type", tag)
                payload = (elem.text or "").strip()
                segments = smart_split_payload(payload) if payload else [""]
                if segments == [""] and list(elem):
                    for child in elem:
                        ctag = strip_ns(child.tag)
                        if ctag not in TAGS_TO_SHOW: continue
                        ctype = child.attrib.get("type", ctag)
                        cdate = child.attrib.get("date", date)
                        ctime = child.attrib.get("time", time)
                        ctext = (child.text or "").strip()
                        for si, seg in enumerate(smart_split_payload(ctext) if ctext else [""], start=1):
                            fields = parse_segment(seg)
                            self.all_rows.append(Row(ctag, ctype, cdate, ctime, si, fields, dict(child.attrib), seg))
                else:
                    for si, seg in enumerate(segments, start=1):
                        fields = parse_segment(seg)
                        self.all_rows.append(Row(tag, etype, date, time, si, fields, dict(elem.attrib), seg))
                elem.clear()
        except ET.ParseError as e:
            messagebox.showerror("XML Parse Error", str(e)); self.var_status.set("Parse error"); return

        self._refresh_type_choices(); self.refresh_view()
        self.var_status.set(f"Loaded {len(self.all_rows):,} rows from {p.name}")

    def _refresh_type_choices(self):
        self.on_tag_changed()

    def on_tag_changed(self):
        tag = self.var_tag.get()
        types = sorted({r.type for r in self.all_rows if tag == "All" or r.tag == tag})
        self._set_type_choices(["All"] + types)
        self.var_type.set("All"); self.refresh_view()

    def _set_type_choices(self, items):
        if not items: items = ["All"]
        self.cb_type["values"] = items
        if self.var_type.get() not in items:
            self.var_type.set(items[0])

    def _current_rows(self):
        q = (self.var_filter.get() or "").lower()
        tag_sel = self.var_tag.get(); type_sel = self.var_type.get()
        out = []
        for r in self.all_rows:
            if tag_sel != "All" and r.tag != tag_sel: continue
            if type_sel != "All" and r.type != type_sel: continue
            hay = " ".join([r.date or "", r.time or "", r.tag or "", r.type or "", r.raw or "", " ".join(f"{k}={v}" for k,v in r.fields.items())]).lower()
            if q and q not in hay: continue
            out.append(r)
        return out

    def refresh_view(self):
        rows = self._current_rows()
        field_keys, seen = [], set()
        for r in rows:
            for k in r.fields.keys():
                if k not in seen: seen.add(k); field_keys.append(k)
        columns = ("date","time","tag","type","segment") + tuple(field_keys)
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col.upper(), command=lambda c=col: self.sort_by(c))
            width = 120 if col in ("date","time") else 90 if col in ("tag","segment") else 150
            self.tree.column(col, width=width, anchor="w")
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            vals = [r.date, r.time, r.tag, r.type, r.segment] + [r.fields.get(k, "") for k in field_keys]
            self.tree.insert("", "end", values=vals)
        self.var_status.set(f"Showing {len(rows):,} row(s)")
        self.on_selection()

    def sort_by(self, column: str):
        items = [(iid, self.tree.item(iid, "values")) for iid in self.tree.get_children("")]
        if not items: return
        cols = self.tree["columns"]
        if column not in cols: return
        idx = cols.index(column)
        asc = True if self._sort_state["col"] != column else not self._sort_state["asc"]
        self._sort_state.update({"col": column, "asc": asc})
        def coerce(x):
            try: return float(x)
            except Exception: return str(x)
        items.sort(key=lambda it: coerce(it[1][idx]), reverse=not asc)
        for i, (iid, _) in enumerate(items): self.tree.move(iid, "", i)

    def on_selection(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            self._set_kv_text(""); self._set_mean_text("Select a row to see parsed fields and their meanings."); return
        vals = self.tree.item(sel[0], "values"); cols = self.tree["columns"]
        data = OrderedDict((c, v) for c, v in zip(cols, vals))
        self._set_kv_text("\n".join(f"{k}: {v}" for k, v in data.items()))
        # meanings only for dynamic value columns
        mean_lines = []
        for k, v in data.items():
            if k in ("date","time","tag","type","segment"): continue
            mean_lines.append(self._describe_key(k, v))
        if not mean_lines:
            mean_lines.append("No glossary entries found for this row. Load a glossary JSON to extend definitions.")
        self._set_mean_text("\n\n".join(mean_lines))

    def on_double_click(self, _evt=None):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0], "values"); cols = self.tree["columns"]
        data = OrderedDict((c, v) for c, v in zip(cols, vals))
        win = tk.Toplevel(self); win.title("Row details"); win.geometry("900x600")
        txt = tk.Text(win, wrap="word"); txt.pack(fill="both", expand=True)
        txt.insert("1.0", "Details (all columns):\n\n")
        for k, v in data.items(): txt.insert("end", f"- {k}: {v}\n")
        txt.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=6)

    def _set_kv_text(self, text: str):
        self.kv.configure(state="normal"); self.kv.delete("1.0","end"); self.kv.insert("1.0", text); self.kv.configure(state="disabled")

    def _set_mean_text(self, text: str):
        self.mean.configure(state="normal"); self.mean.delete("1.0","end"); self.mean.insert("1.0", text); self.mean.configure(state="disabled")

    def _describe_key(self, k: str, v: str) -> str:
        base = k
        if len(k) >= 2 and k[0].isdigit():
            i = 0
            while i < len(k) and k[i].isdigit(): i += 1
            base = "n" + k[i:]
        info = self.glossary.get(base) or self.glossary.get(k)
        if not info:
            # also provide default for positional fields of CNG_DISP_STANDARD
            if base.startswith("F") and self.var_type.get().upper() == "CNG_DISP_STANDARD":
                labels = {"F1": "NOTE_COUNT", "F2": "CASSETTE_ID", "F3": "AMOUNT_MINOR", "F4": "LINE_RESULT"}
                label = labels.get(base, base)
                if base == "F4":
                    fmean = FLAG_MEANINGS.get(str(v).upper(), "—")
                    return f"{k}: {v}\n  • {label}: Result flag\n  • {fmean}"
                else:
                    extra = "Amount in minor currency units" if base == "F3" else ("Cassette source (1..n)" if base == "F2" else "Number of notes on this line")
                    return f"{k}: {v}\n  • {label}: {extra}"
            return f"{k}: {v} — (no definition)"
        label = info.get("label", k); meaning = info.get("meaning", ""); mref = info.get("manual_ref", "")
        return f"{k}: {v}\n  • {label}: {meaning}\n  • {mref}"

# ----------------------------- main -----------------------------
if __name__ == "__main__":
    # Auto-open a local sample if it sits next to this script (optional)
    default_xml = Path(__file__).with_name("20241225.TRC.XML")
    app = App(default_xml if default_xml.exists() else None)
    app.mainloop()
