#!/usr/bin/env python3
"""f2p: turn a FlowJo workspace (.wsp) into one Excel workbook.

Usage:  f2p.py [--no-dialog] [PATH ...]      (default: current folder)
        f2p.py --settings                     (open settings.ini in a text editor)

PATH is a .wsp file or a folder containing .wsp files. Each FlowJo group is one
plate; Freq. of Parent / #cells come from the gate counts, MFI from Median/Mean
statistics saved in the workspace. A dialog lets you tick variables, groups and
gates (gates matching live / CD107a / 107a / IFN / TNF and all groups by default).

Output next to the .wsp, <experiment name>.xlsx (EXPERIMENT NAME keyword;
falls back to <wsp name with yyyymmdd dates>_f2p.xlsx), with
  - "Plates": one well grid per gate, gates side by side, plates stacked down
  - "List":   wells down, one column per plate, grouped by gate
  - "Info":   acquisition, instrument and panel metadata
"""
import configparser
import datetime as _dt
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

WELL_RE = re.compile(r"([A-P])0*(\d{1,2})(?:\.fcs)?\s*$", re.I)
# live (%) below limit -> colour: muted terracotta, clay, sand
LIVE_BANDS = [(50, "D4A496"), (80, "E6C3A1"), (90, "F1E5C6")]
LIVE_RE = re.compile(r"^(live|alive|viable|viability)\b", re.I)
DEFAULT_GATES = [LIVE_RE, re.compile(r"107a", re.I), re.compile(r"ifn", re.I),
                 re.compile(r"tnf", re.I)]
SKIP_GROUPS = {"all samples", "compensation"}
GRADIENT = ("FFFFFF", "F8A35E")
DEFAULT_VARS = {"frequency": True, "cells": False, "mfi": False}
MIN_PRESENCE = 50
HERE = Path(__file__).resolve().parent
SETTINGS = Path(os.environ.get("F2P_SETTINGS", HERE / "settings.ini"))
FREQ = "Freq. of Parent (%)"
COUNT = "#cells"
MFI_STATS = {"median", "mean", "geometric mean", "geom. mean", "cv", "robust cv"}
MONTHS = {m: i for i, m in enumerate(
    "jan feb mar apr may jun jul aug sep oct nov dec".split(), 1)}

BOLD = Font(bold=True)
TITLE = Font(bold=True, size=14)
GATE_FONT = Font(italic=True, size=9, color="6B7785")
HEAD_FILL = PatternFill("solid", fgColor="E7EEF7")
THIN = Side(style="thin", color="B7C3D0")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center")


# ---------------------------------------------------------------- settings
def _list(text):
    return [x.strip() for x in (text or "").split(",") if x.strip()]


def _hex(text, fallback):
    t = (text or "").strip().lstrip("#").upper()
    return t if re.fullmatch(r"[0-9A-F]{6}", t) else fallback


def load_settings():
    """Read settings.ini next to f2p.py (created from settings_default.ini if missing)."""
    global LIVE_BANDS, LIVE_RE, DEFAULT_GATES, SKIP_GROUPS, GRADIENT, MIN_PRESENCE
    default = HERE / "settings_default.ini"
    if not SETTINGS.exists() and default.exists():
        try:
            shutil.copy(default, SETTINGS)
        except OSError:
            pass
    cfg = configparser.ConfigParser()
    try:
        cfg.read([p for p in (default, SETTINGS) if p.exists()], encoding="utf-8")
    except configparser.Error as e:
        print(f"  settings.ini could not be read ({e}), using defaults", file=sys.stderr)
        return
    g = lambda sec, key, fb="": cfg.get(sec, key, fallback=fb)
    words = _list(g("gates", "default"))
    if words:
        DEFAULT_GATES = [re.compile(re.escape(w), re.I) for w in words]
    try:
        MIN_PRESENCE = max(0, min(100, float(g("gates", "min_presence", "50"))))
    except ValueError:
        pass
    for key in DEFAULT_VARS:
        DEFAULT_VARS[key] = g("variables", key, "yes" if DEFAULT_VARS[key] else "no"
                              ).strip().lower() in ("yes", "y", "true", "1", "on")
    skip = _list(g("groups", "skip"))
    if skip:
        SKIP_GROUPS = {x.lower() for x in skip}
    names = _list(g("viability", "names"))
    if names:
        LIVE_RE = re.compile(r"^(" + "|".join(map(re.escape, names)) + r")\b", re.I)
    bands = []
    for item in _list(g("viability", "bands")):
        limit, _, colour = item.partition("=")
        val = to_number(limit)
        if val is not None:
            bands.append((val, _hex(colour, "F1E5C6")))
    if bands:
        LIVE_BANDS = sorted(bands)
    GRADIENT = (_hex(g("colours", "gradient_low"), GRADIENT[0]),
                _hex(g("colours", "gradient_high"), GRADIENT[1]))


def open_settings():
    if not SETTINGS.exists():
        load_settings()
    opener = ["open", "-t"] if sys.platform == "darwin" else ["xdg-open"]
    subprocess.run(opener + [str(SETTINGS)])


# ---------------------------------------------------------------- helpers
def to_number(text):
    t = (text or "").strip().replace("%", "")
    if not t:
        return None
    if "," in t and "." in t:          # 1.234,5 or 1,234.5
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
    else:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def well_of(text):
    m = WELL_RE.search((text or "").strip())
    if not m:
        return None
    row, col = m.group(1).upper(), int(m.group(2))
    return (row, col) if 1 <= col <= 24 else None


def iso_dates(name):
    """'02-Nov-2026' -> '20261102', '2026-11-02' -> '20261102' (rest of the name kept)."""
    def dmy(m):
        mon = MONTHS.get(m.group(2)[:3].lower())
        return f"{m.group(3)}{mon:02d}{int(m.group(1)):02d}" if mon else m.group(0)
    name = re.sub(r"\b(\d{1,2})[-_. ]([A-Za-z]{3,9})[-_. ](\d{4})\b", dmy, name)
    name = re.sub(r"\b(\d{4})-(\d{2})-(\d{2})\b", r"\1\2\3", name)
    return name


def gate_matches(path):
    name = path.split("/")[-1]
    return any(rx.search(name) for rx in DEFAULT_GATES)


# ---------------------------------------------------------------- WSP input
def read_wsp(path):
    """-> groups {name: [sampleID]},
          samples {sampleID: (well, {gatepath: {statname: value}}, samplename)},
          gate order, MFI stat names (e.g. 'MFI Median Comp-PE-A')."""
    root = ET.parse(path).getroot()
    samples, gate_order, mfi_names = {}, [], []
    for s in root.iter("Sample"):
        node, ds = s.find("SampleNode"), s.find("DataSet")
        if node is None:
            continue
        sid = node.get("sampleID") or (ds.get("sampleID") if ds is not None else None)
        kw = {k.get("name"): k.get("value") for k in s.iter("Keyword")}
        well = (well_of(kw.get("WELL ID")) or well_of(kw.get("$WELLID"))
                or well_of(node.get("name")) or well_of(kw.get("$FIL")))
        vals = {}

        def walk(n, prefix, parent_count):
            for p in n.findall("Subpopulations/Population"):
                name = (p.get("name") or "").strip()
                gpath = f"{prefix}/{name}" if prefix else name
                try:
                    cnt = float(p.get("count"))
                except (TypeError, ValueError):
                    cnt = None
                v = vals.setdefault(gpath, {})
                if cnt is not None and cnt >= 0:
                    v[COUNT] = cnt
                    if parent_count and parent_count > 0:
                        v[FREQ] = cnt / parent_count * 100
                # statistics added in FlowJo (Median, Mean, ... of a channel)
                for st in p.findall("Statistic") + p.findall("Subpopulations/Statistic"):
                    sname = (st.get("name") or "").strip()
                    if sname.lower() in MFI_STATS:
                        key = f"MFI {sname} {st.get('id') or ''}".strip()
                        val = to_number(st.get("value"))
                        if val is not None:
                            v[key] = val
                            if key not in mfi_names:
                                mfi_names.append(key)
                if gpath not in gate_order:
                    gate_order.append(gpath)
                walk(p, gpath, cnt)

        try:
            total = float(node.get("count"))
        except (TypeError, ValueError):
            total = None
        walk(node, "", total)
        samples[sid] = (well, vals, node.get("name") or "", kw)
    groups = {}
    for g in root.iter("GroupNode"):
        name = (g.get("name") or "").strip()
        groups[name] = [r.get("sampleID") for r in g.iter("SampleRef")]
    real = {n: ids for n, ids in groups.items() if n.lower() not in SKIP_GROUPS and ids}
    if not real:                       # nothing but "All Samples": use it as the one plate
        real = {n: ids for n, ids in groups.items() if n.lower() == "all samples" and ids}
    return real, samples, gate_order, mfi_names


# ---------------------------------------------------------------- dialog
PICKER_JS = r"""
ObjC.import('Cocoa');
var ALL_BOXES = [];      // enabled gate checkboxes, toggled by "Tick all"
var TICK_ALL = null;
var HANDLER_OK = false;
function run(argv) {
  const txt = $.NSString.stringWithContentsOfFileEncodingError(argv[0], $.NSUTF8StringEncoding, null).js;
  const data = JSON.parse(txt);
  const app = $.NSApplication.sharedApplication;
  app.setActivationPolicy($.NSApplicationActivationPolicyAccessory);
  app.activateIgnoringOtherApps(true);
  const alert = $.NSAlert.alloc.init;
  alert.messageText = data.title;
  alert.informativeText = 'Tick what should go into the Excel file.';
  if (data.icon) {
    const img = $.NSImage.alloc.initWithContentsOfFile($(data.icon));
    if (!img.isNil() && img.isValid) alert.icon = img;
  }
  alert.addButtonWithTitle('Make Excel');
  alert.addButtonWithTitle('Cancel');
  alert.addButtonWithTitle('Settings…');

  // "Tick all" needs a real target object; plain checkboxes have no action at all
  let target = null;
  try {
    ObjC.registerSubclass({
      name: 'F2PTickAll',
      methods: {
        'toggle:': {
          types: ['void', ['id']],
          implementation: function (sender) {
            const on = TICK_ALL.state == 1 ? 1 : 0;
            ALL_BOXES.forEach(cb => { cb.state = on; });
          }
        }
      }
    });
    target = $.F2PTickAll.alloc.init;
    HANDLER_OK = true;
  } catch (e) { HANDLER_OK = false; }

  const H = 22, COLW = 90, PAD = 14;
  let longest = 10;
  data.gates.forEach(g => { longest = Math.max(longest, g.label.length); });
  data.groups.forEach(g => { longest = Math.max(longest, g.label.length - 12); });
  const LW = Math.max(300, Math.min(760, Math.round(longest * 7.2) + 20));
  const W = PAD + LW + COLW * data.vars.length;
  const rows = 1 + 1 + data.groups.length + 1 + 1 + data.gates.length;
  const total = rows * H + 3 * 8;
  const view = $.NSView.alloc.initWithFrame($.NSMakeRect(0, 0, W, total));
  let y = total;

  function heading(text) {
    y -= H;
    const head = $.NSTextField.labelWithString(text);
    head.font = $.NSFont.boldSystemFontOfSize(14);
    head.textColor = $.NSColor.systemBlueColor;
    head.frame = $.NSMakeRect(0, y, W, H - 2);
    view.addSubview(head);
  }
  function label(text, x, w, bold, align) {
    const l = $.NSTextField.labelWithString(text);
    if (bold) l.font = $.NSFont.boldSystemFontOfSize(12);
    l.lineBreakMode = $.NSLineBreakByTruncatingMiddle;
    if (align) l.alignment = $.NSTextAlignmentCenter;
    l.frame = $.NSMakeRect(x, y + 2, w, H - 4);
    l.toolTip = text;
    view.addSubview(l);
    return l;
  }
  function checkbox(title, x, w, on, disabled) {
    const cb = $.NSButton.alloc.initWithFrame($.NSMakeRect(x, y, w, H));
    cb.setButtonType(3);                // NSButtonTypeSwitch
    cb.title = title;
    cb.state = on ? 1 : 0;
    if (disabled) { cb.enabled = false; cb.state = 0; }
    view.addSubview(cb);
    return cb;
  }

  // ---- groups
  heading('Groups');
  y -= H;
  const note = label('each group becomes one plate', PAD, W - PAD, false, false);
  note.font = $.NSFont.systemFontOfSize(11);
  note.textColor = $.NSColor.secondaryLabelColor;
  const groupBoxes = data.groups.map(g => { y -= H; return checkbox(g.label, PAD, W - PAD, g.checked, false); });
  y -= 8;

  // ---- gates table: one row per gate, one column per variable
  heading('Gates');
  y -= H;
  TICK_ALL = checkbox('Tick all', PAD, LW - PAD, false, false);
  if (HANDLER_OK) { TICK_ALL.target = target; TICK_ALL.action = 'toggle:'; }
  data.vars.forEach((v, j) => {
    const l = label(v.label, PAD + LW + j * COLW - 10, COLW, true, true);
    if (v.disabled) { l.textColor = $.NSColor.tertiaryLabelColor; l.toolTip = v.tip; }
  });
  const gateBoxes = data.gates.map(g => {
    y -= H;
    const l = label(g.label, PAD, LW - 8, false, false);
    l.toolTip = g.tip;
    return data.vars.map((v, j) => {
      const dis = v.disabled || g.disabled[j];
      const cb = checkbox('', PAD + LW + j * COLW + COLW / 2 - 19, 22, g.checked[j], dis);
      if (!dis) ALL_BOXES.push(cb);
      return cb;
    });
  });

  const visible = Math.min(total, 520);
  const scroll = $.NSScrollView.alloc.initWithFrame($.NSMakeRect(0, 0, W + 20, visible));
  scroll.hasVerticalScroller = total > visible;
  scroll.documentView = view;
  scroll.drawsBackground = false;
  scroll.borderType = $.NSNoBorder;
  view.scrollPoint($.NSMakePoint(0, total));
  alert.accessoryView = scroll;
  alert.window.initialFirstResponder = alert.buttons.objectAtIndex(0);
  alert.layout;
  // tint the translucent window background blue (keeps the frosted look)
  try {
    const frame = alert.window.contentView.superview.isNil() ? alert.window.contentView
                                                             : alert.window.contentView.superview;
    const tint = $.NSView.alloc.initWithFrame(frame.bounds);
    tint.wantsLayer = true;
    tint.layer.backgroundColor =
      $.NSColor.colorWithSRGBRedGreenBlueAlpha(0.30, 0.62, 0.95, 0.22).CGColor;
    tint.autoresizingMask = $.NSViewWidthSizable | $.NSViewHeightSizable;
    frame.addSubviewPositionedRelativeTo(tint, $.NSWindowBelow, $());
  } catch (e) {}
  const res = alert.runModal;
  if (res == 1002) return 'SETTINGS';
  if (res != 1000) return 'CANCEL';
  // without a working handler, a ticked "Tick all" still means everything
  const all = !HANDLER_OK && TICK_ALL.state == 1;
  return JSON.stringify({
    groups: groupBoxes.map(cb => cb.state == 1),
    gates: gateBoxes.map(row => row.map(cb => cb.isEnabled && (all || cb.state == 1)))
  });
}
"""


def pick(title, payload):
    """payload: {groups: [{label, tip, checked}], vars: [{label, tip, disabled}],
                 gates: [{label, tip, checked: [..per var], disabled: [..per var]}]}
    -> {groups: [bool], gates: [[bool per var]]}, or None to cancel.
    Falls back to the defaults when no dialog is available (not on a Mac, --no-dialog)."""
    defaults = {
        "groups": [bool(g["checked"]) for g in payload["groups"]],
        "gates": [[bool(c) and not d and not v.get("disabled")
                   for c, d, v in zip(g["checked"], g["disabled"], payload["vars"])]
                  for g in payload["gates"]],
    }
    if os.environ.get("F2P_NO_DIALOG") or not shutil.which("osascript"):
        return defaults
    icon = Path(__file__).with_name("icon.png")
    data = dict(payload, title=title, icon=str(icon) if icon.exists() else "")
    with tempfile.TemporaryDirectory() as d:
        jf, sf = Path(d, "in.json"), Path(d, "picker.js")
        jf.write_text(json.dumps(data), encoding="utf-8")
        sf.write_text(PICKER_JS, encoding="utf-8")
        r = subprocess.run(["osascript", "-l", "JavaScript", str(sf), str(jf)],
                           capture_output=True, text=True)
    out = r.stdout.strip()
    if r.returncode != 0 or not out:
        print(f"  dialog not available ({r.stderr.strip()[:200]}), using defaults", file=sys.stderr)
        return defaults
    if out == "SETTINGS":
        open_settings()
        raise SystemExit("Cancelled. Settings opened; run f2p again after saving.")
    if out == "CANCEL":
        return None
    return json.loads(out)


def collect_wsp(path):
    groups, samples, gate_order, mfi_names = read_wsp(path)
    if not groups:
        raise SystemExit(f"No groups with samples found in {path.name}")
    used = [s for s in dict.fromkeys(sid for ids in groups.values() for sid in ids) if s in samples]
    n_used = len(used)
    present = {g: sum(1 for s in used if g in samples[s][1]) for g in gate_order}
    gates = [g for g in gate_order if present[g]]
    name = iso_dates(path.stem)
    has_mfi = {g: any(k.startswith("MFI") for s in used for k in samples[s][1].get(g, {}))
               for g in gates}

    def gate_label(g):
        text = " / ".join(g.split("/"))
        if present[g] < n_used:
            text += f"   (only in {present[g]} of {n_used} samples)"
        return text

    var_keys = [FREQ, COUNT, "MFI"]
    defaults = [DEFAULT_VARS["frequency"], DEFAULT_VARS["cells"], DEFAULT_VARS["mfi"]]
    mfi_tip = ("MFI: " + ", ".join(m[4:] for m in mfi_names) if mfi_names
               else "No Median/Mean statistics in this workspace (add them in FlowJo)")
    payload = {
        "groups": [{"label": f"{n}  ({len(ids)} samples)", "tip": n, "checked": True}
                   for n, ids in groups.items()],
        "vars": [{"label": "Frequency", "tip": "Freq. of Parent (%)"},
                 {"label": "#cells", "tip": "number of cells in the gate"},
                 {"label": "MFI", "tip": mfi_tip, "disabled": not mfi_names}],
        "gates": [],
    }
    for g in gates:
        on = gate_matches(g) and present[g] * 100 >= MIN_PRESENCE * n_used
        payload["gates"].append({
            "label": gate_label(g), "tip": g,
            "checked": [on and d for d in defaults],
            "disabled": [False, False, not has_mfi[g]],
        })
    choice = pick(f"f2p: {name}", payload)
    if choice is None:
        raise SystemExit("Cancelled.")
    chosen_groups = [n for n, keep in zip(groups, choice["groups"]) if keep]
    picked = {var: [g for g, row in zip(gates, choice["gates"]) if row[j]]
              for j, var in enumerate(var_keys)}
    if not any(picked.values()):
        raise SystemExit("Nothing ticked in the gate table.")
    if not chosen_groups:
        raise SystemExit("No groups selected.")
    chosen_vars = [v for v in var_keys if picked[v]]

    # (stat key for the sheet, gate path, value key in the sample) - grouped by variable
    columns = []
    for var in chosen_vars:
        for key in (mfi_names if var == "MFI" else [var]):
            for g in picked[var]:
                columns.append((f"{g} | {key}", g, key))
    plates = []
    for gname in chosen_groups:
        data = {}
        for sid in groups[gname]:
            if sid not in samples:
                continue
            well, vals, sname, _ = samples[sid]
            if well is None:
                print(f"  {gname}: no well id for {sname}, skipped", file=sys.stderr)
                continue
            for st, g, key in columns:
                v = vals.get(g, {}).get(key)
                if v is not None:
                    data.setdefault(st, {})[well] = v
        if data:
            plates.append((gname, data))
    if not plates:
        raise SystemExit("The selected groups contain no samples with well ids.")
    stats = [st for st, _, _ in columns if any(st in d for _, d in plates)]
    meta = build_meta(path, [(g, [samples[sid][3] for sid in groups[g] if sid in samples])
                             for g in chosen_groups if any(p == g for p, _ in plates)])
    return name, plates, stats, path.parent, meta


def _uniq(values):
    return list(dict.fromkeys(v.strip() for v in values if v and v.strip()))


def _date_key(d):
    try:
        return _dt.datetime.strptime(d.title(), "%d-%b-%Y")
    except ValueError:
        return _dt.datetime.max


def build_meta(path, group_kws):
    """Metadata for the Info sheet from the FCS keywords stored in the workspace."""
    all_kw = [kw for _, kws in group_kws for kw in kws]
    get = lambda key: _uniq(kw.get(key) for kw in all_kw)
    join = lambda key: ", ".join(get(key))

    acquisition = []
    for g, kws in group_kws:
        dates = sorted(_uniq(k.get("$DATE") for k in kws), key=_date_key)
        starts = sorted(_uniq(k.get("$BTIM") for k in kws))
        ends = sorted(_uniq(k.get("$ETIM") for k in kws))
        acquisition.append([
            g,
            ", ".join(_uniq(k.get("EXPERIMENT NAME") for k in kws)),
            ", ".join(_uniq(k.get("PLATE NAME") for k in kws)),
            " / ".join(iso_dates(d.title()) for d in dates),
            f"{starts[0][:8]} - {ends[-1][:8]}" if starts and ends else "",
            len(kws),
        ])

    when_who = [
        ("Acquisition date", " / ".join(iso_dates(d.title())
                                        for d in sorted(get("$DATE"), key=_date_key))),
        ("Operator", join("$OP")),
        ("Exported by", join("EXPORT USER NAME")),
        ("Export time", " - ".join(dict.fromkeys(
            [min(get("EXPORT TIME")), max(get("EXPORT TIME"))])) if get("EXPORT TIME") else ""),
        ("Workspace", path.name),
        ("Workspace saved", _dt.datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d %H:%M")),
    ]
    instrument = [
        ("Cytometer", join("$CYT")),
        ("Serial number", join("CYTNUM")),
        ("Software", join("CREATOR")),
        ("Configuration", join("CYTOMETER CONFIG NAME")),
        ("Lasers", ", ".join(_uniq(
            all_kw[0].get(f"LASER{i}NAME") for i in range(1, 10)) if all_kw else [])),
        ("Threshold", join("THRESHOLD")),
        ("CST setup", f"{join('CST SETUP STATUS')}  {join('CST SETUP DATE')[:10]}".strip()),
        ("CST beads expired", join("CST BEADS EXPIRED")),
        ("Compensation applied", join("APPLY COMPENSATION")),
    ]
    panel = []
    if all_kw:
        kw0 = all_kw[0]
        n = int(to_number(kw0.get("$PAR")) or 0)
        for i in range(1, n + 1):
            ch = (kw0.get(f"$P{i}N") or "").strip()
            if not ch or ch.startswith("Comp-"):
                continue
            markers = _uniq(k.get(f"$P{i}S") for k in all_kw)
            volt = kw0.get(f"$P{i}V") or ""
            panel.append((ch, ", ".join(markers), volt))
    experiments = get("EXPERIMENT NAME")
    return {"experiment": "; ".join(experiments), "when_who": when_who,
            "instrument": instrument, "panel": panel, "acquisition": acquisition}


# ---------------------------------------------------------------- naming
def short_names(stats):
    """'A/B/CD107a+ | Freq. of Parent (%)' -> grid label 'CD107a+', list header 'CD107a+ (%)'.
    Gate names that occur more than once get their parent gate: 'CD8/IFNg+'."""
    def split(s):
        path, _, stat = s.partition("|")
        return [p.strip() for p in path.strip().strip("/").split("/")], stat.strip()

    parsed = {s: split(s) for s in stats}
    out = {}
    for s, (parts, stat) in parsed.items():
        depth = 1
        while depth < len(parts):
            label = "/".join(parts[-depth:])
            clash = [o for o, (op, _) in parsed.items()
                     if op != parts and "/".join(op[-depth:]) == label]
            if not clash:
                break
            depth += 1
        gate = "/".join(parts[-depth:])
        if stat == COUNT:
            label = header = f"{gate} #cells"
        elif stat.startswith("MFI"):
            label = header = f"{gate} {stat.replace('Comp-', '')}"
        elif "%" in stat or not stat:
            label = header = f"{gate} (%)" if stat else gate
        else:
            label = header = f"{gate} {stat}"
        out[s] = (label, header)
    return out


def is_live_freq(stat):
    path, _, kind = stat.partition("|")
    gate = path.strip().strip("/").split("/")[-1].strip()
    return bool(LIVE_RE.match(gate)) and "%" in kind


def colour_range(ws, rng, top_left, live):
    if live:   # viability thresholds, muted earthy tones; first match wins
        for limit, colour in LIVE_BANDS:
            fill = PatternFill("solid", start_color=colour, end_color=colour)
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f"AND(ISNUMBER({top_left}),{top_left}<{limit})"],
                fill=fill, stopIfTrue=True))
    else:
        ws.conditional_formatting.add(rng, ColorScaleRule(
            start_type="min", start_color=GRADIENT[0], end_type="max", end_color=GRADIENT[1]))


def row_index(letter):
    return ord(letter) - ord("A")


# ---------------------------------------------------------------- writing
def write_info(wb, meta):
    wi = wb.create_sheet("Info")
    wi["A1"], wi["A1"].font = meta["experiment"] or "Info", TITLE
    r = 3

    def section(name, header, rows):
        nonlocal r
        wi.cell(r, 2, name).font = Font(bold=True, size=12)
        r += 1
        if header:
            for j, h in enumerate(header):
                c = wi.cell(r, 2 + j, h)
                c.font, c.fill = BOLD, HEAD_FILL
            r += 1
        for row in rows:
            for j, v in enumerate(row):
                c = wi.cell(r, 2 + j, v)
                if j == 0 and not header:
                    c.font = BOLD
            r += 1
        r += 1

    section("When and who", None, meta["when_who"])
    section("Plates", ["Group", "Experiment", "Plate", "Acquired", "Time", "Wells"],
            meta["acquisition"])
    section("Instrument", None, meta["instrument"])
    section("Panel", ["Channel", "Marker", "Voltage"], meta["panel"])
    wi.column_dimensions["A"].width = 3
    for col, w in zip("BCDEFG", (24, 34, 22, 12, 20, 8)):
        wi.column_dimensions[col].width = w


def write_workbook(title, plates, all_stats, out_dir, meta=None):
    file_title = f"{title}_f2p"            # fallback: .wsp name
    if meta and meta.get("experiment"):
        title = meta["experiment"]
        file_title = meta["experiment"].replace("; ", " + ")   # name the file after the experiment
    wells = {w for _, d in plates for vals in d.values() for w in vals}
    big = any(row_index(r) > 7 or c > 12 for r, c in wells)
    n_rows, n_cols = (16, 24) if big else (8, 12)
    letters = [chr(ord("A") + i) for i in range(n_rows)]
    names = short_names(all_stats)
    wb = Workbook()

    # ---- Plates sheet: gates left->right, plates top->bottom
    ws = wb.active
    ws.title = "Plates"
    ws["A1"], ws["A1"].font = title, TITLE
    grid_w = n_cols + 1                       # label column + well columns
    band_h = n_rows + 4                       # plate name + header + rows + gate line + gap
    for p, (plate, data) in enumerate(plates):
        top = 3 + p * band_h
        if plate:
            ws.cell(top, 3, plate).font = Font(bold=True, size=12)
        for s_i, stat in enumerate(all_stats):
            left = 3 + s_i * (grid_w + 1)
            hdr = top + 1
            c = ws.cell(hdr, left, names[stat][0])
            c.font, c.fill, c.border = BOLD, HEAD_FILL, BOX
            for j in range(n_cols):
                c = ws.cell(hdr, left + 1 + j, f"{j + 1:02d}")
                c.font, c.fill, c.border, c.alignment = BOLD, HEAD_FILL, BOX, CENTER
            vals = data.get(stat, {})
            fmt = "0" if "#cells" in stat else "0.00"
            for i, letter in enumerate(letters):
                c = ws.cell(hdr + 1 + i, left, letter)
                c.font, c.fill, c.border, c.alignment = BOLD, HEAD_FILL, BOX, CENTER
                for j in range(n_cols):
                    c = ws.cell(hdr + 1 + i, left + 1 + j, vals.get((letter, j + 1)))
                    c.border = BOX
                    c.number_format = fmt
            rng = (f"{get_column_letter(left + 1)}{hdr + 1}:"
                   f"{get_column_letter(left + n_cols)}{hdr + n_rows}")
            colour_range(ws, rng, f"{get_column_letter(left + 1)}{hdr + 1}", is_live_freq(stat))
            # full gating path under the grid (below the "H" row label)
            path = " / ".join(x.strip() for x in stat.partition("|")[0].strip().strip("/").split("/"))
            ws.cell(hdr + n_rows + 1, left, f"Gate: {path}").font = GATE_FONT
    for s_i, stat in enumerate(all_stats):
        left = 3 + s_i * (grid_w + 1)
        ws.column_dimensions[get_column_letter(left)].width = max(8, min(22, len(names[stat][0]) + 2))
        for j in range(n_cols):
            ws.column_dimensions[get_column_letter(left + 1 + j)].width = 6.5
        ws.column_dimensions[get_column_letter(left + grid_w)].width = 3
    ws.column_dimensions["A"].width = ws.column_dimensions["B"].width = 3

    # ---- List sheet: wells down, plates across, grouped by gate
    wl = wb.create_sheet("List")
    wl["A1"], wl["A1"].font = title, TITLE
    order = sorted(wells, key=lambda w: (row_index(w[0]), w[1]))
    for k, (r, c) in enumerate(order):
        wl.cell(4 + k, 2, f"{r}{c:02d}").font = BOLD
    wl.cell(3, 2, "Well").font = BOLD
    col = 3
    for stat in all_stats:
        start = col
        for plate, data in plates:
            h = wl.cell(3, col, plate)
            h.font, h.fill, h.alignment = BOLD, HEAD_FILL, CENTER
            vals = data.get(stat, {})
            fmt = "0" if "#cells" in stat else "0.00"
            for k, w in enumerate(order):
                wl.cell(4 + k, col, vals.get(w)).number_format = fmt
            wl.column_dimensions[get_column_letter(col)].width = max(
                12, min(28, max(len(plate), len(names[stat][1])) + 2))
            col += 1
        wl.cell(2, start, names[stat][1]).font = BOLD
        if col - 1 > start:
            wl.merge_cells(start_row=2, start_column=start, end_row=2, end_column=col - 1)
        wl.cell(2, start).alignment = CENTER
        if order:
            rng = f"{get_column_letter(start)}4:{get_column_letter(col - 1)}{3 + len(order)}"
            colour_range(wl, rng, f"{get_column_letter(start)}4", is_live_freq(stat))
    wl.column_dimensions["A"].width = 3
    wl.freeze_panes = "C4"

    if meta:
        write_info(wb, meta)
    safe = re.sub(r'[\\/:*?"<>|]+', "-", file_title).strip()
    out = out_dir / f"{safe}.xlsx"
    wb.save(out)
    return out


def convert(path):
    path = Path(path).expanduser().resolve()
    if path.is_dir():
        wsps = sorted(p for p in path.glob("*.wsp") if not p.name.startswith((".", "~")))
        if not wsps:
            raise SystemExit(f"No FlowJo workspace (.wsp) found in {path}")
        jobs = [collect_wsp(p) for p in wsps]
    elif path.suffix.lower() == ".wsp":
        jobs = [collect_wsp(path)]
    else:
        raise SystemExit(f"{path.name} is not a FlowJo workspace (.wsp)")
    results = []
    for title, plates, stats, out_dir, meta in jobs:
        out = write_workbook(title, plates, stats, out_dir, meta)
        results.append((out, len(plates), len(stats)))
    return results


def main(argv):
    load_settings()
    if "--settings" in argv:
        open_settings()
        return
    if "--no-dialog" in argv:
        os.environ["F2P_NO_DIALOG"] = "1"
        argv = [a for a in argv if a != "--no-dialog"]
    for p in argv or ["."]:
        for out, n_p, n_s in convert(p):
            print(f"{out}  ({n_p} plates, {n_s} gates)")


if __name__ == "__main__":
    main(sys.argv[1:])
