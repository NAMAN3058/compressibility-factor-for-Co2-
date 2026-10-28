#!/usr/bin/env python3
"""
Peace Software CO2 Z-factor dashboard.

Run:   python app.py        (opens http://127.0.0.1:5057 in your browser)

What it does
  1. You upload an Excel workbook (one or many sheets, e.g. NaCl, MgCl2, CaCl2).
  2. Per sheet it finds the pressure column (header like "psi", "Pressure", "Pi", "100psi")
     and its unit (psi / bar / kPa / MPa) from the header text.
  3. P1 = first pressure of the sheet (repeated down the sheet), P2 = each row's pressure.
  4. Z1 / Z2 are read from the Peace Software CO2 calculator website (NOT CoolProp).
  5. You watch the values fill in live, then download <name>_Z.xlsx with P1, P2, Z1, Z2
     added next to your original data (original cells untouched).

Speed: every distinct pressure is sent to the website only once (cache), requests reuse
one HTTP connection per worker thread, and a few workers run in parallel.
"""
import io
import os
import re
import sys
import json
import time
import uuid
import threading
import traceback
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import urljoin

import openpyxl
import requests
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify, send_file, Response

# --------------------------------------------------------------------------- config
SITE_URLS = [u for u in [
    os.environ.get("PEACE_URL"),
    "https://www.peacesoftware.de/einigewerte/co2_e.html",
    "http://www.peacesoftware.de/einigewerte/co2_e.html",
    "https://peace-software.de/einigewerte/co2_e.html",
    "http://peace-software.de/einigewerte/co2_e.html",
] if u]
PORT = int(os.environ.get("PORT", "5057"))
OUT_DIR = Path(__file__).parent / "outputs"
OUT_DIR.mkdir(exist_ok=True)
DEBUG_HTML = Path(__file__).parent / "debug_last_page.html"
UA = {"User-Agent": "Mozilla/5.0 (compatible; ZDashboard/1.0)"}

# factors to bar, used only if the website has no option for the unit you need
TO_BAR = {"psi": 0.0689475729, "bar": 1.0, "kpa": 0.01, "mpa": 10.0, "atm": 1.01325, "pa": 1e-5}
UNIT_LABEL = {"psi": "psi", "bar": "bar", "kpa": "kPa", "mpa": "MPa", "atm": "atm", "pa": "Pa"}
ALIASES = {"psi": ["psi", "psia", "psig"], "bar": ["bar", "bara"], "kpa": ["kpa"], "mpa": ["mpa"],
           "atm": ["atm"], "pa": ["pa"]}


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def unit_from_text(s):
    s = str(s or "").lower()
    for key, pat in (("kpa", r"kpa"), ("mpa", r"mpa"), ("psi", r"psi"), ("bar", r"bar"),
                     ("atm", r"\batm\b|atm$"), ("pa", r"(?<![a-z])pa(?![a-z])")):
        if re.search(pat, s):
            return key
    return None


# --------------------------------------------------------------------------- website client
NUM = r"([-+]?(?:\d+[.,]?\d*|[.,]\d+)(?:[eE][-+]?\d+)?)"


def to_float(s):
    s = str(s).strip().replace(" ", "")
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    return float(s)


def _prev_text(el, n=1):
    parts = []
    for s in el.find_all_previous(string=True, limit=15):
        t = s.strip()
        if not t or s.parent.name in ("script", "style", "option", "title"):
            continue
        parts.append(t)
        if len(parts) >= n:
            break
    return " ".join(reversed(parts)).lower()


def parse_z(html, extra_values=None):
    """Find the compressibility factor Z in a result page. Returns float or None."""
    soup = BeautifulSoup(html, "html.parser")
    for t in soup(["script", "style"]):
        t.decompose()
    cands = []
    for el in soup.find_all(["input", "textarea"]):
        ctx = ((el.get("name") or "") + " " + (el.get("id") or "") + " " + _prev_text(el)).lower()
        cands.append((ctx, el.get("value") or el.get_text()))
    cands += list(extra_values or [])
    for ctx, val in cands:
        ctx = (ctx or "").lower()
        if re.search(r"compress|z[-_ ]?fact|zfactor|(^|[^a-z])z($|[^a-z])", ctx or ""):
            try:
                v = to_float(val)
                if 0 < v < 10:
                    return v
            except Exception:
                pass
    text = soup.get_text(" ", strip=True)
    for pat in (r"compressib[a-z]*(?:\s+factor)?(?:\s*\(?\s*z\s*\)?)?[^0-9+\-.,]{0,20}" + NUM,
                r"(?<![A-Za-z])Z\s*[=:]\s*" + NUM):
        for m in re.finditer(pat, text, re.I):
            try:
                v = to_float(m.group(1))
                if 0 < v < 10:
                    return v
            except Exception:
                continue
    return None


NON_TEXT_TYPES = {"hidden", "submit", "button", "reset", "image", "file", "checkbox", "radio",
                  "password", "color", "range"}
T_LABELS = {"c", "k", "f", "degc", "degk", "degf", "kelvin", "celsius", "fahrenheit"}


def fmtnum(v):
    return format(float(v), ".10g")


def elements_summary(html, limit=30):
    """Compact list of the page's input/select/button elements (for error messages)."""
    soup = BeautifulSoup(html, "html.parser")
    out = []
    for el in soup.find_all(["input", "select", "button", "textarea"]):
        d = [el.name]
        for a in ("type", "name", "id", "value", "onclick", "onchange"):
            v = el.get(a)
            if v:
                d.append(f"{a}={str(v)[:30]}")
        out.append(" ".join(d))
        if len(out) >= limit:
            break
    return "; ".join(out) or "(no input/select/button elements at all)"


class FormSpec:
    """What we learned about the calculator (fields, unit controls, method...)."""

    def __init__(self):
        self.action = None
        self.method = "get"
        self.base = {}
        self.t_name = self.p_name = None      # keys (name / id / position)
        self.t_css = self.p_css = None        # selectors for the browser
        self.t_sel = self.p_sel = None        # (key, [(label, value)], css, kind)
        self.t_fixed = None                   # 'c'/'k' when the unit is only printed as text
        self.p_fixed = None
        self.submit = None
        self.guessed = False
        self.formless = False
        self.page_url = None

    def describe(self):
        return {"action": self.action, "method": self.method, "has_html_form": not self.formless,
                "temperature_field": self.t_name, "pressure_field": self.p_name,
                "temperature_units": [o[0] for o in self.t_sel[1]] if self.t_sel else f"fixed: {self.t_fixed}",
                "pressure_units": [o[0] for o in self.p_sel[1]] if self.p_sel else f"fixed: {self.p_fixed}",
                "fields_guessed_by_order": self.guessed}


def discover(html, page_url):
    soup = BeautifulSoup(html, "html.parser")
    ipos = {id(el): k for k, el in enumerate(soup.find_all("input"))}
    spos = {id(el): k for k, el in enumerate(soup.find_all("select"))}

    def ikey(el):
        return el.get("name") or el.get("id") or f"input#{ipos[id(el)]}"

    def icss(el):
        if el.get("name"):
            return f"[name='{el['name']}']"
        if el.get("id"):
            return f"#{el['id']}"
        return f"xpath=(//input)[{ipos[id(el)] + 1}]"

    def near(el):
        nxt = el.find_next(string=lambda s: s and s.strip() and s.parent.name not in ("script", "style"))
        return (_prev_text(el, 2) + " " + (nxt or "")).lower()

    containers = [(f, False) for f in soup.find_all("form")] + [(soup, True)]
    for cont, formless in containers:
        texts = [i for i in cont.find_all("input")
                 if (i.get("type") or "text").lower() not in NON_TEXT_TYPES and (formless or i.get("name"))]
        if len(texts) < 2:
            continue
        spec = FormSpec()
        spec.page_url = page_url
        spec.formless = formless
        spec.action = urljoin(page_url, (cont.get("action") or "") if not formless else "")
        spec.method = "get" if formless else (cont.get("method") or "get").lower()
        if not formless:
            for el in cont.find_all(["input", "select", "textarea"]):
                name = el.get("name")
                if not name:
                    continue
                typ = (el.get("type") or "text").lower()
                if el.name == "select":
                    opts = el.find_all("option")
                    sel = next((o for o in opts if o.has_attr("selected")), opts[0] if opts else None)
                    spec.base[name] = (sel.get("value", sel.get_text(strip=True)) if sel else "")
                elif typ in ("checkbox", "radio"):
                    if el.has_attr("checked"):
                        spec.base[name] = el.get("value", "on")
                elif typ in ("submit", "button", "reset", "image", "file"):
                    if typ == "submit" and spec.submit is None:
                        spec.submit = (name, el.get("value", ""))
                else:
                    spec.base[name] = el.get("value", "")
        t = p = None
        for use_label in (False, True):
            for i in texts:
                ctx = (ikey(i) + " " + (i.get("id") or "")).lower()
                if use_label:
                    ctx += " " + _prev_text(i).lower()
                if t is None and i is not p and re.search(r"temp|(^|[^a-z])t($|[^a-z])", ctx):
                    t = i
                elif p is None and i is not t and re.search(r"press|druck|(^|[^a-z])p($|[^a-z])", ctx):
                    p = i
        if t is None or p is None or t is p:
            t, p = texts[0], texts[1]
            spec.guessed = True
        spec.t_name, spec.p_name = ikey(t), ikey(p)
        spec.t_css, spec.p_css = icss(t), icss(p)

        def assign(group, last):
            labels = [norm(o[0]) for o in group[1]]
            is_t = any(x in T_LABELS for x in labels)
            is_p = any(unit_from_text(o[0]) for o in group[1])
            if is_p and not is_t and spec.p_sel is None:
                spec.p_sel = group
            elif is_t and spec.t_sel is None:
                spec.t_sel = group
            elif last == "p" and spec.p_sel is None and len(group[1]) > 1:
                spec.p_sel = group
            elif last == "t" and spec.t_sel is None and len(group[1]) > 1:
                spec.t_sel = group

        last = None
        radios = {}
        for el in cont.find_all(["input", "select"]):
            if el.name == "input":
                typ = (el.get("type") or "text").lower()
                if el is t:
                    last = "t"
                elif el is p:
                    last = "p"
                elif typ == "radio" and el.get("name"):
                    lab = el.find_next(string=lambda s: s and s.strip())
                    radios.setdefault(el["name"], []).append(
                        ((lab or el.get("value") or "").strip(), el.get("value") or (lab or "").strip()))
            else:
                key = el.get("name") or el.get("id") or f"select#{spos[id(el)]}"
                css = (f"[name='{el['name']}']" if el.get("name") else f"#{el['id']}" if el.get("id")
                       else f"xpath=(//select)[{spos[id(el)] + 1}]")
                opts = [(o.get_text(strip=True), o.get("value", o.get_text(strip=True)))
                        for o in el.find_all("option")]
                assign((key, opts, css, "select"), last)
        for name, opts in radios.items():
            assign((name, opts, f"input[name='{name}']", "radio"), None)
        if spec.p_sel is None:
            spec.p_fixed = unit_from_text(near(p))
        if spec.t_sel is None:
            spec.t_fixed = "k" if re.search(r"(^|[^a-z])k($|[^a-z])|kelvin", near(t)) else "c"
        return spec
    return None


def pick_option(sel, key):
    """Find the option value of a unit control for a unit key."""
    if not sel:
        return None
    for label, value in sel[1]:
        for n in (norm(label), norm(value)):
            if key == "c" and n in ("c", "degc", "celsius"):
                return value
            if key == "k" and n in ("k", "kelvin", "degk"):
                return value
            if key not in ("c", "k") and n in ALIASES[key]:
                return value
    return None


class SiteError(Exception):
    pass


class ParseFail(SiteError):
    pass


def plan_units(sp, p, p_unit, t, t_unit):
    """Return (p, pressure_option, t, temperature_option). Converts only if the site lacks the unit."""
    popt = topt = None
    if sp.p_sel:
        popt = pick_option(sp.p_sel, p_unit)
        if popt is None:
            popt = pick_option(sp.p_sel, "bar")
            if popt is None:
                raise SiteError(f"Website has no option for {UNIT_LABEL[p_unit]} or bar")
            p = p * TO_BAR[p_unit]
    else:
        if sp.p_fixed is None:
            raise SiteError("Could not find the pressure-unit control on the page - refusing to guess the unit")
        p = p * TO_BAR[p_unit] / TO_BAR[sp.p_fixed]
    if sp.t_sel:
        topt = pick_option(sp.t_sel, t_unit)
        if topt is None:
            other = "k" if t_unit == "c" else "c"
            topt = pick_option(sp.t_sel, other)
            if topt is None:
                raise SiteError("Website has no °C or K option")
            t = t + 273.15 if other == "k" else t - 273.15
    elif (sp.t_fixed or "c") != t_unit:
        t = t + 273.15 if t_unit == "c" else t - 273.15
    return p, popt, t, topt


class RequestsBackend:
    name = "requests"

    def __init__(self, log):
        self.log = log
        self.spec = None
        self.local = threading.local()
        diag = []
        for url in SITE_URLS:
            try:
                r = requests.get(url, headers=UA, timeout=20)
                html = r.content.decode(r.encoding or r.apparent_encoding or "utf-8", "replace")
                DEBUG_HTML.write_text(html, encoding="utf-8", errors="replace")
                soup = BeautifulSoup(html, "html.parser")
                title = soup.title.get_text(strip=True) if soup.title else ""
                diag.append(f"{url} -> {r.url}, status {r.status_code}, title '{title}', "
                            f"forms {len(soup.find_all('form'))}, inputs {len(soup.find_all('input'))}, "
                            f"iframes {len(soup.find_all(['iframe', 'frame']))}, "
                            f"elements: {elements_summary(html, 12)}")
                r.raise_for_status()
                spec = discover(html, r.url)
                if spec and not spec.formless:
                    self.spec = spec
                    self.log(f"Website reached: {r.url}")
                    return
                if spec:
                    diag.append("calculator inputs found but there is no HTML form (JavaScript page)")
            except Exception as e:  # noqa
                diag.append(f"{url} -> error: {e}")
        raise SiteError("No HTML calculator form. " + " || ".join(diag))

    def _session(self):
        s = getattr(self.local, "s", None)
        if s is None:
            s = requests.Session()
            s.headers.update(UA)
            self.local.s = s
        return s

    def z(self, p, p_unit, t, t_unit):
        sp = self.spec
        p, popt, t, topt = plan_units(sp, p, p_unit, t, t_unit)
        payload = dict(sp.base)
        if popt is not None:
            payload[sp.p_sel[0]] = popt
        if topt is not None:
            payload[sp.t_sel[0]] = topt
        payload[sp.p_name] = fmtnum(p)
        payload[sp.t_name] = fmtnum(t)
        if sp.submit and sp.submit[0]:
            payload[sp.submit[0]] = sp.submit[1]
        for attempt in range(3):
            try:
                kw = {"params": payload} if sp.method == "get" else {"data": payload}
                r = self._session().request(sp.method, sp.action, timeout=25, **kw)
                r.raise_for_status()
                html = r.content.decode(r.encoding or r.apparent_encoding or "utf-8", "replace")
                z = parse_z(html)
                if z is not None:
                    return z
                DEBUG_HTML.write_text(html, encoding="utf-8", errors="replace")
                txt = BeautifulSoup(html, "html.parser").get_text(" ", strip=True)[:400]
                raise ParseFail(f"Website answered but no Z found. Sent {sp.method.upper()} {sp.action} "
                                f"{ {k: v for k, v in payload.items()} }. Page text: {txt!r}")
            except ParseFail:
                raise
            except Exception as e:  # noqa
                if attempt == 2:
                    raise SiteError(str(e))
                time.sleep(0.6 * 2 ** attempt)


class BrowserBackend:
    """Drives the real calculator page in a headless browser (for JavaScript calculators).
    Needs:  pip install playwright   and   playwright install chromium"""
    name = "browser"
    single_thread = True

    def __init__(self, log):
        from playwright.sync_api import sync_playwright
        self.log = log
        self.spec = None
        self.pw = sync_playwright().start()
        self.browser = self.pw.chromium.launch(headless=True)
        self.page = self.browser.new_page()
        diag = []
        for url in SITE_URLS:
            try:
                self.page.goto(url, timeout=30000, wait_until="domcontentloaded")
                self.page.wait_for_timeout(500)
                html = self.page.content()
                DEBUG_HTML.write_text(html, encoding="utf-8", errors="replace")
                spec = discover(html, self.page.url)
                if spec:
                    self.spec = spec
                    log(f"Calculator page opened: {self.page.url}")
                    return
                diag.append(f"{url}: loaded, but no calculator inputs. Elements: {elements_summary(html)}")
            except Exception as e:  # noqa
                diag.append(f"{url}: {str(e)[:100]}")
        self.close()
        raise SiteError("Browser could not find the calculator: " + " | ".join(diag))

    def _read_z(self):
        pg = self.page
        vals = pg.evaluate("""()=>[...document.querySelectorAll('input,textarea')].map(e=>[
            (e.name||'')+' '+(e.id||'')+' '+(e.closest('tr')?e.closest('tr').innerText:''), e.value])""")
        body = pg.inner_text("body").replace("<", " ").replace("&", " ")
        return parse_z("<html><body>" + body + "</body></html>", vals)

    def _set_unit(self, sel, opt):
        pg = self.page
        if sel[3] == "radio":
            pg.check(f"{sel[2]}[value='{opt}']")
        else:
            try:
                pg.select_option(sel[2], value=opt)
            except Exception:
                pg.select_option(sel[2], label=opt)

    def z(self, p, p_unit, t, t_unit):
        sp, pg = self.spec, self.page
        p, popt, t, topt = plan_units(sp, p, p_unit, t, t_unit)
        if popt is not None:
            self._set_unit(sp.p_sel, popt)
        if topt is not None:
            self._set_unit(sp.t_sel, topt)
        pg.fill(sp.t_css, fmtnum(t))
        pg.fill(sp.p_css, fmtnum(p))
        prev = self._read_z()
        btns = pg.query_selector_all("input[type=submit], input[type=button], button")
        pick = next((b for b in btns if re.search(
            r"calc|berech|comput|submit|run|start|ok|go",
            ((b.get_attribute("value") or "") + " " + (b.inner_text() or "")).lower())), btns[0] if btns else None)
        if pick:
            pick.click()
        else:
            pg.press(sp.p_css, "Enter")
        z = None
        for _ in range(14):
            pg.wait_for_timeout(150)
            z = self._read_z()
            if z is not None and z != prev:
                return z
        if z is not None:
            return z
        raise SiteError("Z not found on the page after calculating")

    def close(self):
        try:
            self.browser.close()
            self.pw.stop()
        except Exception:
            pass


def make_backend(log):
    err1 = None
    try:
        be = RequestsBackend(log)
        be.z(10, "bar", 30, "c")   # smoke test: does the plain-HTTP path really return Z?
        return be
    except SiteError as e:
        err1 = str(e)
    log("Plain HTTP mode not possible; switching to a real browser (Playwright)...")
    try:
        be = BrowserBackend(log)
    except ImportError:
        raise SiteError("Needs a browser engine: pip install playwright && playwright install chromium. "
                        "(Plain mode said: " + err1[:600] + ")")
    try:
        be.z(10, "bar", 30, "c")
    except Exception as e:  # noqa
        html = be.page.content()
        be.close()
        raise SiteError(f"Browser opened the page but could not read Z ({e}). Spec: {be.spec.describe()}. "
                        f"Elements: {elements_summary(html)}")
    return be


# --------------------------------------------------------------------------- Excel analysis
P_KEY = re.compile(r"press|psi|bar|kpa|mpa|^pi$|^p$|^p\d*$|pressure|\bp[ _]?(in|i)\b|^\d+(\.\d+)?\s*(psi|bar|kpa|mpa)", re.I)


def analyze_sheet(ws_val, default_unit):
    """Return dict(header_row, p_col, t_col, unit, rows[(row, P)]) or None."""
    max_r, max_c = ws_val.max_row, ws_val.max_column
    header_row = None
    for r in range(1, min(max_r, 15) + 1):
        for c in range(1, max_c + 1):
            v = ws_val.cell(r, c).value
            if isinstance(v, str) and P_KEY.search(v.strip()) and "time" not in v.lower():
                header_row = r
                break
        if header_row:
            break
    if header_row is None:
        header_row = 1
    heads = {c: ws_val.cell(header_row, c).value for c in range(1, max_c + 1)}

    def numeric_count(c):
        n = 0
        for r in range(header_row + 1, max_r + 1):
            v = ws_val.cell(r, c).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                n += 1
        return n

    p_col = None
    for c, h in heads.items():
        if isinstance(h, str) and P_KEY.search(h.strip()) and "time" not in h.lower() and numeric_count(c):
            p_col = c
            break
    if p_col is None:  # first numeric column that is not time
        for c, h in heads.items():
            if not (isinstance(h, str) and "time" in h.lower()) and numeric_count(c):
                p_col = c
                break
    if p_col is None:
        return None
    t_col = next((c for c, h in heads.items() if isinstance(h, str) and "time" in h.lower()), None)
    unit = unit_from_text(heads[p_col]) or unit_from_text(ws_val.title) or default_unit
    rows = []
    for r in range(header_row + 1, max_r + 1):
        v = ws_val.cell(r, p_col).value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            rows.append((r, float(v)))
    if not rows:
        return None
    return {"header_row": header_row, "p_col": p_col, "t_col": t_col, "unit": unit, "rows": rows,
            "header": str(heads[p_col]), "last_col": max(c for c, h in heads.items() if h not in (None, ""))}


# --------------------------------------------------------------------------- jobs
JOBS = {}


class Job:
    def __init__(self):
        self.id = uuid.uuid4().hex[:10]
        self.events = []
        self.lock = threading.Lock()
        self.cancel = False
        self.done = False
        self.error = None
        self.out_path = None
        self.total = 0
        self.finished = 0
        self.failed = 0
        self.t0 = time.time()

    def emit(self, **ev):
        with self.lock:
            self.events.append(ev)

    def log(self, msg):
        self.emit(t="log", msg=msg)


def run_job(job, data, filename, temp, temp_unit, default_unit, workers):
    backend = None
    try:
        wb_val = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        wb_out = openpyxl.load_workbook(io.BytesIO(data))
        plans = {}
        for ws in wb_val.worksheets:
            info = analyze_sheet(ws, default_unit)
            if info is None:
                job.log(f"Sheet '{ws.title}': no pressure column found - skipped")
                continue
            plans[ws.title] = info
            job.log(f"Sheet '{ws.title}': pressure column '{info['header']}' unit {UNIT_LABEL[info['unit']]}, "
                    f"{len(info['rows'])} rows")
        if not plans:
            raise SiteError("No sheet with a numeric pressure column was found.")
        job.total = sum(len(p["rows"]) + 1 for p in plans.values())
        meta = []
        for name, info in plans.items():
            ws = wb_val[name]
            meta.append({"sheet": name, "unit": UNIT_LABEL[info["unit"]], "p1": info["rows"][0][1],
                         "rows": [[r, (ws.cell(r, info["t_col"]).value if info["t_col"] else None), p]
                                  for r, p in info["rows"]]})
        job.emit(t="init", sheets=meta, temp=temp, temp_unit=temp_unit.upper())

        job.log("Connecting to the Peace Software website...")
        backend = make_backend(job.log)
        job.log(f"Connected ({backend.name} mode). Calculating...")
        if getattr(backend, "single_thread", False):
            workers = 1
        cache = {}
        clock = threading.Lock()

        def get_z(p, unit):
            key = (unit, round(p, 10))
            with clock:
                if key in cache:
                    return cache[key]
            z = backend.z(p, unit, temp, temp_unit)
            with clock:
                cache[key] = z
            return z

        for name, info in plans.items():
            if job.cancel:
                break
            unit, rows = info["unit"], info["rows"]
            p1 = rows[0][1]
            try:
                z1 = get_z(p1, unit)
            except SiteError as e:
                raise SiteError(f"Z1 failed for sheet '{name}': {e}")
            job.finished += 1
            job.emit(t="z1", sheet=name, z=z1)
            zs = {}
            uniq = list(dict.fromkeys(p for _, p in rows))

            def work(p):
                if job.cancel:
                    return p, None, "cancelled"
                try:
                    return p, get_z(p, unit), None
                except Exception as e:  # noqa
                    return p, None, str(e)

            rows_by_p = {}
            for r, p in rows:
                rows_by_p.setdefault(p, []).append(r)
            def record(p, z, err):
                for r in rows_by_p[p]:
                    job.finished += 1
                    if z is None:
                        job.failed += 1
                    job.emit(t="z2", sheet=name, row=r, z=z, err=err)
                zs[p] = z

            if getattr(backend, "single_thread", False):
                for p in uniq:      # the browser may only be used from this thread
                    record(*work(p))
            else:
                with ThreadPoolExecutor(max_workers=workers) as ex:
                    for res in ex.map(work, uniq):
                        record(*res)
            # write into output workbook (original cells untouched)
            ws = wb_out[name]
            hr, col = info["header_row"], info["last_col"] + 1
            for k, h in enumerate(("P1", "P2", "Z1", "Z2")):
                cell = ws.cell(hr, col + k, h)
                cell.font = openpyxl.styles.Font(bold=True)
            for r, p in rows:
                ws.cell(r, col, p1)
                ws.cell(r, col + 1, p)
                ws.cell(r, col + 2, z1)
                c = ws.cell(r, col + 3, zs.get(p))
                for k in (2, 3):
                    ws.cell(r, col + k).number_format = "0.00000"
            for k in range(4):
                ws.column_dimensions[openpyxl.utils.get_column_letter(col + k)].width = 12
        info_ws = wb_out.create_sheet("Z_settings")
        info_ws.append(["Source", "Peace Software online CO2 calculator"])
        info_ws.append(["Temperature", f"{temp} {temp_unit.upper()}"])
        info_ws.append(["Generated", time.strftime("%Y-%m-%d %H:%M:%S")])
        for n, i in plans.items():
            info_ws.append([f"Sheet '{n}' pressure unit", UNIT_LABEL[i["unit"]]])
        out = OUT_DIR / (Path(filename).stem + "_Z.xlsx")
        wb_out.save(out)
        job.out_path = out
        job.emit(t="done", file=out.name, failed=job.failed, cancelled=job.cancel)
    except Exception as e:  # noqa
        job.error = str(e)
        job.emit(t="error", msg=str(e))
        traceback.print_exc()
    finally:
        job.done = True
        if backend is not None and hasattr(backend, "close"):
            backend.close()


# --------------------------------------------------------------------------- web app
app = Flask(__name__)


@app.get("/")
def index():
    return Response(PAGE, mimetype="text/html")


@app.post("/start")
def start():
    f = request.files.get("file")
    if not f:
        return jsonify(error="No file uploaded"), 400
    try:
        temp = float(request.form.get("temp", "30"))
    except ValueError:
        return jsonify(error="Temperature must be a number"), 400
    temp_unit = request.form.get("temp_unit", "c").lower()
    default_unit = request.form.get("unit", "psi").lower()
    workers = max(1, min(8, int(request.form.get("workers", "4"))))
    job = Job()
    JOBS[job.id] = job
    threading.Thread(target=run_job, args=(job, f.read(), f.filename, temp, temp_unit, default_unit, workers),
                     daemon=True).start()
    return jsonify(id=job.id)


@app.get("/events/<jid>")
def events(jid):
    job = JOBS.get(jid)
    if not job:
        return jsonify(error="unknown job"), 404
    after = int(request.args.get("after", 0))
    with job.lock:
        ev = job.events[after:after + 3000]
    el = time.time() - job.t0
    rate = job.finished / el if el > 0 else 0
    eta = (job.total - job.finished) / rate if rate > 0 else None
    return jsonify(events=ev, next=after + len(ev), done=job.done, total=job.total,
                   finished=job.finished, failed=job.failed, elapsed=el, eta=eta)


@app.post("/cancel/<jid>")
def cancel(jid):
    if jid in JOBS:
        JOBS[jid].cancel = True
    return jsonify(ok=True)


@app.get("/download/<jid>")
def download(jid):
    job = JOBS.get(jid)
    if not job or not job.out_path:
        return "Not ready", 404
    return send_file(job.out_path, as_attachment=True, download_name=job.out_path.name)


@app.get("/probe")
def probe():
    """Connection test: shows what the app found on the website and one Z value."""
    msgs = []
    try:
        be = make_backend(msgs.append)
        unit = request.args.get("unit", "bar")
        z = be.z(10, unit, 30, "c")
        out = {"ok": True, "mode": be.name, "form": be.spec.describe(), "test": f"Z at 10 {unit}, 30 C = {z}"}
        if hasattr(be, "close"):
            be.close()
        return jsonify(out)
    except Exception as e:  # noqa
        return jsonify(ok=False, error=str(e), log=msgs, hint=f"Page saved to {DEBUG_HTML.name} for inspection")


PAGE = r"""<!doctype html><html><head><meta charset="utf-8"><title>CO2 Z-factor dashboard</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
:root{--bg:#0f1419;--card:#18202a;--line:#2a3441;--tx:#e6edf3;--mut:#8b98a5;--ac:#4cc2ff;--ok:#3fb950;--bad:#f85149}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font:14px system-ui,Segoe UI,sans-serif}
.wrap{max-width:1100px;margin:0 auto;padding:20px}h1{font-size:20px;margin:0 0 4px}.sub{color:var(--mut);margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px;margin-bottom:14px}
.row{display:flex;flex-wrap:wrap;gap:12px;align-items:end}label{display:flex;flex-direction:column;gap:4px;color:var(--mut);font-size:12px}
input,select{background:#0f1419;color:var(--tx);border:1px solid var(--line);border-radius:6px;padding:7px 9px;font-size:14px}
button{background:var(--ac);color:#04121c;border:0;border-radius:6px;padding:9px 14px;font-weight:600;cursor:pointer}
button.sec{background:transparent;color:var(--tx);border:1px solid var(--line)}button:disabled{opacity:.4;cursor:not-allowed}
.bar{height:10px;background:#0f1419;border-radius:6px;overflow:hidden;border:1px solid var(--line)}.bar>div{height:100%;width:0;background:var(--ac);transition:width .2s}
.stats{display:flex;gap:18px;color:var(--mut);margin-top:8px;flex-wrap:wrap}.stats b{color:var(--tx)}
.tabs{display:flex;gap:6px;margin-bottom:8px;flex-wrap:wrap}.tab{padding:5px 11px;border:1px solid var(--line);border-radius:6px;cursor:pointer;color:var(--mut)}.tab.on{color:var(--tx);border-color:var(--ac)}
.tw{max-height:380px;overflow:auto;border:1px solid var(--line);border-radius:6px}table{border-collapse:collapse;width:100%;font-variant-numeric:tabular-nums}
th,td{padding:5px 10px;text-align:right;border-bottom:1px solid var(--line)}th{position:sticky;top:0;background:var(--card);color:var(--mut)}
td.z{color:var(--ok)}td.e{color:var(--bad)}td.w{color:var(--mut)}
#log{height:120px;overflow:auto;font:12px ui-monospace,monospace;color:var(--mut);white-space:pre-wrap}
#probe{white-space:pre-wrap;font:12px ui-monospace,monospace;color:var(--mut);margin-top:8px}
</style></head><body><div class="wrap">
<h1>CO&#8322; Z-factor dashboard</h1><div class="sub">Z values come from the Peace Software online CO&#8322; calculator &mdash; each pressure is entered into the website automatically.</div>
<div class="card"><div class="row">
<label>Excel workbook<input type="file" id="file" accept=".xlsx,.xlsm"></label>
<label>Temperature<input type="number" id="temp" value="30" step="any" style="width:90px"></label>
<label>Unit<select id="tu"><option value="c">&deg;C</option><option value="k">K</option></select></label>
<label>Pressure unit if not in header<select id="unit"><option>psi</option><option>bar</option><option>kpa</option><option>mpa</option></select></label>
<label>Parallel requests<input type="number" id="workers" value="4" min="1" max="8" style="width:70px"></label>
<button id="go">Start</button><button class="sec" id="stop" disabled>Stop</button><button class="sec" id="pr">Test website connection</button>
</div><div id="probe"></div></div>
<div class="card"><div class="bar"><div id="fill"></div></div>
<div class="stats"><span>Done <b id="sd">0</b>/<b id="st">0</b></span><span>Failed <b id="sf">0</b></span><span>Elapsed <b id="se">0s</b></span><span>ETA <b id="sa">-</b></span><span id="msg"></span></div>
<div style="margin-top:10px"><a id="dl" style="display:none"><button>Download Excel</button></a></div></div>
<div class="card"><div class="tabs" id="tabs"></div><div class="tw"><table><thead><tr id="thr"></tr></thead><tbody id="tb"></tbody></table></div></div>
<div class="card"><div id="log"></div></div></div>
<script>
const $=id=>document.getElementById(id);let jid=null,after=0,timer=null,S={},cur=null;
const fmt=(n,d=5)=>n==null?'':Number(n).toFixed(d);
function render(){const s=S[cur];if(!s)return;$('thr').innerHTML='<th>Row</th>'+(s.hasT?'<th>Time</th>':'')+'<th>P1 ('+s.unit+')</th><th>P2 ('+s.unit+')</th><th>Z1</th><th>Z2</th>';
 const rows=s.rows;let h='';for(let i=0;i<rows.length;i++){const r=rows[i];h+='<tr><td>'+r.r+'</td>'+(s.hasT?'<td>'+(r.t??'')+'</td>':'')+'<td>'+s.p1+'</td><td>'+r.p+'</td><td class="'+(s.z1==null?'w':'z')+'">'+(s.z1==null?'...':fmt(s.z1))+'</td><td id="c'+i+'" class="'+(r.err?'e':r.z==null?'w':'z')+'">'+(r.err?'ERR':r.z==null?'...':fmt(r.z))+'</td></tr>'}
 $('tb').innerHTML=h}
function tabs(){$('tabs').innerHTML=Object.keys(S).map(k=>'<div class="tab '+(k==cur?'on':'')+'" onclick="cur=\''+k.replace(/'/g,"\\'")+'\';tabs();render()">'+k+'</div>').join('')}
function log(m){const l=$('log');l.textContent+=m+'\n';l.scrollTop=l.scrollHeight}
function dur(s){s=Math.round(s);return s>=60?Math.floor(s/60)+'m '+(s%60)+'s':s+'s'}
function apply(ev){for(const e of ev){
 if(e.t=='init'){S={};for(const m of e.sheets){const idx={};m.rows.forEach((r,i)=>idx[r[0]]=i);S[m.sheet]={unit:m.unit,p1:m.p1,z1:null,idx,hasT:m.rows.some(r=>r[1]!=null),rows:m.rows.map(r=>({r:r[0],t:r[1],p:r[2],z:null,err:null}))}}cur=Object.keys(S)[0];tabs();render()}
 else if(e.t=='z1'){S[e.sheet].z1=e.z;if(e.sheet==cur)render()}
 else if(e.t=='z2'){const s=S[e.sheet],i=s.idx[e.row];s.rows[i].z=e.z;s.rows[i].err=e.err;if(e.sheet==cur){const c=$('c'+i);if(c){c.className=e.err?'e':'z';c.textContent=e.err?'ERR':fmt(e.z);c.scrollIntoView({block:'nearest'})}}}
 else if(e.t=='log')log(e.msg);else if(e.t=='error'){log('ERROR: '+e.msg);$('msg').textContent='Error: '+e.msg}
 else if(e.t=='done'){$('dl').href='/download/'+jid;$('dl').style.display='inline-block';$('msg').textContent=e.cancelled?'Stopped (partial file ready)':e.failed?('Finished with '+e.failed+' failed rows'):'Finished'}}}
async function poll(){const r=await (await fetch('/events/'+jid+'?after='+after)).json();after=r.next;apply(r.events);
 $('sd').textContent=r.finished;$('st').textContent=r.total;$('sf').textContent=r.failed;$('se').textContent=dur(r.elapsed);$('sa').textContent=r.eta==null||r.done?'-':dur(r.eta);
 $('fill').style.width=(r.total?100*r.finished/r.total:0)+'%';
 if(r.done&&r.events.length==0){clearInterval(timer);$('go').disabled=false;$('stop').disabled=true}}
$('go').onclick=async()=>{const f=$('file').files[0];if(!f){alert('Choose an Excel file first');return}
 const fd=new FormData();fd.append('file',f);fd.append('temp',$('temp').value);fd.append('temp_unit',$('tu').value);fd.append('unit',$('unit').value);fd.append('workers',$('workers').value);
 $('go').disabled=true;$('stop').disabled=false;$('dl').style.display='none';$('log').textContent='';$('msg').textContent='';after=0;
 const r=await (await fetch('/start',{method:'POST',body:fd})).json();if(r.error){alert(r.error);$('go').disabled=false;return}
 jid=r.id;timer=setInterval(poll,400)};
$('stop').onclick=()=>fetch('/cancel/'+jid,{method:'POST'});
$('pr').onclick=async()=>{$('probe').textContent='Testing...';const r=await (await fetch('/probe?unit='+$('unit').value)).json();$('probe').textContent=JSON.stringify(r,null,2)};
</script></body></html>"""


if __name__ == "__main__":
    threading.Timer(1.0, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
    print(f"Dashboard running at http://127.0.0.1:{PORT}  (Ctrl+C to stop)")
    app.run(host="127.0.0.1", port=PORT, threaded=True)
