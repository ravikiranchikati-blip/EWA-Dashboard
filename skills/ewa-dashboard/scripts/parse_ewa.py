#!/usr/bin/env python3
"""
parse_ewa.py v3.4 — SAP EarlyWatch Alert (.doc / .docx) -> DEFAULT_DATA JSON

Handles all three SAP EWA file formats:
  1. Standard OOXML .docx (ZIP-based)
  2. SAP OOXML .docx with purl.oclc.org namespace variants
  3. Word 2003 XML .doc  — the native SAP ECS EWA format
     Correctly handles wx:sect section wrappers by recursively walking the body.

v3.3 changes:
  - Full capacity extraction via table-based parsing (no regex fallback for capacity):
      * HANA memory: extracts from 'HANA instance / Memory usage of SAP HANA [GB]' table
      * Physical RAM: extracts from hardware table 'Physical Memory [GB]' column
      * Disk volumes: extracts from 'Available Disk Space [GB]' table by Usage Type
        (Available = Total disk, Used = HANA files, Free = Total-Used)
      * DB Size and Growth: from 'Database Space Management' merged-cell rows
  - Fiori user count: fixed to use 'System Performance' / 'Fiori Users' row
  - SID: extracted from filename as reliable fallback
  - Product: S/4HANA release from paragraph text, not table EoM status

v3.4 changes:
  - Customer name: extracted from EWA cover page "SAP HANA Database <Name> Session No."
    Handles zero-whitespace cell concatenation in Word 2003 XML (native SAP ECS format)
  - Cloud/hardware provider: extended detection for Google Compute Engine → GCP,
    Amazon EC2 → AWS; uses hardware-context tracking to avoid false positives
  - Both fields emitted as d["customer"] and d["cloud"] in output JSON

Usage:
    python parse_ewa.py --input "path/to/ewa.doc" --output "ewa_data.json" [--indent 2]
"""
import argparse, json, re, sys, os, zipfile, io, math

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    print("ERROR: python-docx not installed. Run: pip install python-docx", file=sys.stderr)
    sys.exit(1)

# ── OOXML namespace compatibility ────────────────────────────────────────────────────
_NS_FIXES = {
    "http://purl.oclc.org/ooxml/wordprocessingml/main":
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "http://purl.oclc.org/ooxml/officeDocument/relationships":
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/drawingml/main":
        "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing":
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "http://purl.oclc.org/ooxml/officeDocument/math":
        "http://schemas.openxmlformats.org/officeDocument/math",
}


# ── Format detection ────────────────────────────────────────────────────────────
def _detect_format(path):
    with open(path, "rb") as f:
        header = f.read(12)
    if header[:2] == b"PK":
        return "ooxml"
    stripped = header.lstrip(b"\xef\xbb\xbf")
    if stripped[:5] == b"<?xml":
        return "word2003xml"
    if header[:4] == b"\xd0\xcf\x11\xe0":
        return "binary_doc"
    with open(path, "rb") as f:
        chunk = f.read(512)
    if b"wordDocument" in chunk or b"<w:" in chunk:
        return "word2003xml"
    return "ooxml"


# ── OOXML loader ─────────────────────────────────────────────────────────────────
def _fix_bytes(b):
    t = b.decode("utf-8", errors="replace")
    for old, new in _NS_FIXES.items():
        t = t.replace(old, new)
    return t.encode("utf-8")

def _open_ooxml(path):
    try:
        return Document(path)
    except Exception:
        pass
    try:
        with zipfile.ZipFile(path) as z:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
                for name in z.namelist():
                    data = z.read(name)
                    if name.endswith(".xml") or name.endswith(".rels"):
                        data = _fix_bytes(data)
                    out.writestr(name, data)
        buf.seek(0)
        return Document(buf)
    except Exception:
        return None


# ── Word 2003 XML loader ────────────────────────────────────────────────────────
def _load_word2003_xml(path):
    """
    Parse a Word 2003 XML .doc file.
    Recursively walks the body to handle wx:sect section wrappers.
    Returns (paras: list[str], rows: list[list[str]]).
    """
    try:
        from lxml import etree as ET
        _lxml = True
    except ImportError:
        import xml.etree.ElementTree as ET
        _lxml = False

    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]

    if _lxml:
        parser = ET.XMLParser(recover=True, encoding="utf-8")
        root = ET.fromstring(raw, parser=parser)
    else:
        root = ET.fromstring(raw)

    W = root.tag.split("}")[0].lstrip("{") if "}" in root.tag \
        else "http://schemas.microsoft.com/office/word/2003/wordml"

    def para_text(elem):
        return "".join(t.text or "" for t in elem.iter(f"{{{W}}}t")).strip()

    def cell_text(tc_elem):
        parts = [para_text(p) for p in tc_elem.findall(f"{{{W}}}p") if para_text(p)]
        return " ".join(parts).strip()

    paras = []
    rows  = []

    def walk(elem):
        for child in elem:
            local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if local == "p":
                txt = para_text(child)
                if txt:
                    paras.append(txt)
            elif local == "tbl":
                for tr in child.iter(f"{{{W}}}tr"):
                    cells = []
                    seen = set()
                    for tc in tr.findall(f"{{{W}}}tc"):
                        txt = cell_text(tc)
                        key = txt[:60]
                        if key not in seen:
                            seen.add(key)
                            cells.append(txt)
                    if any(cells):
                        rows.append(cells)
            elif local not in ("tr", "tc", "fonts", "styles", "docPr",
                                "docSuppData", "shapeDefaults", "listsDefn"):
                walk(child)

    _body = root.find(f"{{{W}}}body")
    body = _body if _body is not None else root
    walk(body)
    return paras, rows


# ── Number utilities ──────────────────────────────────────────────────────────────
def _eu(s):
    """Parse European-formatted number: '495,8' -> 495.8, '1.234,5' -> 1234.5"""
    if not s or not str(s).strip(): return None
    s = str(s).strip()
    if ',' in s:
        # European: period=thousands, comma=decimal
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except Exception:
        return None


# ── Main parser class ───────────────────────────────────────────────────────────
class EWAParser:
    def __init__(self, path):
        self._path = path
        fmt = _detect_format(path)
        print(f"[INFO] Format detected: {fmt}  ({path})", file=sys.stderr)
        self.paras = []; self.rows = []; self.doc = None

        if fmt == "ooxml":
            self.doc = _open_ooxml(path)
            if self.doc is None:
                raise ValueError("Failed to open OOXML file.")
            self._load_ooxml()
        elif fmt == "word2003xml":
            self.paras, self.rows = _load_word2003_xml(path)
        elif fmt == "binary_doc":
            raise ValueError(
                "Old binary .doc format is not supported. "
                "Open in Word and save as .docx."
            )
        else:
            self.doc = _open_ooxml(path)
            if self.doc:
                self._load_ooxml()
            else:
                self.paras, self.rows = _load_word2003_xml(path)

        self.corpus = (
            "\n".join(self.paras) + "\n" +
            "\n".join("\t".join(r) for r in self.rows)
        )
        print(f"[INFO] Loaded {len(self.paras)} paragraphs, {len(self.rows)} table rows",
              file=sys.stderr)

    def _load_ooxml(self):
        for para in self.doc.paragraphs:
            t = para.text.strip()
            if t:
                self.paras.append(t)
        for table in self.doc.tables:
            for row in table.rows:
                cells = []
                seen = set()
                for cell in row.cells:
                    txt = cell.text.strip()
                    key = txt[:60]
                    if key not in seen:
                        seen.add(key)
                        cells.append(txt)
                if any(cells):
                    self.rows.append(cells)

    # ── helpers ────────────────────────────────────────────────────────────────
    def find(self, patterns, default=None):
        for pat in patterns:
            try:
                m = re.search(pat, self.corpus, re.I | re.M)
                if m:
                    g = (m.group(1) if m.lastindex else m.group(0)).strip()
                    if g: return g
            except Exception: pass
        return default

    def find_n(self, patterns, default=None):
        v = self.find(patterns)
        if v is None: return default
        try: return float(re.sub(r"[^0-9.\-]", "", v.replace(",", "")))
        except Exception: return default

    def tr(self, label_re, col=1):
        pat = re.compile(label_re, re.I)
        for row in self.rows:
            if row and pat.search(row[0]):
                if col < len(row):
                    v = row[col].strip()
                    return v if v else None
        return None

    def tr_n(self, label_re, col=1, default=None):
        v = self.tr(label_re, col)
        if v is None: return default
        try:
            c = re.sub(r"[^0-9.\-]", "", v.replace(",", ""))
            return float(c) if c else default
        except Exception: return default

    def tr_first_n(self, label_re, default=None):
        pat = re.compile(label_re, re.I)
        for row in self.rows:
            if row and pat.search(row[0]):
                for cell in row[1:]:
                    try:
                        n = float(re.sub(r"[^0-9.\-]", "", cell.replace(",", "")))
                        if n >= 0: return n
                    except Exception: pass
        return default

    def extract_workload(self, type_re):
        pat = re.compile(r"^\s*" + type_re + r"\s*$", re.I)
        best = None; best_n = -1
        for row in self.rows:
            if not row or not pat.match(row[0].strip()): continue
            nums = sum(1 for c in row[1:] if self._safe_n(c) is not None)
            if nums > best_n: best = row; best_n = nums
        if not best: return {}
        nums = []
        for cell in best[1:]:
            if re.search(r"(yellow|amber|green|red|ok|warning|n/a|threshold)", cell, re.I): continue
            n = self._safe_n(cell)
            if n is not None: nums.append(n)
        if not nums: return {}
        result = {"resp": nums[0]}
        steps_i = None
        for i in range(len(nums) - 1, 0, -1):
            if nums[i] > 1000 and nums[i] == int(nums[i]): steps_i = i; break
        if steps_i is None and len(nums) > 3: steps_i = len(nums) - 1
        if steps_i is not None:
            result["steps"] = nums[steps_i]
            rest = [n for i, n in enumerate(nums[1:], 1) if i != steps_i]
            if rest: result["db"] = rest[0]
            if len(rest) > 1: result["cpu"] = rest[1]
        elif len(nums) >= 3:
            result["db"] = nums[1]; result["cpu"] = nums[2]
        return result

    def _safe_n(self, s):
        if not s or not s.strip(): return None
        c = re.sub(r"[^0-9.\-]", "", s.replace(",", ""))
        try: return float(c) if c else None
        except Exception: return None

    # ── capacity extraction ─────────────────────────────────────────────────────────
    def _extract_capacity(self):
        """
        Extract HANA memory, disk volume, and physical hardware capacity
        from structured EWA tables.

        EWA disk table interpretation:
          'Available Disk Space [GB]' = TOTAL volume size (not free!)
          'Used Disk Space [GB]'      = HANA data/log file usage
          'Percentage of free Disk Space' = (Total-Used)/Total * 100
          Free GB = Total - Used

        EWA HANA memory table:
          Header: ['HANA instance', 'Memory usage of SAP HANA [GB]',
                   'Global allocation limit [GB]', 'Row store size [GB]',
                   'Column store size in main memory [GB]',
                   'Memory usage of indexserver [GB]',
                   'Effective allocation limit of indexserver [GB]']
          Data row: [hostname_SID_NN, mem_used, alloc_limit, row, col, idx_mem, idx_limit]

        EWA hardware table (NUMA nodes may be absent from data row):
          Header: ['Host','CPU Type','CPU Frequency','CPU Cores','Threads',
                   'Sockets','NUMA Nodes','Physical Memory [GB]',...]
          Maps col position by header name to handle missing columns.
        """
        cap = {}

        # 1. HANA memory instance table
        hana_mem_hdr = False
        for row in self.rows:
            joined = ' '.join(row).lower()
            if 'memory usage of sap hana' in joined and 'global allocation limit' in joined:
                hana_mem_hdr = True
                continue
            if hana_mem_hdr and len(row) >= 6:
                # row: [instance, mem_used_gb, alloc_limit_gb, row_store_gb,
                #       col_store_gb, indexserver_mem_gb, indexserver_limit_gb]
                vals = [_eu(c) for c in row]
                if any(v is not None and v > 0 for v in vals[1:]):
                    if vals[1] is not None: cap['hanaMemUsed']     = round(vals[1], 1)
                    if vals[2] is not None: cap['hanaAllocLimit']  = round(vals[2], 1)
                    if vals[3] is not None: cap['rowStore']        = round(vals[3], 1)
                    if vals[4] is not None: cap['colStore']        = round(vals[4], 1)
                    if vals[5] is not None: cap['indexserverMem']  = round(vals[5], 1)
                    if len(vals) > 6 and vals[6] is not None:
                        cap['indexserverLimit'] = round(vals[6], 1)
                    hana_mem_hdr = False
            elif hana_mem_hdr and row and row[0].lower() in ('', 'hana instance',
                                                              'hana role', 'service'):
                pass  # still in header area
            elif hana_mem_hdr:
                hana_mem_hdr = False

        # 2. Physical hardware table — map by header column names
        hw_hdr_row = None
        for i, row in enumerate(self.rows):
            joined = ' '.join(row).lower()
            if 'physical memory' in joined and 'cpu' in joined and 'host' in row[0].lower():
                hw_hdr_row = row
                # Find Physical Memory column index
                phys_idx = alloc_idx = None
                for ci, cell in enumerate(row):
                    cl = cell.lower()
                    if 'physical memory' in cl: phys_idx = ci
                    if 'allocation limit' in cl and phys_idx is not None: alloc_idx = ci
                # Next non-header data row
                for j in range(i + 1, min(i + 5, len(self.rows))):
                    dr = self.rows[j]
                    if not dr or dr[0].lower() in ('host', ''): continue
                    if phys_idx and phys_idx < len(dr):
                        v = _eu(dr[phys_idx])
                        if v and 50 < v < 10000:
                            cap['hanaPhysRam'] = round(v, 1)
                    if alloc_idx and alloc_idx < len(dr) and cap.get('hanaAllocLimit') is None:
                        v = _eu(dr[alloc_idx])
                        if v and 50 < v < 10000:
                            cap['hanaAllocLimit'] = round(v, 1)
                    break
                break

        # Fallback for physRam from ini: allocationlimit in MB
        if cap.get('hanaPhysRam') is None:
            for row in self.rows:
                if 'allocationlimit' in ' '.join(row).lower() and len(row) >= 4:
                    v_mb = _eu(row[3])
                    if v_mb and v_mb > 1000:
                        cap.setdefault('hanaAllocLimit', round(v_mb / 1024, 1))

        # 3. Disk space table
        # Header: Host | Available Disk Space [GB] | Used Disk Space [GB] |
        #         Percentage of free Disk Space | Usage Types | File system | Rating
        # IMPORTANT: 'Available Disk Space' = TOTAL disk, NOT free
        disk_hdr_idx = None
        for i, row in enumerate(self.rows):
            joined = ' '.join(row).lower()
            if ('available disk space' in joined and
                    'used disk space' in joined and
                    'host' in row[0].lower()):
                disk_hdr_idx = i
                # Map column positions
                total_idx = used_idx = pct_idx = usage_idx = None
                for ci, cell in enumerate(row):
                    cl = cell.lower()
                    if 'available disk space' in cl: total_idx = ci
                    elif 'used disk space' in cl:    used_idx  = ci
                    elif 'percentage' in cl:         pct_idx   = ci
                    elif 'usage type' in cl:         usage_idx = ci
                for j in range(i + 1, min(i + 15, len(self.rows))):
                    dr = self.rows[j]
                    if not dr: continue
                    if dr[0].lower() == 'host': continue
                    # Stop if no longer a disk row
                    if len(dr) < 3: break
                    total = _eu(dr[total_idx]) if total_idx and total_idx < len(dr) else None
                    used  = _eu(dr[used_idx])  if used_idx  and used_idx  < len(dr) else None
                    pct_s = dr[pct_idx].replace('%','').strip() if pct_idx and pct_idx < len(dr) else None
                    pct   = _eu(pct_s) if pct_s else None
                    usage = dr[usage_idx] if usage_idx and usage_idx < len(dr) else (dr[4] if len(dr) > 4 else '')
                    # Skip bogus rows (filesystem volumes > 1PB)
                    if total and total > 100000: continue
                    if 'DATA' in usage and cap.get('dataVolTotal') is None and total and used:
                        free = round(total - used)
                        cap['dataVolTotal'] = round(total)
                        cap['dataVolUsed']  = round(used)
                        cap['dataVolFree']  = free
                        cap['dataFreePct']  = round(pct) if pct is not None else (
                            round(free / total * 100) if total else None)
                    elif 'LOG' in usage and cap.get('logVolTotal') is None and total and used:
                        free = round(total - used)
                        cap['logVolTotal'] = round(total)
                        cap['logVolUsed']  = round(used)
                        cap['logVolFree']  = free
                        cap['logFreePct']  = round(pct) if pct is not None else (
                            round(free / total * 100) if total else None)
                break

        # Fallback: /hana/data and /hana/log path-based lookup
        if cap.get('dataVolTotal') is None:
            for row in self.rows:
                r0 = row[0].lower() if row else ''
                if '/hana/data' in r0 or 'data volume' in r0:
                    nums = [_eu(c) for c in row[1:] if _eu(c) is not None and 0 < (_eu(c) or 0) < 100000]
                    if len(nums) >= 3:
                        cap['dataVolUsed'] = round(nums[0])
                        cap['dataVolFree'] = round(nums[1])
                        cap['dataVolTotal'] = round(nums[2])
                    break
        if cap.get('logVolTotal') is None:
            for row in self.rows:
                r0 = row[0].lower() if row else ''
                if '/hana/log' in r0 or 'log volume' in r0:
                    nums = [_eu(c) for c in row[1:] if _eu(c) is not None and 0 < (_eu(c) or 0) < 100000]
                    if len(nums) >= 3:
                        cap['logVolUsed'] = round(nums[0])
                        cap['logVolFree'] = round(nums[1])
                        cap['logVolTotal'] = round(nums[2])
                    break

        # 4. DB Size and Growth from 'Database Space Management' table
        for row in self.rows:
            if not row: continue
            joined = ' '.join(row)
            if 'Database Space Management' in joined and 'DB Size' in joined:
                m = re.search(r'([\d,\.]+)\s*GB', joined)
                if m and cap.get('dbSize') is None:
                    raw = m.group(1).replace(',', '.').replace(' ', '')
                    try: cap['dbSize'] = round(float(raw), 2)
                    except Exception: pass
            if 'Database Space Management' in joined and 'DB Growth' in joined:
                m = re.search(r'([\d,\.]+)\s*GB', joined)
                if m and cap.get('dbGrowth') is None:
                    raw = m.group(1).replace(',', '.').replace(' ', '')
                    try: cap['dbGrowth'] = round(float(raw), 2)
                    except Exception: pass

        return cap

    # ── extraction ──────────────────────────────────────────────────────────
    def parse(self):
        d = {}

        # ── SID: prefer filename extraction, fallback to table ──
        sid_from_fn = None
        fn = os.path.basename(self._path)
        m_fn = re.match(r'^([A-Z][A-Z0-9]{2})_', fn)
        if m_fn:
            sid_from_fn = m_fn.group(1)

        sid_from_table = None
        for row in self.rows:
            if not row: continue
            r0 = row[0]
            # Look for rows like: ['PS4', 'SPS Stack', ...] but not ['SID', 'SPS Stack', ...]
            if (re.match(r'^[A-Z][A-Z0-9]{2}$', r0.strip()) and
                    len(row) > 1 and row[0] != 'SID'):
                sid_from_table = r0.strip()
                break

        d["sid"] = sid_from_fn or sid_from_table or (
            self.find([r"\bSID[:\s]+([A-Z][A-Z0-9]{2})\b",
                       r"System\s*ID[:\s]+([A-Z][A-Z0-9]{2})\b"]) or "UNKNOWN"
        )

        # ── Product: prefer paragraph scan for S/4HANA version ──
        product_from_para = None
        for p in self.paras:
            m = re.search(r'(SAP\s+S/4HANA\s+[0-9]{4}(?:\s+FPS[0-9]{2})?)', p)
            if m:
                product_from_para = m.group(1).strip()
                break
        if not product_from_para:
            for row in self.rows:
                joined = ' '.join(row)
                m = re.search(r'(SAP\s+S/4HANA\s+[0-9]{4}(?:\s+FPS[0-9]{2})?)', joined)
                if m:
                    product_from_para = m.group(1).strip()
                    break
        d["product"] = product_from_para or (
            self.find([r"(SAP\s+S/4HANA\s+[0-9]{4}(?:\s+FPS[0-9]+)?)",
                       r"(S/4HANA\s+[0-9]{4})", r"(SAP\s+ERP\s+[0-9.]+)",
                       r"(SAP\s+ECC\s+[0-9.]+)"]) or "SAP System"
        )

        d["hanaVersion"] = (
            self.tr(r"SAP\s+HANA(?:\s+(?:DB|Database))?(?:\s+(?:Version|Release|Revision))?") or
            self.find([r"(2\.00\.[0-9]{3}\.[0-9]+)",
                       r"SAP\s+HANA[^0-9]*(2\.[0-9]+\.[0-9]+\.[0-9]+)"])
        )
        d["hanaSPS"] = (
            self.find([r"(SPS\s*0?[0-9]{1,2})", r"HANA\s+SPS[:\s]+([0-9]{1,2})"]) or
            self.tr(r"Support\s+Package\s+Stack|^SPS$")
        )
        d["kernelVersion"] = (
            self.tr(r"(?:SAP\s+)?[Kk]ernel(?:\s+(?:Release|Version))?") or
            self.find([r"[Kk]ernel[\s:]+([0-9]+\s*(?:PL|Patch|patch)\s*[0-9]+)"])
        )
        d["kernelAgeMo"] = self.find_n(
            [r"[Kk]ernel[^\n]*([0-9]+)\s+months?", r"([0-9]+)\s+months?[^\n]*[Kk]ernel"]
        )
        d["period"] = (
            self.tr(r"(?:Monitoring\s+)?Period|Reporting\s+(?:Period|Week)") or
            self.find([r"(\d{2}\.\d{2}\.\d{4}\s*[\u2013\-]+\s*\d{2}\.\d{2}\.\d{4})",
                       r"Period[:\s]+(\d{2}\.\d{2}\.\d{4}[^\n]{3,30})"])
        )
        d["session"] = (
            self.tr(r"Session(?:\s+(?:No\.?|Number|ID))?|EWA\s+Session") or
            self.find([r"[Ss]ession[:\s#]+([0-9]{6,12})"])
        )
        # Installation from filename
        m_inst = re.search(r'^[A-Z][A-Z0-9]{2}_([0-9]{6,13})_', fn)
        d["installation"] = (m_inst.group(1) if m_inst else None) or (
            self.tr(r"Installation(?:\s+(?:No\.?|Number))?") or
            self.find([r"[Ii]nstallation[:\s#]+([0-9]{6,13})"])
        )
        # ── Customer name: from EWA cover page "SAP HANA Database <Name> Session No." ──
        # EWA cover page table cells are concatenated without spaces in Word 2003 XML:
        # "SAP HANA DatabaseCardinal Health Inc.Session No." (zero-ws between tokens)
        # Use \s* (zero or more) instead of \s+ to handle both zero and whitespace cases
        _cust_pat = re.compile(
            r'SAP\s*HANA\s*(?:Database|DB)\s*([A-Z][^\t\n\r<]{3,70}?)\s*Session\s*No\.',
            re.I
        )
        d["customer"] = None
        _cm = _cust_pat.search(self.corpus)
        if _cm:
            _raw = _cm.group(1).strip()
            if re.search(r'[A-Za-z]{2}', _raw) and 4 <= len(_raw) <= 80:
                d["customer"] = _raw

        d["cloud"] = None
        hw_context = False
        for row in self.rows:
            joined = ' '.join(row)
            jl = joined.lower()
            # Track hardware/manufacturer table context
            if any(h in jl for h in ['manufacturer', 'hardware', 'virtualization',
                                      'cpu type', 'cpu freq', 'physical memory']):
                hw_context = True
            # Specific long-form patterns first
            if re.search(r'Amazon\s+(?:Web\s+Services|EC2|Elastic\s+Compute)', joined, re.I):
                d["cloud"] = "AWS"; break
            if re.search(r'Microsoft\s+Azure|Azure\s+Virtual\s+Machine', joined, re.I):
                d["cloud"] = "Azure"; break
            if re.search(r'Google\s+(?:Cloud|Compute\s+Engine)', joined, re.I):
                d["cloud"] = "GCP"; break
            # Hardware-context generic vendor names
            if hw_context:
                if re.search(r'\bGoogle\b', joined):
                    d["cloud"] = "GCP"; break
                if re.search(r'\bAmazon\b|\bEC2\b|\bAWS\b', joined, re.I):
                    d["cloud"] = "AWS"; break
                if re.search(r'\bAzure\b', joined, re.I):
                    d["cloud"] = "Azure"; break
        if not d["cloud"]:
            d["cloud"] = self.find([r"(AWS)", r"(Azure)", r"(GCP)",
                                    r"(On[- ]?[Pp]rem(?:ise)?)"])
        d["sysReplication"] = (
            self.tr(r"System\s+Replication|HANA\s+(?:System\s+)?Replication|HSR") or
            self.find([r"[Ss]ystem\s+[Rr]eplication[:\s]+([^\n]{3,60})"])
        )
        d["spAgeMo"] = (
            self.tr_n(r"Support\s+Package\s+Age|SP\s+Age") or
            self.find_n([r"SP\s+[Aa]ge[^0-9]*([0-9]+)\s*months?",
                         r"([0-9]+)\s+months?[^\n]*[Ss]upport\s+[Pp]ackage",
                         r"[Ss]upport\s+[Pp]ackage[^0-9]*([0-9]+)\s*months?"])
        )
        d["availability"] = self.find_n(
            [r"[Aa]vailability[:\s]+(\d+(?:\.\d+)?)\s*%",
             r"(\d+(?:\.\d+)?)\s*%\s*[Aa]vailability"], 100
        )
        d["activeUsers"] = (
            self.tr_first_n(r"Active\s+Users?|Users?\s+Logged(?:\s+On|\s+In)?") or
            self.find_n([r"([0-9]+)\s+[Aa]ctive\s+[Uu]sers?"])
        )
        # Fiori users: look for 'System Performance' + 'Fiori Users' row specifically
        fiori_val = None
        for i, row in enumerate(self.rows):
            if 'System Performance' in row and 'Fiori Users' in row:
                # Next numeric cell in same or next row
                for cell in row:
                    if cell not in ('System Performance', 'Fiori Users', '') and self._safe_n(cell):
                        v = self._safe_n(cell)
                        if v and v < 100000: fiori_val = v; break
                if fiori_val: break
            if 'Fiori Users' in ' '.join(row):
                for cell in row:
                    if 'Fiori' not in cell and 'System' not in cell and self._safe_n(cell):
                        v = self._safe_n(cell)
                        if v and v < 100000: fiori_val = v; break
                if fiori_val: break
        d["fioriUsers"] = fiori_val or self.find_n([r"[Ff]iori[^0-9]*([0-9]+)\s+[Uu]sers?"])
        # Total users: look for 'Users' header table
        total_u = None
        for row in self.rows:
            if 'Total Users' in row and 'Users' == row[0]:
                # Next row has counts
                break
        d["totalUsers"] = total_u or (
            self.tr_first_n(r"Total(?:\s+[Uu]sers?)?|All\s+[Uu]sers?") or
            self.find_n([r"[Tt]otal\s+[Uu]sers?[:\s]+([0-9]+)"])
        )

        dlg = self.extract_workload(r"Dialog(?:\s+(?:Steps?|Tasks?|Work\s+Processes?))?|DIALOG")
        rfc = self.extract_workload(r"RFC(?:\s+(?:Steps?|Tasks?))?")
        bch = self.extract_workload(r"[Bb]atch(?:\s+(?:Steps?|Jobs?))?|Background")
        upd = self.extract_workload(r"Update(?:\s+Tasks?)?")
        htp = self.extract_workload(r"HTTPS?(?:\s+(?:Steps?|Tasks?))?|Web\s+Services?")
        d["dialogRespMs"] = dlg.get("resp"); d["dialogDbMs"]  = dlg.get("db")
        d["dialogCpuMs"]  = dlg.get("cpu"); d["stepsDialog"] = dlg.get("steps")
        d["rfcRespMs"]    = rfc.get("resp"); d["rfcDbMs"]     = rfc.get("db")
        d["rfcCpuMs"]     = rfc.get("cpu"); d["stepsRfc"]    = rfc.get("steps")
        d["batchRespMs"]  = bch.get("resp"); d["batchDbMs"]   = bch.get("db")
        d["batchCpuMs"]   = bch.get("cpu"); d["stepsBatch"]  = bch.get("steps")
        d["updateRespMs"] = upd.get("resp"); d["updateDbMs"]  = upd.get("db")
        d["updateCpuMs"]  = upd.get("cpu"); d["stepsUpdate"] = upd.get("steps")
        d["httpsRespMs"]  = htp.get("resp"); d["httpsDbMs"]   = htp.get("db")
        d["httpsCpuMs"]   = htp.get("cpu"); d["stepsHttps"]  = htp.get("steps")

        d["maxDbCpu"]  = self.find_n([r"[Dd][Bb][^\n]*?([0-9]+)\s*%[^\n]*CPU",
                                      r"CPU[^\n]*(?:[Dd]atabase|DB)[^\n]*?([0-9]+)\s*%"])
        d["maxAppCpu"] = self.find_n([r"[Aa]pp(?:lication)?[^\n]*CPU[^\n]*?([0-9]+)\s*%"])

        # ── Capacity: use table-based extraction ──
        cap = self._extract_capacity()
        d["hanaPhysRam"]      = cap.get("hanaPhysRam")
        d["hanaAllocLimit"]   = cap.get("hanaAllocLimit")
        d["hanaMemUsed"]      = cap.get("hanaMemUsed")
        d["colStore"]         = cap.get("colStore")
        d["rowStore"]         = cap.get("rowStore")
        d["indexserverMem"]   = cap.get("indexserverMem")
        d["indexserverLimit"] = cap.get("indexserverLimit") or cap.get("hanaAllocLimit")
        d["dataVolTotal"]     = cap.get("dataVolTotal")
        d["dataVolUsed"]      = cap.get("dataVolUsed")
        d["dataVolFree"]      = cap.get("dataVolFree")
        d["dataFreePct"]      = cap.get("dataFreePct")
        d["logVolTotal"]      = cap.get("logVolTotal")
        d["logVolUsed"]       = cap.get("logVolUsed")
        d["logVolFree"]       = cap.get("logVolFree")
        d["logFreePct"]       = cap.get("logFreePct")
        d["dbSize"]           = cap.get("dbSize") or (
            self.tr_n(r"(?:Total\s+)?(?:DB|Database)\s+Size") or
            self.find_n([r"[Dd]atabase\s+[Ss]ize[^\n]*?([0-9,]+)\s*GB"]))
        d["dbGrowth"]         = cap.get("dbGrowth") or (
            self.find_n([r"[Gg]rowth[^\n]*?([0-9]+\.?[0-9]*)\s*GB"]))

        d["backupFailures"]     = None  # Avoid matching SAP Note numbers
        d["backupDays"]         = self.find([r"[Bb]ackup[^\n]*?([0-9]+/[0-9]+\s*(?:days?)?)"])
        # Infer backup health from log backup rows
        for row in self.rows:
            joined = ' '.join(row).lower()
            if 'automatic log backup' in joined and len(row) >= 2:
                if 'yes' in row[1].lower():
                    if not d["backupDays"]:
                        d["backupDays"] = '7/7'
                    d["backupFailures"] = 0
                break
        d["lastConsistencyCheck"] = self.find([r"[Cc]onsistency[^\n]*?(\d{2}\.\d{2}\.\d{4})"])
        d["abapDumps"] = (
            self.tr_n(r"(?:ABAP\s+)?[Rr]untime\s+[Ee]rrors?|ST22|[Dd]umps?") or
            self.find_n([r"([0-9]+)\s+(?:ABAP\s+)?[Rr]untime\s+[Ee]rrors?",
                         r"(?:ABAP\s+)?[Dd]umps?[^\n]*?([0-9]+)"])
        )
        d["peakDumpDate"]  = self.find([r"[Pp]eak[^\n]*[Dd]ump[^\n]*?(\d{2}\.\d{2}\.\d{4})"])
        d["peakDumpCount"] = self.find_n([r"[Pp]eak[^\n]*[Dd]ump[^\n]*?([0-9]+)"])
        d["transportSeqErrors"] = (
            self.tr_n(r"[Ss]equence[^\n]*[Ee]rrors?|[Oo]vertaken|[Bb]ypassed") or
            self.find_n([r"([0-9]+)[^\n]*[Ss]equence\s+[Ee]rrors?"])
        )
        d["transportCount"]    = self.find_n([r"([0-9]+)\s+[Tt]ransports?\s+(?:imported|released)"])
        d["failedImports"]     = self.find_n([r"[Ff]ailed\s+[Ii]mports?[^\n]*?([0-9]+)"], 0)
        d["updateErrors"]      = self.find_n([r"[Uu]pdate\s+[Ee]rrors?[^\n]*?([0-9]+)"], 0)
        d["shortLeadTime"]     = self.find_n([r"[Ss]hort\s+[Ll]ead\s+[Tt]ime[^\n]*?([0-9]+)"])
        d["emergencyTransports"] = self.find_n([r"[Ee]mergency[^\n]*[Tt]ransport[^\n]*?([0-9]+)"], 0)
        d["debugUsers"] = (
            self.tr_n(r"S_DEVELOP|DEBUG\s+[Uu]sers?|[Dd]ebug[^\n]*[Uu]sers?") or
            self.find_n([r"([0-9]+)\s+[Uu]sers?[^\n]*DEBUG", r"S_DEVELOP[^\n]*?([0-9]+)"])
        )
        d["defaultPasswords"] = bool(self.find([
            r"(?:default|standard)\s+password",
            r"SAP\*[^\n]*password", r"SAPCPIC[^\n]*password", r"TMSADM[^\n]*password"
        ]))
        d["firefighterAccounts"] = self.find_n([r"[Ff]irefighter[^\n]*?([0-9]+)"])
        d["hanaAlert59"] = bool(self.find([r"[Aa]lert\s+59|[Bb]locked?\s+[Tt]ransactions?"]))
        d["blockingTxnCount"] = self.find_n([r"[Bb]locked?\s+[Tt]ransactions?[^\n]*?([0-9]+)"])
        d["topSqlHash"]    = self.find([r"\b([0-9a-f]{8})\b"])
        d["topSqlTable"]   = self.find([r"[Tt]able[:\s]+([A-Z][A-Z0-9_]{2,})",
                                        r"INSERT\s+(?:INTO\s+)?([A-Z][A-Z0-9_]{2,})"])
        d["topSqlProgram"] = self.find([r"[Pp]rogram[:\s]+([A-Z][A-Z0-9_/]{2,})"])
        d["topSqlExecCount"] = self.find_n([r"([0-9,]+)\s*[Ee]xecutions?"])
        d["topSqlLockPct"]   = self.find_n([r"[Ll]ock[^\n]*?([0-9]+\.?[0-9]*)\s*%"])

        top_txns = []
        for row in self.rows:
            if len(row) < 2: continue
            r0 = row[0]
            if re.match(r"^[A-Z][A-Z0-9/_]{1,19}$", r0) and len(r0) >= 2:
                pct = None; desc = ""
                for cell in row[1:]:
                    n = self._safe_n(cell)
                    if n is not None and 0 < n <= 100 and pct is None: pct = n
                    elif len(cell) > 4 and not re.match(r"^[0-9.,% ]+$", cell): desc = cell[:100]
                if pct: top_txns.append({"code": r0, "pct": pct, "desc": desc})
        top_txns.sort(key=lambda x: x["pct"], reverse=True)
        d["topTxns"] = top_txns[:5] if top_txns else []

        largest = []
        for row in self.rows:
            if len(row) < 2: continue
            r0 = row[0]
            if re.match(r"^[A-Z][A-Z0-9/_]{2,}$", r0):
                size = None; delta = None
                for cell in row[1:]:
                    n = self._safe_n(cell)
                    if n is not None and n > 10:
                        if size is None: size = n
                        elif delta is None: delta = n
                if size and size > 100:
                    largest.append({"name": r0, "sizeMB": round(size),
                                    "deltaMB": round(delta) if delta else None})
        largest.sort(key=lambda x: x["sizeMB"], reverse=True)
        # Exclude the SID itself from largest tables list
        largest = [t for t in largest if t["name"] != d.get("sid")]
        if largest:
            d["largestTable"]  = largest[0]["name"]
            d["baldatSizeMB"]  = largest[0]["sizeMB"]
            d["baldatDeltaMB"] = largest[0]["deltaMB"]

        check_overview = []; in_check = False
        for row in self.rows:
            if not row: continue
            r0l = row[0].lower()
            if re.search(r"check\s+overview|category|area.*(?:check|alert|status)", r0l):
                in_check = True; continue
            if not in_check: continue
            if len(row) < 2: continue
            area = row[0].strip()
            if not area or len(area) < 5: continue
            checks_n = None; alerts_n = None; rag = "ok"
            for cell in row[1:]:
                n = self._safe_n(cell)
                if n is not None:
                    if checks_n is None: checks_n = int(n)
                    elif alerts_n is None: alerts_n = int(n)
                cl = cell.lower()
                if any(w in cl for w in ("red","fail","critical")): rag = "err"
                elif any(w in cl for w in ("amber","yellow","warning")): rag = "warn"
                elif any(w in cl for w in ("green","ok","pass")): rag = "ok"
            if checks_n is not None:
                if alerts_n and alerts_n > 0 and rag == "ok": rag = "warn"
                check_overview.append({"area": area, "checks": checks_n or 0,
                                       "alerts": alerts_n or 0, "rag": rag})
        d["checkOverview"] = check_overview[:20]

        alert_overview = []; in_alert = False; alert_id = 1
        for row in self.rows:
            if not row: continue
            r0l = row[0].lower()
            if re.search(r"alert\s+overview|priority.*alert|alert.*priority", r0l):
                in_alert = True; continue
            if not in_alert: continue
            if len(row) < 2: continue
            prio_text = ""
            for cell in row[:3]:
                m = re.search(r"\b(high|medium|low|critical)\b", cell, re.I)
                if m: prio_text = m.group(1).lower(); break
            if not prio_text: continue
            pmap = {"critical": "crit", "high": "high", "medium": "med", "low": "low"}
            prio = pmap.get(prio_text, "med")
            desc = ""; action = ""
            for cell in row:
                if len(cell) > 20 and not re.match(r"^\d+$", cell) and "priority" not in cell.lower():
                    if not desc: desc = cell[:300]
                    elif not action: action = cell[:300]
            if desc:
                alert_overview.append({"id": str(alert_id), "area": "", "priority": prio,
                                        "count": 1, "desc": desc, "action": action,
                                        "owner": "Basis", "tl": "14 days"})
                alert_id += 1
        d["alertOverview"] = alert_overview[:15]

        d["abapBufferProg"]  = self.find_n([r"[Pp]rogram\s+[Bb]uffer[^\n]*?([0-9]+\.?[0-9]*)\s*%"])
        d["abapBufferTable"] = self.find_n([r"[Tt]able\s+[Bb]uffer[^\n]*?([0-9]+\.?[0-9]*)\s*%"])
        d["abapBufferField"] = self.find_n([r"[Ff]ield\s+[Bb]uffer[^\n]*?([0-9]+\.?[0-9]*)\s*%"])
        d["abapBufferPfRole"]= self.find_n([r"[Pp]rofile[^\n]*[Bb]uffer[^\n]*?([0-9]+\.?[0-9]*)\s*%"])
        d["enqueueErrors"]   = self.find_n([r"[Ee]nqueue[^\n]*[Ee]rrors?[^\n]*?([0-9]+)"], 0)
        d["enqueueAvgMs"]    = self.find_n([r"[Ee]nqueue[^\n]*[Aa]vg[^\n]*?([0-9]+)\s*ms"])
        d["lockTimeoutPct"]  = self.find_n([r"[Ll]ock\s+[Tt]imeout[^\n]*?([0-9]+\.?[0-9]*)\s*%"])
        d["deltaMergeErrors"]= self.find_n([r"[Dd]elta\s+[Mm]erge[^\n]*[Ee]rrors?[^\n]*?([0-9]+)"], 0)
        d["savepointDurationS"] = self.find_n([r"[Ss]avepoint[^\n]*[Dd]uration[^\n]*?([0-9]+\.?[0-9]*)\s*s"])
        d["savepointCount"]  = self.find_n([r"[Ss]avepoint[^\n]*[Cc]ount[^\n]*?([0-9,]+)"])
        d["hanaPersistenceErrors"] = self.find_n([r"[Pp]ersistence[^\n]*[Ee]rrors?[^\n]*?([0-9]+)"], 0)
        d["icmMaxConnections"]    = self.find_n([r"ICM[^\n]*[Mm]ax[^\n]*[Cc]onnections?[^\n]*?([0-9]+)"])
        d["icmActiveConnections"] = self.find_n([r"ICM[^\n]*[Aa]ctive[^\n]*[Cc]onnections?[^\n]*?([0-9]+)"])
        d["icmMaxThreadsPct"]     = self.find_n([r"ICM[^\n]*[Tt]hread[^\n]*?([0-9]+)\s*%"])
        d["qrfcErrors"]  = self.find_n([r"qRFC[^\n]*[Ee]rrors?[^\n]*?([0-9]+)"], 0)
        d["trfcErrors"]  = self.find_n([r"tRFC[^\n]*[Ee]rrors?[^\n]*?([0-9]+)"], 0)
        d["spoolIssues"] = self.find_n([r"[Ss]pool[^\n]*[Ii]ssues?[^\n]*?([0-9]+)"], 0)
        d["missingStatistics"]  = self.find_n([r"[Mm]issing\s+[Ss]tatistics?[^\n]*?([0-9]+)"], 0)
        d["outdatedStatistics"] = self.find_n([r"[Oo]utdated\s+[Ss]tatistics?[^\n]*?([0-9]+)"], 0)
        d["osVersion"] = (self.tr(r"[Oo]perating\s+[Ss]ystem") or self.find([r"(SUSE\s+Linux[^\n]{3,40})", r"(Red\s+Hat[^\n]{3,40})", r"(Windows\s+Server[^\n]{3,30})"]))
        tech_debt = []
        if (d.get("spAgeMo") or 0) >= 18:
            tech_debt.append({"area":"SP Stack","priority":"high" if (d.get("spAgeMo") or 0)>=22 else "med","effort":"High","item":"Core components %s months old - SP update required"%d.get("spAgeMo","?"),"impact":"Security coverage + supportability"})
        if d.get("kernelAgeMo") and d["kernelAgeMo"]>=6:
            tech_debt.append({"area":"Kernel","priority":"med","effort":"Low","item":"Kernel %s (%s months old) - update to latest patch level"%(d.get("kernelVersion","?"),d["kernelAgeMo"]),"impact":"Security + stability"})
        if d.get("debugUsers") and d["debugUsers"]>0:
            tech_debt.append({"area":"Authorization","priority":"crit","effort":"Low","item":"S_DEVELOP DEBUG for %s production user(s) - immediate removal required"%d["debugUsers"],"impact":"Security compliance (critical risk)"})
        if d.get("defaultPasswords"):
            tech_debt.append({"area":"Passwords","priority":"crit","effort":"Low","item":"Default SAP*/SAPCPIC passwords active - immediate reset required","impact":"Security compliance (critical risk)"})
        if (d.get("baldatSizeMB") or 0)>10000:
            tech_debt.append({"area":"Housekeeping","priority":"med","effort":"Low","item":"%s %s MB - implement automated deletion (SAP Note 195157)"%(d.get("largestTable","App log"),d.get("baldatSizeMB","?")),"impact":"Performance + capacity freeing"})
        if d.get("hanaVersion"):
            tech_debt.append({"area":"HANA DB","priority":"info","effort":"Low","item":"SAP HANA %s - verify ECS-supported release via SAP Note 2196476"%d["hanaVersion"],"impact":"ECS compliance / supportability"})
        d["techDebt"] = tech_debt
        rec_notes = [{"note":"2196476","title":"Standard DB Software Releases in SAP ECS","area":"Database","priority":"info"}]
        if d.get("debugUsers") and d["debugUsers"]>0: rec_notes.append({"note":"863362","title":"S_DEVELOP DEBUG authorization","area":"Security","priority":"crit"})
        if d.get("defaultPasswords"):
            rec_notes.append({"note":"1749142","title":"SAP* user: standard password","area":"Security","priority":"crit"})
            rec_notes.append({"note":"1414256","title":"TMSADM user: standard password","area":"Security","priority":"crit"})
        if d.get("hanaAlert59"): rec_notes.append({"note":"2081856","title":"HANA Alert 59 - Blocked transactions","area":"HANA","priority":"high"})
        if (d.get("abapDumps") or 0)>20: rec_notes.append({"note":"195157","title":"Reorganize and delete the application log","area":"Operations","priority":"med"})
        if (d.get("spAgeMo") or 0)>=18: rec_notes.append({"note":"2348574","title":"SAP S/4HANA upgrades FAQ","area":"Upgrade","priority":"med"})
        if d.get("kernelVersion"): rec_notes.append({"note":"2083594","title":"Maintenance of SAP kernel packages","area":"Basis","priority":"med"})
        d["recommendedNotes"] = rec_notes
        d["_filename"] = os.path.basename(self._path)
        return d


def main():
    parser = argparse.ArgumentParser(description="Parse SAP EWA .doc/.docx -> JSON")
    parser.add_argument("--input","-i",required=True)
    parser.add_argument("--output","-o",required=True)
    parser.add_argument("--indent","-n",type=int,default=None)
    args = parser.parse_args()
    if not os.path.exists(args.input):
        print("ERROR: File not found: "+args.input,file=__import__("sys").stderr); __import__("sys").exit(1)
    ep = EWAParser(args.input)
    data = ep.parse()
    import sys
    print("[parse] SID=%s | Product=%s"%(data.get("sid"),data.get("product")),file=sys.stderr)
    print("[parse] HANA=%s | Period=%s"%(data.get("hanaVersion"),data.get("period")),file=sys.stderr)
    print("[parse] hanaPhysRam=%s | hanaMemUsed=%s | hanaAllocLimit=%s"%(data.get("hanaPhysRam"),data.get("hanaMemUsed"),data.get("hanaAllocLimit")),file=sys.stderr)
    print("[parse] dataVol: total=%s used=%s free=%s (%s%%)"%(data.get("dataVolTotal"),data.get("dataVolUsed"),data.get("dataVolFree"),data.get("dataFreePct")),file=sys.stderr)
    print("[parse] logVol: total=%s used=%s free=%s (%s%%)"%(data.get("logVolTotal"),data.get("logVolUsed"),data.get("logVolFree"),data.get("logFreePct")),file=sys.stderr)
    import json
    with open(args.output,"w",encoding="utf-8") as f:
        json.dump(data,f,indent=args.indent,ensure_ascii=False)
    print("Parsing complete: "+args.output)


if __name__ == "__main__":
    main()
