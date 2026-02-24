"""
Question Bank Parser - Extracts questions from .docx files
Handles: SA, MCQ, Part-B | Images embedded | Multiple format detection
"""
import zipfile
import xml.etree.ElementTree as ET
import re
import base64
import os
from dataclasses import dataclass, field
from typing import Optional

# ─── Namespaces ───────────────────────────────────────────
W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
A   = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
VML = 'urn:schemas-microsoft-com:vml'

UNIT_NAMES = {
    '1': 'Sequential Circuit Design',
    '2': 'Asynchronous Sequential Circuit Design',
    '3': 'Testing',
    '4': 'PLD Design',
    '5': 'Digital Systems & Programming Tools',
}

CO_UNIT_MAP = {'CO1':'1','CO2':'2','CO3':'3','CO4':'4','CO5':'5'}

@dataclass
class Question:
    id:        str = ''
    type:      str = 'SA'        # SA | MCQ | PARTB
    unit:      str = '1'
    co:        str = 'CO1'
    btl:       str = 'K2'
    marks:     int = 2
    qno:       str = ''
    text:      str = ''
    options:   list = field(default_factory=list)   # [A,B,C,D]
    correct:   str  = ''
    images:    list = field(default_factory=list)   # list of base64 data-URI strings
    source:    str  = 'aqa'      # ct | es | aqa
    filename:  str  = ''
    subject:   str  = ''

# ─── Helpers ──────────────────────────────────────────────
def _tag(ns, name): return f'{{{ns}}}{name}'

def cell_paragraphs(tc):
    """Return list of (text, has_drawing, image_rids) per paragraph in cell"""
    result = []
    for p in tc.findall(f'.//{{{W}}}p'):
        # collect text
        runs = p.findall(f'.//{{{W}}}r')
        text = ''
        for r in runs:
            for t in r.findall(f'{{{W}}}t'):
                text += (t.text or '')
        # collect image rIds
        rids = []
        for blip in p.findall(f'.//{{{A}}}blip'):
            for k,v in blip.attrib.items():
                if 'embed' in k: rids.append(v)
        for imgdata in p.findall(f'.//{{{VML}}}imagedata'):
            rid = imgdata.get(f'{{{R}}}id','') or imgdata.get(f'{{{R}}}href','')
            if rid: rids.append(rid)
        has_draw = p.find(f'.//{{{W}}}drawing') is not None or bool(rids)
        result.append({'text': text.strip(), 'has_draw': has_draw, 'rids': rids})
    return result

def cell_full_text(tc):
    paras = cell_paragraphs(tc)
    return '\n'.join(p['text'] for p in paras if p['text'])

def cell_has_image(tc):
    return (tc.find(f'.//{{{W}}}drawing') is not None or
            tc.find(f'.//{{{A}}}blip')    is not None or
            tc.find(f'.//{{{VML}}}imagedata') is not None)

def extract_co_btl(text):
    """Extract CO and K-level from text like 'CO1, K-3' or 'CO2,K2'"""
    co_m  = re.search(r'CO\s*([1-5])', text, re.I)
    btl_m = re.search(r'K[-–]?\s*([1-6])', text, re.I)
    co  = f'CO{co_m.group(1)}' if co_m else None
    btl = f'K{btl_m.group(1)}' if btl_m else None
    return co, btl

def parse_mcq_options(paras):
    """
    Given list of paragraph dicts, separate question text from options.
    Handles:
      - 4 separate paragraphs (one per option) after question para
      - Inline  a) ... b) ... c) ... d) ...
      - Single line  1 b) 2 c) 3 d) 4
    Returns (question_text, [optA, optB, optC, optD])
    """
    if not paras:
        return '', []

    q_text = paras[0]['text']
    rest   = [p['text'] for p in paras[1:] if p['text']]

    # Case 1: 4 separate option paras  → strip any leading a) b) c) d)
    if len(rest) == 4:
        opts = []
        for opt in rest:
            clean = re.sub(r'^[a-dA-D][.)]\s*', '', opt).strip()
            opts.append(clean)
        return q_text, opts

    # Case 2: Options are 3 separate paras (sometimes D is inline with C)
    if len(rest) == 3:
        opts = [re.sub(r'^[a-dA-D][.)]\s*', '', o).strip() for o in rest]
        return q_text, opts + ['']

    # Case 3: Inline  a) X  b) Y  c) Z  d) W
    combined = ' '.join([q_text] + rest)
    inline = re.split(r'\b[a-dA-D]\s*[.)]\s*', combined)
    if len(inline) >= 5:
        return inline[0].strip(), [x.strip() for x in inline[1:5]]

    # Case 4: numeric inline  "1  b) 2  c) 3  d) 4"
    if rest:
        inline2 = re.split(r'\b[b-dB-D]\s*[.)]\s*', rest[0])
        if len(inline2) >= 3:
            optA = rest[0].split(inline2[0])[0].strip() if inline2[0] else ''
            # fallback – grab first token as A
            tokens = rest[0].split()
            return q_text, [tokens[0] if tokens else ''] + [x.strip() for x in inline2[1:4]] + ['']

    return q_text, []

def is_section_header(text):
    """Detect section dividers, not real questions"""
    patterns = [
        r'^Part\s*[A-C]\b', r'^\(OR\)$', r'^Answer ALL', r'^PART\s*[–-]',
        r'^Unit\s*-?\s*[1-5]\b', r'^CO,?\s*BTL', r'^Q\.\s*No', r'^UNIT-\d',
        r'^AQA\s*[–-]', r'^Syllabus:', r'^Hyperlinks', r'^Short Qu', r'^Multiple Ch',
        r'^PART\s*–\s*[AB]$', r'^Questions$', r'^Mark$', r'^Marks$',
        r'^(I|II|III|M|E|D)$',   # Portion/level codes
    ]
    t = text.strip()
    for pat in patterns:
        if re.match(pat, t, re.I): return True
    if len(t) < 3: return True
    return False

def is_marks_cell(text):
    return bool(re.match(r'^\s*\d{1,2}\s*$', text.strip()))

def detect_type(marks, paras, co_text=''):
    """Decide SA / MCQ / PARTB"""
    if marks >= 8: return 'PARTB'
    # check if options exist
    all_text = ' '.join(p['text'] for p in paras)
    q_text   = paras[0]['text'] if paras else ''

    # 4 separate option paras
    rest = [p['text'] for p in paras[1:] if p['text']]
    if len(rest) >= 3 and all(len(r) > 0 for r in rest[:3]):
        # likely options
        return 'MCQ'

    # Inline a) b) c) d)
    if re.search(r'\b[a-d]\)', all_text, re.I) and re.search(r'\b[b-d]\)', all_text, re.I):
        return 'MCQ'

    # "Justify Your Answer" marker from ES paper
    if 'Justify Your Answer' in q_text or 'Justify the Answer' in q_text:
        if len(rest) >= 3:
            return 'MCQ'

    return 'SA'

# ─── Image extractor ──────────────────────────────────────
def build_rel_map(zf):
    """rId -> filename (just basename, e.g. 'image1.png')"""
    try:
        with zf.open('word/_rels/document.xml.rels') as f:
            rel_root = ET.parse(f).getroot()
    except KeyError:
        return {}
    rel_map = {}
    for rel in rel_root.findall(f'{{{REL}}}Relationship'):
        rid    = rel.get('Id','')
        target = rel.get('Target','')
        name   = target.split('/')[-1]
        rel_map[rid] = name
    return rel_map

def build_image_map(zf):
    """filename -> data URI base64"""
    img_map = {}
    for path in zf.namelist():
        if not path.startswith('word/media/'): continue
        ext = path.rsplit('.',1)[-1].lower()
        if ext == 'emf': continue   # skip vector graphics for now
        mime = {'png':'image/png','jpg':'image/jpeg','jpeg':'image/jpeg',
                'gif':'image/gif','bmp':'image/bmp','webp':'image/webp'}.get(ext)
        if not mime: continue
        with zf.open(path) as f:
            b64 = base64.b64encode(f.read()).decode()
        img_map[path.split('/')[-1]] = f'data:{mime};base64,{b64}'
    return img_map

def rids_to_images(rids, rel_map, img_map):
    imgs = []
    for rid in rids:
        fname = rel_map.get(rid,'')
        if fname and fname in img_map:
            imgs.append(img_map[fname])
    return imgs

def collect_cell_images(tc, rel_map, img_map):
    """All images embedded in a table cell"""
    rids = []
    for blip in tc.findall(f'.//{{{A}}}blip'):
        for k,v in blip.attrib.items():
            if 'embed' in k: rids.append(v)
    for imgdata in tc.findall(f'.//{{{VML}}}imagedata'):
        rid = imgdata.get(f'{{{R}}}id','') or ''
        if rid: rids.append(rid)
    return rids_to_images(rids, rel_map, img_map)

# ─── Subject detector ─────────────────────────────────────

def extract_co_statements(text_lines):
    """
    Extract CO statements from document text lines.
    Handles patterns like:
      CO1 / Construct the synchronous sequential circuits.
      CO2: Solve hazards...
      "CO Statements" header followed by CO+text pairs
    Returns dict: {'CO1': 'statement...', 'CO2': '...', ...}
    """
    co_stmts = {}
    n = len(text_lines)
    i = 0
    while i < n:
        line = text_lines[i].strip()
        # Detect a bare CO label line (e.g. "CO1", "CO 2")
        bare_co = re.match(r'^CO\s*([1-5])\s*$', line, re.I)
        if bare_co:
            co_key = f'CO{bare_co.group(1)}'
            # Next non-empty line is the statement
            j = i + 1
            while j < n and not text_lines[j].strip():
                j += 1
            if j < n:
                stmt = text_lines[j].strip()
                # Make sure it's not another CO label or a header
                if not re.match(r'^CO\s*[1-5]\s*$', stmt, re.I) and len(stmt) > 5:
                    co_stmts[co_key] = stmt
                    i = j + 1
                    continue
        # Detect inline "CO1: statement" or "CO1 – statement"
        inline = re.match(r'^CO\s*([1-5])\s*[:\u2013\-]\s*(.+)', line, re.I)
        if inline:
            co_key = f'CO{inline.group(1)}'
            co_stmts[co_key] = inline.group(2).strip()
        i += 1
    return co_stmts


def extract_unit_topics(text_lines):
    """
    Extract unit topic/title lines from document text.
    Handles patterns like:
      "Syllabus: Unit -1 Sequential Circuit Design"
      "Unit -1: Sequential Circuit Design"
      "UNIT 1 – Sequential Circuit Design"
    Returns dict: {'1': 'Sequential Circuit Design', '2': '...', ...}
    """
    topics = {}
    for line in text_lines:
        # Pattern: "Syllabus: Unit -1 Sequential Circuit Design"
        m = re.match(r'Syllabus:\s*Unit\s*[-–]?\s*([1-5])\s+(.+)', line, re.I)
        if m:
            topics[m.group(1)] = m.group(2).strip()
            continue
        # Pattern: "Unit -1: Title" or "Unit–1 – Title" or "UNIT 1 | Title"
        m = re.match(r'Unit\s*[-–]?\s*([1-5])\s*[:\u2013\-|]\s*(.+)', line, re.I)
        if m:
            topics[m.group(1)] = m.group(2).strip()
            continue
        # Pattern: "UNIT-1" alone then next line might be the title (handled separately)
    return topics


def detect_subject(zf):
    """Scan first few paragraphs for subject/course info, CO statements, and unit topics"""
    try:
        with zf.open('word/document.xml') as f:
            root = ET.parse(f).getroot()
    except Exception:
        return {
            'subject': 'Unknown Subject', 'code': '', 'dept': '', 'semester': '',
            'co_statements': {}, 'unit_topics': {}
        }

    info = {
        'subject': '', 'code': '', 'dept': '', 'semester': '',
        'co_statements': {}, 'unit_topics': {}
    }
    all_paras = root.findall(f'.//{{{W}}}p')
    text_lines = []
    for p in all_paras[:150]:
        t = ''.join(x.text or '' for x in p.findall(f'.//{{{W}}}t')).strip()
        if t: text_lines.append(t)

    for line in text_lines:
        if re.search(r'VEC\d{3}', line):
            info['code'] = re.search(r'VEC\d{3}', line).group(0)
        if re.search(r'Digital VLSI', line, re.I):
            info['subject'] = 'Digital VLSI Design and Technology'
        if re.search(r'Electronics and Communication', line, re.I):
            info['dept'] = 'ECE'
        if re.search(r'(fifth|third|fourth|sixth)\s+semester', line, re.I):
            info['semester'] = re.search(r'(fifth|third|fourth|sixth)\s+semester', line, re.I).group(0).title()
        if re.search(r'Semester:\s*(\S+)', line, re.I):
            info['semester'] = re.search(r'Semester:\s*(.+)', line, re.I).group(1).strip()

    if not info['subject']:
        # Generic fallback from any subject-like line
        for line in text_lines:
            if len(line) > 20 and re.search(r'(Design|Technology|Circuit|System|Signal|Network)', line, re.I):
                if not re.search(r'(Examination|Assessment|Semester|Department|regulation)', line, re.I):
                    info['subject'] = line[:80]
                    break

    if not info['subject']:
        info['subject'] = 'Unknown Subject'

    # Extract CO statements and unit topics
    info['co_statements'] = extract_co_statements(text_lines)
    info['unit_topics']   = extract_unit_topics(text_lines)

    return info

# ─── Main Table Parser ────────────────────────────────────
def parse_table(tbl, rel_map, img_map, source, subject_info):
    """Parse a single docx table → list[Question]"""
    questions = []
    rows = tbl.findall(f'{{{W}}}tr')

    # Detect column layout from header row
    # Standard layout: [CO/BTL | Q.No | Question | Marks]  (4 cols)
    # AQA layout:      [Q.No | CO | K-Level | Question | Marks | Portion | Level] (7 cols)
    # We'll detect by scanning the first 3 rows

    layout = None
    # Scan ALL rows to detect layout (tables may have info rows before question rows)
    for row in rows:
        cells = row.findall(f'{{{W}}}tc')
        texts = [cell_full_text(c).lower() for c in cells]
        joined = ' '.join(texts)
        # Header row detection
        if 'question' in joined and ('co' in joined or 'btl' in joined or 'knowledge' in joined):
            if len(cells) == 4:
                layout = '4col'; break
            elif len(cells) >= 6:
                layout = '7col'; break
        # Data row detection: CO/BTL in first cell, question in third
        if len(cells) == 4:
            t0 = cell_full_text(cells[0])
            t2 = cell_full_text(cells[2]) if len(cells) > 2 else ''
            if re.search(r'CO\s*[1-5]', t0) and len(t2) > 10:
                layout = '4col'; break
        elif len(cells) >= 6:
            t1 = cell_full_text(cells[1])
            t3 = cell_full_text(cells[3]) if len(cells) > 3 else ''
            if re.search(r'CO\s*[1-5]', t1) and len(t3) > 10:
                layout = '7col'; break

    if layout is None:
        return questions

    # Track current unit from section headers
    current_unit = '1'
    current_section = 'SA'   # SA | MCQ | PARTB
    q_counter = 1

    for row in rows:
        cells = row.findall(f'{{{W}}}tc')
        if not cells: continue

        # ── Single-cell section header rows ──
        if len(cells) == 1:
            full = cell_full_text(cells[0])
            um = re.search(r'UNIT\s*[-–]?\s*([1-5])', full, re.I)
            if um: current_unit = um.group(1)
            if re.search(r'PART\s*[–-]\s*B', full, re.I): current_section = 'PARTB'
            elif re.search(r'MCQ|Multiple Choice', full, re.I): current_section = 'MCQ'
            elif re.search(r'PART\s*[–-]\s*A.*Short', full, re.I): current_section = 'SA'
            continue

        # ── Two-cell info rows ──
        if len(cells) == 2:
            t0 = cell_full_text(cells[0])
            t1 = cell_full_text(cells[1])
            if 'Part A' in t0 or 'Part B' in t0 or 'Part C' in t0:
                if 'B' in t0: current_section = 'PARTB'
                elif 'A' in t0: current_section = 'SA'
            continue

        # ── Data rows ──
        if layout == '4col' and len(cells) >= 4:
            co_cell   = cells[0]
            qno_cell  = cells[1]
            q_cell    = cells[2]
            mrk_cell  = cells[3]

            co_txt  = cell_full_text(co_cell).strip()
            qno_txt = cell_full_text(qno_cell).strip()
            mrk_txt = cell_full_text(mrk_cell).strip()

            co, btl = extract_co_btl(co_txt)
            if not co: continue
            # detect unit from CO
            unit = CO_UNIT_MAP.get(co, current_unit)

            try: marks = int(re.search(r'\d+', mrk_txt).group())
            except: marks = 2

            paras = cell_paragraphs(q_cell)
            q_text_first = paras[0]['text'] if paras else ''
            if is_section_header(q_text_first): continue
            if not q_text_first: continue

            # Images from q_cell
            imgs = collect_cell_images(q_cell, rel_map, img_map)
            # Also images from qno_cell (some papers put image there)
            imgs += collect_cell_images(qno_cell, rel_map, img_map)

            # Detect section from surrounding context
            um = re.search(r'UNIT\s*[-–]?\s*([1-5])', co_txt, re.I)
            if um: current_unit = um.group(1)

            # Classify
            q_type = detect_type(marks, paras, co_txt)
            if q_type == 'MCQ':
                q_text, opts = parse_mcq_options(paras)
            else:
                q_text = '\n'.join(p['text'] for p in paras if p['text'])
                opts = []

            q = Question(
                id=f'{source}_{q_counter}',
                type=q_type, unit=unit, co=co, btl=btl or 'K2',
                marks=marks, qno=qno_txt or str(q_counter),
                text=q_text.strip(), options=opts, correct='',
                images=imgs, source=source, filename='',
                subject=subject_info.get('subject','')
            )
            questions.append(q)
            q_counter += 1

        elif layout == '7col' and len(cells) >= 7:
            qno_cell = cells[0]
            co_cell  = cells[1]
            k_cell   = cells[2]
            q_cell   = cells[3]
            mrk_cell = cells[4]

            qno_txt = cell_full_text(qno_cell).strip()
            co_txt  = cell_full_text(co_cell).strip()
            k_txt   = cell_full_text(k_cell).strip()
            mrk_txt = cell_full_text(mrk_cell).strip()

            # Section tracking
            if re.search(r'PART\s*[–-]\s*B', cell_full_text(q_cell), re.I):
                current_section = 'PARTB'; continue
            if re.search(r'MCQ|Multiple Choice', cell_full_text(q_cell), re.I):
                current_section = 'MCQ'; continue

            co,  _    = extract_co_btl(co_txt)
            btl, _    = extract_co_btl(f'K{k_txt.strip()}')
            if not btl:
                km = re.search(r'K\s*([1-6])', k_txt, re.I)
                btl = f'K{km.group(1)}' if km else 'K2'
            if not co: continue

            unit = CO_UNIT_MAP.get(co, current_unit)

            try: marks = int(re.search(r'\d+', mrk_txt).group())
            except: marks = 2

            paras = cell_paragraphs(q_cell)
            q_text_first = paras[0]['text'] if paras else ''
            if is_section_header(q_text_first): continue
            if not q_text_first: continue

            imgs = collect_cell_images(q_cell, rel_map, img_map)

            # In 7col layout, use current_section to help classify
            if current_section == 'MCQ' or detect_type(marks, paras) == 'MCQ':
                q_text, opts = parse_mcq_options(paras)
                q_type = 'MCQ'
            elif marks >= 8 or current_section == 'PARTB':
                q_text = '\n'.join(p['text'] for p in paras if p['text'])
                opts = []
                q_type = 'PARTB'
            else:
                q_text = '\n'.join(p['text'] for p in paras if p['text'])
                opts = []
                q_type = 'SA'

            q = Question(
                id=f'{source}_{q_counter}',
                type=q_type, unit=unit, co=co, btl=btl,
                marks=marks, qno=qno_txt or str(q_counter),
                text=q_text.strip(), options=opts, correct='',
                images=imgs, source=source, filename='',
                subject=subject_info.get('subject','')
            )
            questions.append(q)
            q_counter += 1

    return questions

# ─── Entry Point ─────────────────────────────────────────
def parse_docx(filepath, source='aqa'):
    """Parse a .docx file → list[Question]"""
    questions = []
    try:
        with zipfile.ZipFile(filepath, 'r') as zf:
            subject_info = detect_subject(zf)
            rel_map  = build_rel_map(zf)
            img_map  = build_image_map(zf)

            with zf.open('word/document.xml') as f:
                root = ET.parse(f).getroot()

            tables = root.findall(f'.//{{{W}}}tbl')
            for tbl in tables:
                qs = parse_table(tbl, rel_map, img_map, source, subject_info)
                questions.extend(qs)

    except Exception as e:
        print(f'[ERROR] {filepath}: {e}')
        import traceback; traceback.print_exc()

    # Tag filename
    basename = os.path.basename(filepath)
    for q in questions:
        q.filename = basename

    # Deduplicate by (co, text[:50])
    seen = set()
    unique = []
    for q in questions:
        key = (q.co, q.text[:50].strip().lower())
        if key not in seen:
            seen.add(key)
            unique.append(q)

    # Ensure subject_info always has co_statements and unit_topics keys
    subject_info.setdefault('co_statements', {})
    subject_info.setdefault('unit_topics', {})

    return unique, subject_info

if __name__ == '__main__':
    # Quick test
    import sys, json
    path = sys.argv[1] if len(sys.argv) > 1 else '/mnt/user-data/uploads/VEC311_AQA_.docx'
    src  = sys.argv[2] if len(sys.argv) > 2 else 'aqa'
    qs, info = parse_docx(path, src)
    print(f"Subject: {info}")
    print(f"Total questions: {len(qs)}")
    by_type = {}
    for q in qs:
        by_type.setdefault(q.type, 0)
        by_type[q.type] += 1
    print(f"By type: {by_type}")
    for q in qs[:5]:
        print(f"\n[{q.type}] {q.co}/{q.btl} U{q.unit} {q.marks}m")
        print(f"  Q: {q.text[:100]}")
        if q.options: print(f"  Opts: {q.options}")
        if q.images:  print(f"  Images: {len(q.images)}")
