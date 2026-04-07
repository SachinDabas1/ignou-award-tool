"""
IGNOU Award/Grade List document generator.
Produces a Word (.docx) exactly matching the official IGNOU form format.
"""
import io
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─────────────────────────────────────────────────────────────
# Low-level XML helpers  (work on raw <w:tc> elements)
# ─────────────────────────────────────────────────────────────

def _tc_is_continuation(tc):
    """True if tc is a vMerge continuation cell (bare <w:vMerge/> = no val attr)."""
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        return False
    vm = tcPr.find(qn('w:vMerge'))
    return vm is not None and vm.get(qn('w:val')) is None


def _get_or_create_tcPr(tc):
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    return tcPr


def apply_borders(tc, sz='4', color='000000'):
    """Add single-line borders to a tc. Skip continuation cells."""
    if _tc_is_continuation(tc):
        return
    tcPr = _get_or_create_tcPr(tc)
    for ex in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(ex)
    tcB = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        t = OxmlElement(f'w:{edge}')
        t.set(qn('w:val'),   'single')
        t.set(qn('w:sz'),    sz)
        t.set(qn('w:space'), '0')
        t.set(qn('w:color'), color)
        tcB.append(t)
    # Schema order: tcW → gridSpan → vMerge → tcBorders → shd → vAlign → hideMark
    # Insert before shd if present, else before vAlign, else append
    shd = tcPr.find(qn('w:shd'))
    va  = tcPr.find(qn('w:vAlign'))
    if shd is not None:
        shd.addprevious(tcB)
    elif va is not None:
        va.addprevious(tcB)
    else:
        tcPr.append(tcB)


def apply_bg(tc, fill):
    """Add cell background shading. Skip continuation cells."""
    if _tc_is_continuation(tc):
        return
    tcPr = _get_or_create_tcPr(tc)
    for ex in tcPr.findall(qn('w:shd')):
        tcPr.remove(ex)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  fill)
    # Insert before vAlign (correct: shd → vAlign)
    va = tcPr.find(qn('w:vAlign'))
    if va is not None:
        va.addprevious(shd)
    else:
        tcPr.append(shd)


def set_valign(tc, val='center'):
    """Set vertical alignment on tc. Skip continuation cells."""
    if _tc_is_continuation(tc):
        return
    tcPr = _get_or_create_tcPr(tc)
    for ex in tcPr.findall(qn('w:vAlign')):
        tcPr.remove(ex)
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), val)
    tcPr.append(va)   # vAlign is last before hideMark


def set_cell_text(tc, text, bold=False, size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER):
    """Write text into a tc's first paragraph. Skip continuation cells."""
    if _tc_is_continuation(tc):
        return
    set_valign(tc)
    # Find or create paragraph
    p_el = tc.find(qn('w:p'))
    if p_el is None:
        p_el = OxmlElement('w:p')
        tc.append(p_el)
    # Clear existing runs
    for r in p_el.findall(qn('w:r')):
        p_el.remove(r)
    # ── pPr: spacing BEFORE jc (schema order) ──
    pPr = p_el.find(qn('w:pPr'))
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        p_el.insert(0, pPr)
    for tag in (qn('w:spacing'), qn('w:jc')):
        for ex in pPr.findall(tag):
            pPr.remove(ex)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), '20')
    sp.set(qn('w:after'),  '20')
    pPr.append(sp)                      # spacing first
    jc_map = {
        WD_ALIGN_PARAGRAPH.CENTER: 'center',
        WD_ALIGN_PARAGRAPH.LEFT:   'left',
        WD_ALIGN_PARAGRAPH.RIGHT:  'right',
    }
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), jc_map.get(align, 'center'))
    pPr.append(jc)                      # jc after spacing
    # ── run ──
    r_el = OxmlElement('w:r')
    rPr  = OxmlElement('w:rPr')
    if bold:
        rPr.append(OxmlElement('w:b'))
    fnt = OxmlElement('w:rFonts')
    fnt.set(qn('w:ascii'), 'Times New Roman')
    fnt.set(qn('w:hAnsi'), 'Times New Roman')
    rPr.append(fnt)
    sz_el = OxmlElement('w:sz')
    sz_el.set(qn('w:val'), str(int(size * 2)))
    rPr.append(sz_el)
    r_el.append(rPr)
    t_el = OxmlElement('w:t')
    t_el.text = text
    if text != text.strip():
        t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r_el.append(t_el)
    p_el.append(r_el)


def set_col_width(tc, dxa):
    """Forcibly set tcW on a tc."""
    tcPr = _get_or_create_tcPr(tc)
    for ex in tcPr.findall(qn('w:tcW')):
        tcPr.remove(ex)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'),    str(dxa))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.insert(0, tcW)


# ─────────────────────────────────────────────────────────────
# Main generator
# ─────────────────────────────────────────────────────────────

def generate_award_list(candidates, course_code):
    """
    Generate the IGNOU Award/Grade List for Assignments as a .docx BytesIO.

    candidates : list of dicts with keys: enrollment, name, programme
    course_code: string shown in Course Code field
    """
    doc = Document()

    # ── Page: A4, 2 cm margins ──
    sec = doc.sections[0]
    sec.page_width    = Cm(21)
    sec.page_height   = Cm(29.7)
    sec.top_margin    = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin   = Cm(2.0)
    sec.right_margin  = Cm(2.0)
    CW = 9638  # content width in DXA = (21-4) cm × 567

    # ─── paragraph helpers ───
    def para(text='', bold=False, size=11, align=WD_ALIGN_PARAGRAPH.CENTER,
             sb=0, sa=2, underline=False):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(sb)
        p.paragraph_format.space_after  = Pt(sa)
        if text:
            r = p.add_run(text)
            r.bold = bold; r.underline = underline
            r.font.size = Pt(size); r.font.name = 'Times New Roman'
        return p

    def info_line(lbl1, val1, lbl2, val2, sa=2):
        """Two-column info line using a tab stop."""
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(sa)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(3.35))
        for txt, ul in [(lbl1, False), (val1, True), ('\t', False), (lbl2, False), (val2, True)]:
            r = p.add_run(txt)
            r.underline = ul
            r.font.size = Pt(9)
            r.font.name = 'Times New Roman'

    # ── HEADER ──
    para('INDIRA GANDHI NATIONAL OPEN UNIVERSITY', bold=True, size=13, sa=1)
    para('Study Centre & Code : Naval Headquarters, DNE, RK Puram , New Delhi (Code -7101)',
         size=9, sa=2)
    para('AWARD/GRADE LIST FOR ASSIGNMENTS', bold=True, size=11, underline=True, sa=1)
    para('For TEE Dec 2024 Session', size=9, sa=4)

    info_line('Programme: ', 'x', 'Course Code: ', course_code)
    info_line('Study Centre: ', '7101, NAVY', 'Assignment No: ', '_' * 10)
    info_line('Place: ', 'WEST BLOK - V, PORTA CABIN, SECTOR-1, RK PURAM',
              '  For MA Maximum Marks: ', '100', sa=4)
    para('Please arrange Enrolment Nos. in ascending order only and write', size=9, sa=0)
    para('Complete and correct enrolment number in nine digits.', size=9, sa=4)

    # ── TABLE ──
    # 8 logical columns: Sl.No | Enrolment | Name | Programme | TMA-I | TMA-II | TMA-III | Remarks
    raw_w = [520, 1480, 2480, 1440, 860, 860, 860, 1138]
    scale = CW / sum(raw_w)
    COL_W = [int(w * scale) for w in raw_w]
    COL_W[-1] = CW - sum(COL_W[:-1])   # absorb rounding

    data_rows  = max(25, len(candidates))
    total_rows = 2 + data_rows          # 2 header rows

    tbl = doc.add_table(rows=total_rows, cols=8)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths on every cell BEFORE any merges
    for row in tbl.rows:
        for ci, cell in enumerate(row.cells):
            set_col_width(cell._tc, COL_W[ci])

    # ── MERGES (do all merges before setting any text) ──
    # Row 0: cols 4-6 → "Grade/Award" span
    tbl.rows[0].cells[4].merge(tbl.rows[0].cells[6])
    # Rows 0-1: cols 0,1,2,3,7 → vertical span
    for col in [0, 1, 2, 3, 7]:
        tbl.rows[0].cells[col].merge(tbl.rows[1].cells[col])

    # ── Access RAW XML tc elements (not python-docx proxies) ──
    # Row 0 XML has 6 actual tc: [Sl, Enrol, Name, Prog, Grade/Award(gs=3), Remarks]
    # Row 1 XML has 8 actual tc: [cont×4, TMA-I, TMA-II, TMA-III, cont]
    rows_xml = tbl._tbl.findall(qn('w:tr'))
    r0_tcs   = rows_xml[0].findall(qn('w:tc'))
    r1_tcs   = rows_xml[1].findall(qn('w:tc'))

    HBG   = 'C0CCE0'   # light blue-grey header bg
    HSIZE = 8

    # Row 0 headers
    r0_data = [
        (r0_tcs[0], 'Sl.\nNo',            WD_ALIGN_PARAGRAPH.CENTER),
        (r0_tcs[1], 'Enrolment No.',       WD_ALIGN_PARAGRAPH.CENTER),
        (r0_tcs[2], 'Name of Candidate',   WD_ALIGN_PARAGRAPH.CENTER),
        (r0_tcs[3], 'Programme',           WD_ALIGN_PARAGRAPH.CENTER),
        (r0_tcs[4], 'Grade/Award',         WD_ALIGN_PARAGRAPH.CENTER),
        (r0_tcs[5], 'Remarks\nif any',     WD_ALIGN_PARAGRAPH.CENTER),
    ]
    for tc, text, align in r0_data:
        set_cell_text(tc, text, bold=True, size=HSIZE, align=align)
        apply_borders(tc, sz='6')
        apply_bg(tc, HBG)

    # Row 1: only the 3 TMA cells (indices 4, 5, 6 in raw XML) get content
    for tc, text in [(r1_tcs[4], 'TMA-I'), (r1_tcs[5], 'TMA-II'), (r1_tcs[6], 'TMA-III')]:
        set_cell_text(tc, text, bold=True, size=HSIZE)
        apply_borders(tc, sz='6')
        apply_bg(tc, HBG)
    # Row 1 continuation cells (indices 0,1,2,3,7) are left as-is (only tcW allowed)

    # ── DATA ROWS ──
    aligns = [
        WD_ALIGN_PARAGRAPH.CENTER,  # Sl.No
        WD_ALIGN_PARAGRAPH.LEFT,    # Enrolment
        WD_ALIGN_PARAGRAPH.LEFT,    # Name
        WD_ALIGN_PARAGRAPH.LEFT,    # Programme
        WD_ALIGN_PARAGRAPH.CENTER,  # TMA-I
        WD_ALIGN_PARAGRAPH.CENTER,  # TMA-II
        WD_ALIGN_PARAGRAPH.CENTER,  # TMA-III
        WD_ALIGN_PARAGRAPH.CENTER,  # Remarks
    ]

    for i in range(data_rows):
        row_xml = rows_xml[2 + i]
        # Row height
        trPr = row_xml.find(qn('w:trPr'))
        if trPr is None:
            trPr = OxmlElement('w:trPr')
            row_xml.insert(0, trPr)
        for ex in trPr.findall(qn('w:trHeight')):
            trPr.remove(ex)
        trH = OxmlElement('w:trHeight')
        trH.set(qn('w:val'), '240')   # ~12pt
        trPr.append(trH)

        row_tcs = row_xml.findall(qn('w:tc'))
        if i < len(candidates):
            c    = candidates[i]
            vals = [str(i + 1), c['enrollment'], c['name'], c['programme'],
                    '', '', '', '']
        else:
            vals = [str(i + 1), '', '', '', '', '', '', '']

        for j, (tc, val) in enumerate(zip(row_tcs, vals)):
            set_cell_text(tc, val, size=9, align=aligns[j])
            apply_borders(tc, sz='4')

    # ── SIGNATURE BLOCK ──
    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(10)
    p_gap.paragraph_format.space_after  = Pt(0)

    sig_lines = [
        ('Signature of Co-ordinator________________',
         'Signature of Evaluation________________'),
        ('Date____________________',
         'Date____________________'),
        ('Office Stamp',
         'Name & Address_________________________________'),
    ]
    for left, right in sig_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after  = Pt(5)
        p.paragraph_format.tab_stops.add_tab_stop(Inches(3.35))
        for txt in [left, '\t', right]:
            r = p.add_run(txt)
            r.font.size = Pt(9)
            r.font.name = 'Times New Roman'

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf
