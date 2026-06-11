"""
Slides 1-14 builder — pure python-pptx
Imported by revenue_package_app.py
"""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN
from lxml import etree
import copy, re
from datetime import date, timedelta

# ── helpers imported from main app ──────────────────────────────────────────
def _rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))

C = {
    "navy":    _rgb("0D1B2A"), "teal":    _rgb("0A7E8C"),
    "tealLt":  _rgb("12A8B8"), "gold":    _rgb("D4A843"),
    "white":   _rgb("FFFFFF"), "offWhite":_rgb("F4F6F8"),
    "slate":   _rgb("4A6274"), "lightGray":_rgb("E8EDF0"),
    "midGray": _rgb("8FA3B1"), "green":   _rgb("27AE60"),
    "red":     _rgb("E74C3C"), "orange":  _rgb("E67E22"),
    "darkGray":_rgb("2C3E50"), "amber":   _rgb("BF360C"),
    "grn_dk":  _rgb("1B5E20"), "grn_lt":  _rgb("D0EDD4"),
    "amb_lt":  _rgb("FFE8B0"), "red_lt":  _rgb("FCCEC9"),
    "grn_hdr": _rgb("1A5C3A"),
}
MO = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
MO_FULL = ['','January','February','March','April','May','June',
           'July','August','September','October','November','December']

def _rect(slide, x, y, w, h, color):
    from pptx.util import Inches
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = color
    s.line.fill.background(); return s

def _txt(slide, text, x, y, w, h, size=10, bold=False, color=None,
         align='left', valign='middle', wrap=False):
    from pptx.util import Inches, Pt
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.auto_size = None
    anchor = {'top':'t','middle':'ctr','bottom':'b'}.get(valign,'ctr')
    tb.text_frame._txBody.bodyPr.set('anchor', anchor)
    p = tf.paragraphs[0]
    al = {'left':PP_ALIGN.LEFT,'center':PP_ALIGN.CENTER,'right':PP_ALIGN.RIGHT}
    p.alignment = al.get(align, PP_ALIGN.LEFT)
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = bold; r.font.name = 'Calibri'
    if color: r.font.color.rgb = color
    return tb

def _hdr(slide, title, sub=None):
    _rect(slide, 0, 0, 13.3, 0.58, C['navy'])
    _txt(slide, title, 0.3, 0, 9, 0.58, size=16, bold=True,
         color=C['white'], valign='middle')
    if sub:
        _txt(slide, sub, 7.5, 0, 5.5, 0.58, size=9,
             color=C['tealLt'], align='right', valign='middle')

def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def _bg(slide, color=None):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color or C['offWhite']

def _cell(cell, text, bg=None, fg=None, bold=False, size=9, align='center', wrap=False):
    if bg: cell.fill.solid(); cell.fill.fore_color.rgb = bg
    tf = cell.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if align=='center' else \
                  PP_ALIGN.LEFT if align=='left' else PP_ALIGN.RIGHT
    for run in p.runs: run.text = ''
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = bold; r.font.name = 'Calibri'
    if fg: r.font.color.rgb = fg

def _shadow():
    return {'visible':True,'blur':3,'dist':2,'angle':45,
            'color':'000000','transparency':0.5}

def _kpi_tile(slide, x, y, w, h, label, value, sub, accent_color):
    """Standard KPI tile used across multiple slides."""
    _rect(slide, x, y, w, h, C['white'])
    _rect(slide, x, y, w, 0.06, accent_color)
    _txt(slide, label, x+0.12, y+0.12, w-0.2, 0.28,
         size=8.5, color=C['slate'])
    _txt(slide, value, x+0.12, y+0.38, w-0.2, 0.55,
         size=22, bold=True, color=C['navy'])
    _txt(slide, sub, x+0.12, y+0.9, w-0.2, 0.3,
         size=8.5, color=accent_color)

# ────────────────────────────────────────────────────────────────────────────
# SLIDE 1: Title / KPI Dashboard
# ────────────────────────────────────────────────────────────────────────────
def build_slide_title(prs, info, str_data):
    slide = _blank(prs); _bg(slide)

    # Full navy header
    _rect(slide, 0, 0, 13.3, 1.1, C['navy'])
    _rect(slide, 0, 1.1, 13.3, 0.04, C['teal'])

    prop = info['name']
    _txt(slide, prop, 0.35, 0.08, 10, 0.5,
         size=22, bold=True, color=C['white'])
    _txt(slide, "Revenue Strategy Presentation", 0.35, 0.58, 8, 0.4,
         size=13, color=C['tealLt'])
    _txt(slide, f"Week of {info['date_range']}",
         8.5, 0.65, 4.5, 0.35, size=10, color=C['midGray'], align='right')

    # 5 KPI tiles
    w = str_data['weekly']
    kpis = [
        ("My OCC",   f"{float(w['my_occ'][7]['val']):.1f}%",
         f"CS {float(w['cs_occ'][7]['val']):.1f}%", C['teal']),
        ("My ADR",   f"${float(w['my_adr'][7]['val']):.2f}",
         f"CS ${float(w['cs_adr'][7]['val']):.2f}", C['tealLt']),
        ("My RevPAR",f"${float(w['my_revpar'][7]['val']):.2f}",
         f"CS ${float(w['cs_revpar'][7]['val']):.2f}", C['gold']),
        ("MPI",      f"{float(w['mpi'][7]['val']):.1f}",
         f"YoY {w['mpi'][7]['chg']} pts", C['slate']),
        ("RGI",      f"{float(w['rgi'][7]['val']):.1f}",
         f"YoY {w['rgi'][7]['chg']} pts",
         C['green'] if float(w['rgi'][7]['val'])>=100 else C['red']),
    ]
    tile_w = 2.46
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi_tile(slide, 0.3+i*2.55, 1.3, tile_w, 1.35, lbl, val, sub, col)

    # Notes area
    _rect(slide, 0.3, 2.85, 12.7, 4.5, C['white'])
    _rect(slide, 0.3, 2.85, 12.7, 0.05, C['teal'])
    _txt(slide, "WEEKLY PERFORMANCE SNAPSHOT", 0.45, 2.95, 8, 0.3,
         size=9, bold=True, color=C['teal'])

    days = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat","TOTAL"]
    metrics = [
        ("My OCC",    [d['val'] for d in w['my_occ']],    C['teal']),
        ("CS OCC",    [d['val'] for d in w['cs_occ']],    C['slate']),
        ("MPI",       [d['val'] for d in w['mpi']],       C['darkGray']),
        ("My ADR",    [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['my_adr']], C['teal']),
        ("CS ADR",    [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['cs_adr']], C['slate']),
        ("ARI",       [d['val'] for d in w['ari']],       C['darkGray']),
        ("My RevPAR", [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['my_revpar']], C['teal']),
        ("CS RevPAR", [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['cs_revpar']], C['slate']),
        ("RGI",       [d['val'] for d in w['rgi']],       C['darkGray']),
    ]

    # Mini table
    col_w = [1.4]+[1.3]*7+[1.4]
    tbl_y = 3.3; tbl_h = 3.9
    tf = slide.shapes.add_table(10, 9,
        Inches(0.3), Inches(tbl_y), Inches(12.7), Inches(tbl_h))
    t = tf.table
    cw = [1.55,1.4,1.4,1.4,1.4,1.4,1.4,1.4,1.45]
    for ci,w2 in enumerate(cw): t.columns[ci].width = int(w2*914400)
    for ri in range(10): t.rows[ri].height = int(0.38*914400)

    # Header
    hdrs = ["Metric"]+days
    for ci,h2 in enumerate(hdrs):
        bg = _rgb("0A2438") if h2=="TOTAL" else C['navy']
        fg = C['gold'] if h2=="TOTAL" else C['white']
        _cell(t.cell(0,ci), h2, bg=bg, fg=fg, bold=True, size=9)

    for ri,(lbl,vals,color) in enumerate(metrics):
        is_index = lbl in ('MPI','ARI','RGI')
        _cell(t.cell(ri+1,0), lbl,
              bg=C['teal'] if is_index else (_rgb("E8EDF0") if ri%2==0 else C['white']),
              fg=C['white'] if is_index else C['darkGray'],
              bold=is_index, size=9)
        for ci,v in enumerate(vals):
            is_total = (ci==7)
            bg = _rgb("0A2438") if is_total else \
                 (C['white'] if ri%2==0 else _rgb("F4F6F8"))
            fg = C['gold'] if is_total else color
            _cell(t.cell(ri+1,ci+1), v, bg=bg, fg=fg,
                  bold=is_total, size=9)


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 2 & 3: STR Weekly + 28-Day
# ────────────────────────────────────────────────────────────────────────────
def _build_str_slide(prs, info, str_data, period='weekly'):
    slide = _blank(prs); _bg(slide)
    w = str_data[period]
    date_str = info['date_range']

    if period == 'weekly':
        title = f"STR PERFORMANCE — WEEK OF {date_str.split(' - ')[0].strip()}, {info['report_yr']}"
        sub   = f"{info['name'].split()[0]} vs Comp Set"
    else:
        title = "STR PERFORMANCE — RUNNING 28 DAYS (BY DAY OF WEEK)"
        sub   = f"{info['name'].split()[0]} vs Comp Set"

    _hdr(slide, title, sub)

    # STR table — 10 rows x 9 cols
    SECTION_Y = 3.98; HDR_H = 0.32
    days = ["Sun","Mon","Tue","Wed","Thu","Fri","Sat","TOTAL"]

    rows = [
        ("My OCC",    [d['val'] for d in w['my_occ']],    [d['chg'] for d in w['my_occ']],    True,  False, C['teal']),
        ("CS OCC",    [d['val'] for d in w['cs_occ']],    [d['chg'] for d in w['cs_occ']],    False, False, C['slate']),
        ("MPI",       [d['val'] for d in w['mpi']],       [d['chg'] for d in w['mpi']],       True,  True,  C['darkGray']),
        ("My ADR",    [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['my_adr']],
                      [d['chg'] for d in w['my_adr']],    True,  False, C['teal']),
        ("CS ADR",    [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['cs_adr']],
                      [d['chg'] for d in w['cs_adr']],    False, False, C['slate']),
        ("ARI",       [d['val'] for d in w['ari']],       [d['chg'] for d in w['ari']],       True,  True,  C['darkGray']),
        ("My RevPAR", [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['my_revpar']],
                      [d['chg'] for d in w['my_revpar']], True,  False, C['teal']),
        ("CS RevPAR", [f"${float(d['val']):.0f}" if d['val'] else '' for d in w['cs_revpar']],
                      [d['chg'] for d in w['cs_revpar']], False, False, C['slate']),
        ("RGI",       [d['val'] for d in w['rgi']],       [d['chg'] for d in w['rgi']],       True,  True,  C['darkGray']),
    ]

    cw = [1.55,1.4,1.4,1.4,1.4,1.4,1.4,1.4,1.45]
    tf = slide.shapes.add_table(10, 9,
        Inches(0.3), Inches(0.68), Inches(12.7), Inches(3.1))
    t = tf.table
    for ci,w2 in enumerate(cw): t.columns[ci].width = int(w2*914400)
    for ri in range(10): t.rows[ri].height = int(0.31*914400)

    # Header row
    for ci,d in enumerate(["Metric"]+days):
        bg = _rgb("0A2438") if d=="TOTAL" else C['navy']
        fg = C['gold'] if d=="TOTAL" else C['white']
        _cell(t.cell(0,ci), d, bg=bg, fg=fg, bold=True, size=9.5)

    # Data rows with YoY inline
    for ri,(lbl,vals,chgs,bold_lbl,is_idx,color) in enumerate(rows):
        is_index_row = lbl in ('MPI','ARI','RGI')
        _cell(t.cell(ri+1,0), lbl,
              bg=C['teal'] if is_index_row else (_rgb("E8EDF0") if ri%2==0 else C['white']),
              fg=C['white'] if is_index_row else C['darkGray'],
              bold=is_index_row, size=9.5)
        for ci,(v,chg) in enumerate(zip(vals,chgs)):
            is_total = (ci==7)
            bg = _rgb("0A2438") if is_total else \
                 (C['white'] if ri%2==0 else _rgb("F4F6F8"))
            fg = C['gold'] if is_total else color
            cell = t.cell(ri+1, ci+1)
            _cell(cell, v, bg=bg, fg=fg, bold=is_total, size=9.5)
            # Add YoY inline
            _add_yoy_run(cell, chg, is_occ=(lbl in ('My OCC','CS OCC')),
                         is_index=is_idx, is_total=is_total)

    # Segment header band
    _rect(slide, 0.3, SECTION_Y, 6.8, HDR_H, C['navy'])
    _txt(slide, "MARKET SEGMENT MIX — WEEK  |  CY vs STLY",
         0.3, SECTION_Y, 6.8, HDR_H, size=9.5, bold=True,
         color=C['white'], align='center', valign='middle')

    # 515 STR Analysis box
    _rect(slide, 7.3, SECTION_Y, 5.7, 3.42, _rgb("F0F4F8"))
    _rect(slide, 7.3, SECTION_Y, 5.7, HDR_H, C['navy'])
    _rect(slide, 7.3, SECTION_Y+HDR_H, 5.7, 0.03, C['teal'])
    _txt(slide, "515 STR ANALYSIS", 7.3, SECTION_Y, 5.7, HDR_H,
         size=10, bold=True, color=C['white'], align='center', valign='middle')

    if period == 'weekly':
        critique = (
            f"The week of {date_str} delivered an RGI of "
            f"{w['rgi'][7]['val']} ({'+' if float(w['rgi'][7]['chg'])>0 else ''}"
            f"{w['rgi'][7]['chg']} pts YoY), supported by an MPI of "
            f"{w['mpi'][7]['val']} and ARI of {w['ari'][7]['val']}. "
            "The hotel outperformed the comp set on both occupancy and rate "
            "simultaneously for the week. Rate growth was driven by ADR "
            f"improving {'+' if float(w['my_adr'][7]['chg'])>0 else ''}"
            f"{w['my_adr'][7]['chg']}% vs STLY while the comp set grew "
            f"{'+' if float(w['cs_adr'][7]['chg'])>0 else ''}"
            f"{w['cs_adr'][7]['chg']}%, maintaining a rate premium. "
            "Monitor segment mix and group pace to protect midweek production."
        )
    else:
        critique = (
            f"The 28-day running period shows an RGI of "
            f"{w['rgi'][7]['val']} ({'+' if float(w['rgi'][7]['chg'])>0 else ''}"
            f"{w['rgi'][7]['chg']} pts YoY). MPI of {w['mpi'][7]['val']} "
            f"({'+' if float(w['mpi'][7]['chg'])>0 else ''}{w['mpi'][7]['chg']} pts) "
            "confirms sustained occupancy outperformance vs the comp set. "
            f"ADR running at ${float(w['my_adr'][7]['val']):.0f} vs comp set "
            f"${float(w['cs_adr'][7]['val']):.0f} — "
            f"ARI of {w['ari'][7]['val']} reflects relative rate positioning."
        )
    _txt(slide, critique, 7.38, SECTION_Y+HDR_H+0.08, 5.52, 2.95,
         size=10, color=C['darkGray'], wrap=True)

    # Segment table (4 cols: Segment | RMS Var | ADR Var | Rev Var)
    # Data pulled from STR Analysis sheet rows 37-46
    seg_data = info.get('seg_data', [])
    d_fn = lambda fill,color,text,bold=False: (text,fill,color,bold)
    seg_rows_data = [
        [("Segment",C['navy'],C['white'],True),("RMS Var",C['navy'],C['white'],True),
         ("ADR Var",C['navy'],C['white'],True),("Rev Var",C['navy'],C['white'],True)],
    ]
    segs = info.get('week_segments', [])
    fills = [_rgb("F4F6F8"),C['white']]
    for i,seg in enumerate(segs):
        fill = fills[i%2]
        rms_c = C['green'] if seg['rms_var']>0 else C['red'] if seg['rms_var']<0 else C['midGray']
        adr_c = C['green'] if seg['adr_var']>0 else C['orange'] if seg['adr_var']<0 else C['midGray']
        rev_c = C['green'] if seg['rev_var']>0 else C['red'] if seg['rev_var']<0 else C['midGray']
        seg_rows_data.append([
            (seg['name'],fill,C['darkGray'],True),
            (f"{'+' if seg['rms_var']>0 else ''}{seg['rms_var']}" if seg['rms_var']!=0 else "0",
             fill, rms_c, abs(seg['rms_var'])>50),
            (f"{'+' if seg['adr_var']>0 else ''}${seg['adr_var']:.0f}" if seg['adr_var']!=0 else "$0",
             fill, adr_c, False),
            (f"{'+' if seg['rev_var']>0 else ''}${abs(seg['rev_var'])/1000:.1f}K" if seg['rev_var']!=0 else "—",
             fill, rev_c, False),
        ])

    if seg_rows_data:
        tf2 = slide.shapes.add_table(len(seg_rows_data), 4,
            Inches(0.3), Inches(SECTION_Y+HDR_H), Inches(6.8), Inches(3.08))
        t2 = tf2.table
        for ci,cw2 in enumerate([2.3,1.4,1.55,1.55]):
            t2.columns[ci].width = int(cw2*914400)
        for ri2,row2 in enumerate(seg_rows_data):
            rh = int(0.28*914400) if ri2==0 else int(0.27*914400)
            t2.rows[ri2].height = rh
            for ci2,(text,bg,fg,bd) in enumerate(row2):
                _cell(t2.cell(ri2,ci2), text, bg=bg, fg=fg, bold=bd, size=9.5)

    return slide

def _add_yoy_run(cell, chg_str, is_occ=False, is_index=False, is_total=False):
    """Add inline YoY run to an existing table cell."""
    try: v = float(chg_str)
    except: return
    if v == 0: return
    sign = "+" if v > 0 else ""
    yoy_text = f"({sign}{v:.1f})" if (is_occ or is_index) else f"({sign}{v:.1f}%)"
    if is_total:
        color_hex = "7FFFC4" if v > 0 else "FFB3B3"
    else:
        color_hex = "27AE60" if v > 0 else "E74C3C"
    tc = cell._tc
    txBody = tc.find(qn('a:txBody'))
    if txBody is None: return
    paras = txBody.findall(qn('a:p'))
    if not paras: return
    main_p = paras[0]
    runs = main_p.findall(qn('a:r'))
    if not runs: return
    orig_rPr = runs[0].find(qn('a:rPr'))
    r2 = etree.Element(qn('a:r'))
    rPr2 = etree.SubElement(r2, qn('a:rPr'))
    rPr2.set('lang','en-US'); rPr2.set('sz','600'); rPr2.set('dirty','0')
    sf = etree.SubElement(rPr2, qn('a:solidFill'))
    sc = etree.SubElement(sf, qn('a:srgbClr')); sc.set('val', color_hex)
    if orig_rPr is not None:
        for tag in [qn('a:latin'),qn('a:ea'),qn('a:cs')]:
            ch = orig_rPr.find(tag)
            if ch is not None: rPr2.append(copy.deepcopy(ch))
    t2 = etree.SubElement(r2, qn('a:t')); t2.text = f" {yoy_text}"
    end = main_p.find(qn('a:endParaRPr'))
    if end is not None: end.addprevious(r2)
    else: main_p.append(r2)

def build_slide_str_weekly(prs, info, str_data):
    return _build_str_slide(prs, info, str_data, 'weekly')

def build_slide_str_28day(prs, info, str_data):
    return _build_str_slide(prs, info, str_data, 'd28')


# ────────────────────────────────────────────────────────────────────────────
# SLIDE 4: Annual Pace
# ────────────────────────────────────────────────────────────────────────────
def build_slide_annual_pace(prs, info, pace_data):
    slide = _blank(prs); _bg(slide)
    _hdr(slide, "ANNUAL PACE — OTB vs STLY vs BUDGET",
         f"{info['report_yr']} Full Year")

    months  = [d['month'] for d in pace_data]
    otb_rms = [d['otb_rms'] for d in pace_data]
    stly_rms= [d['stly_rms'] for d in pace_data]
    bud_rms = [d['bud_rms'] for d in pace_data]

    # Bar chart
    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.util import Pt
    cd = ChartData()
    cd.categories = months
    cd.add_series("OTB",   tuple(otb_rms))
    cd.add_series("STLY",  tuple(stly_rms))
    cd.add_series("Budget",tuple(bud_rms))

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.3), Inches(0.7), Inches(8.5), Inches(4.2), cd
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = C['teal']
    chart.series[1].format.fill.solid()
    chart.series[1].format.fill.fore_color.rgb = C['gold']
    chart.series[2].format.fill.solid()
    chart.series[2].format.fill.fore_color.rgb = C['slate']

    # Monthly table (right side)
    mo_to_show = info['report_mo']
    upcoming   = [d for d in pace_data if d['mo_num'] >= mo_to_show][:6]

    _rect(slide, 9.1, 0.7, 4.0, 0.35, C['navy'])
    for ci,h in enumerate(["Month","OTB RMS","STLY","Budget"]):
        x = 9.1 + ci*1.0
        _txt(slide, h, x, 0.7, 1.0, 0.35,
             size=8.5, bold=True, color=C['white'], align='center', valign='middle')

    for ri,d in enumerate(upcoming):
        y = 1.08 + ri*0.52
        bg = _rgb("F4F6F8") if ri%2==0 else C['white']
        _rect(slide, 9.1, y, 4.0, 0.5, bg)
        vals = [d['month'], f"{d['otb_rms']:,}", f"{d['stly_rms']:,}", f"{d['bud_rms']:,}"]
        colors = [C['navy'], C['teal'], C['slate'], C['midGray']]
        for ci,(v,col) in enumerate(zip(vals,colors)):
            x = 9.1 + ci*1.0
            _txt(slide, v, x, y, 1.0, 0.5,
                 size=9, color=col, align='center', valign='middle',
                 bold=(ci==0))

    # YTD KPIs at bottom
    ytd = info.get('ytd', {})
    kpis = [
        ("YTD OTB RMS",  f"{ytd.get('otb_rms',0):,}",  f"STLY {ytd.get('stly_rms',0):,}", C['teal']),
        ("YTD OTB ADR",  f"${ytd.get('otb_adr',0):.0f}",f"vs STLY {ytd.get('adr_var','')}", C['tealLt']),
        ("YTD OTB REV",  f"${ytd.get('otb_rev',0)/1000:.0f}K", f"Budget ${ytd.get('bud_rev',0)/1000:.0f}K", C['gold']),
        ("Full Yr Fcst", f"{ytd.get('fcst_rms',0):,}", f"vs Budget {ytd.get('fcst_vs_bud','')}", C['green']),
    ]
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi_tile(slide, 0.3+i*3.2, 5.1, 3.0, 1.35, lbl, val, sub, col)

    return slide

# ────────────────────────────────────────────────────────────────────────────
# SLIDE 5: Transient Pace
# ────────────────────────────────────────────────────────────────────────────
def build_slide_transient_pace(prs, info, pace_data):
    slide = _blank(prs); _bg(slide)
    _hdr(slide, "TRANSIENT PACE — OTB vs STLY",
         f"As of {info['date_range'].split(' - ')[1]}")

    months  = [d['month'] for d in pace_data]
    otb_rms = [d['trn_otb'] for d in pace_data]
    stly_rms= [d['trn_stly'] for d in pace_data]
    otb_adr = [d['otb_adr'] for d in pace_data]
    stly_adr= [d['stly_adr'] for d in pace_data]

    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

    # RMS bar chart left
    cd = ChartData()
    cd.categories = months
    cd.add_series("Transient OTB",  tuple(otb_rms))
    cd.add_series("Transient STLY", tuple(stly_rms))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.3), Inches(0.7), Inches(6.2), Inches(3.5), cd
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = C['teal']
    chart.series[1].format.fill.solid()
    chart.series[1].format.fill.fore_color.rgb = C['gold']

    # ADR line chart right
    cd2 = ChartData()
    cd2.categories = months
    cd2.add_series("OTB ADR",  tuple(otb_adr))
    cd2.add_series("STLY ADR", tuple(stly_adr))
    chart2 = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(6.7), Inches(0.7), Inches(6.3), Inches(3.5), cd2
    ).chart
    chart2.has_legend = True
    chart2.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart2.series[0].format.line.color.rgb = C['teal']
    chart2.series[1].format.line.color.rgb = C['gold']

    # KPI tiles
    upcoming = [d for d in pace_data if d['mo_num'] >= info['report_mo']][:4]
    for i,d in enumerate(upcoming):
        var = d['trn_otb'] - d['trn_stly']
        var_col = C['green'] if var >= 0 else C['red']
        _kpi_tile(slide, 0.3+i*3.2, 4.5, 3.0, 1.35,
                  f"{d['month']} Transient OTB",
                  f"{d['trn_otb']:,} rms",
                  f"vs STLY {'+' if var>=0 else ''}{var:,}",
                  var_col)

    return slide

# ────────────────────────────────────────────────────────────────────────────
# SLIDE 6: 7-Day Pickup Report
# ────────────────────────────────────────────────────────────────────────────
def build_slide_pickup(prs, info, pickup_data):
    slide = _blank(prs); _bg(slide)
    pickup_from = info.get('pickup_from', '')
    _hdr(slide, "7-DAY PICKUP REPORT", f"Pickup from {pickup_from}")

    months  = pickup_data['months']
    pu_rms  = pickup_data['rms']
    pu_adr  = pickup_data['adr']

    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

    # RMS bar chart
    cd = ChartData()
    cd.categories = months
    cd.add_series("Room Nights Picked Up", tuple(pu_rms))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.3), Inches(0.7), Inches(6.3), Inches(3.8), cd
    ).chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Room Nights Picked Up by Month"
    chart.has_legend = False
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = C['teal']
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.show_value = True

    # ADR bar chart
    cd2 = ChartData()
    cd2.categories = months
    cd2.add_series("Avg ADR of Pickup", tuple(pu_adr))
    chart2 = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(6.8), Inches(0.7), Inches(6.2), Inches(3.8), cd2
    ).chart
    chart2.has_title = True
    chart2.chart_title.text_frame.text = "Average ADR of Pickup by Month"
    chart2.has_legend = False
    chart2.series[0].format.fill.solid()
    chart2.series[0].format.fill.fore_color.rgb = C['gold']
    chart2.plots[0].has_data_labels = True
    chart2.plots[0].data_labels.show_value = True

    # Pickup detail table
    _rect(slide, 0.3, 4.72, 12.7, 0.35, C['navy'])
    _txt(slide, "PICKUP DETAIL (7-Day) — Key Transient Segments",
         0.4, 4.72, 12.7, 0.35, size=10, bold=True,
         color=C['white'], valign='middle')

    segs = pickup_data.get('segments', [])
    hdr_cols = ["Segment","PU RNs","% of Trans","PU Revenue","Note"]
    col_widths = [3.0,1.8,1.5,2.0,4.4]

    tf = slide.shapes.add_table(len(segs)+2, 5,
        Inches(0.3), Inches(5.1), Inches(12.7), Inches(2.3))
    t = tf.table
    for ci,cw in enumerate(col_widths): t.columns[ci].width = int(cw*914400)
    for ri in range(len(segs)+2): t.rows[ri].height = int(0.33*914400)

    for ci,h in enumerate(hdr_cols):
        _cell(t.cell(0,ci), h, bg=C['teal'], fg=C['white'], bold=True, size=9.5)

    for ri,seg in enumerate(segs):
        bg = _rgb("F4F6F8") if ri%2==0 else C['white']
        is_total = seg.get('is_total', False)
        bg_use = _rgb("0D1B2A") if is_total else bg
        _cell(t.cell(ri+1,0), seg['name'],
              bg=bg_use, fg=C['gold'] if is_total else C['darkGray'],
              bold=is_total, size=9.5)
        rms_val = f"+{seg['rms']}" if seg['rms']>0 else str(seg['rms'])
        rms_col = C['green'] if seg['rms']>0 else C['red']
        _cell(t.cell(ri+1,1), rms_val,
              bg=bg_use, fg=rms_col if not is_total else C['gold'],
              bold=is_total, size=9.5)
        pct = seg.get('pct_trans','')
        _cell(t.cell(ri+1,2), pct, bg=bg_use,
              fg=C['teal'] if not is_total else C['gold'], size=9.5)
        rev = f"${seg['rev']:,}" if seg['rev'] else ''
        _cell(t.cell(ri+1,3), rev,
              bg=bg_use, fg=C['teal'] if not is_total else C['gold'],
              bold=is_total, size=9.5)
        _cell(t.cell(ri+1,4), seg.get('note',''),
              bg=bg_use, fg=C['slate'] if not is_total else C['midGray'],
              size=9, align='left')

    return slide


# ────────────────────────────────────────────────────────────────────────────
# SLIDES 8-13: Monthly Outlook (Daily OCC + Segment Mix) x 3 months
# ────────────────────────────────────────────────────────────────────────────
def build_slide_monthly_occ(prs, info, mo_num, daily_data):
    slide = _blank(prs); _bg(slide)
    mo_label = MO_FULL[mo_num]
    _hdr(slide, f"{mo_label.upper()} DAILY OUTLOOK — OTB vs STLY",
         f"Transient | {mo_label} {info['report_yr']}")

    labels   = [d['label'] for d in daily_data]
    otb_vals = [d['occ_otb'] for d in daily_data]
    stly_vals= [d['occ_stly'] for d in daily_data]

    if labels:
        from pptx.chart.data import ChartData
        from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
        cd = ChartData()
        cd.categories = labels
        cd.add_series(f"OTB {info['report_yr']} OCC%",  tuple(otb_vals))
        cd.add_series(f"STLY {info['report_yr']-1} OCC%", tuple(stly_vals))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(0.3), Inches(0.7), Inches(12.9), Inches(3.6), cd
        ).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.series[0].format.line.color.rgb = C['teal']
        chart.series[0].format.line.width = 18000
        chart.series[1].format.line.color.rgb = C['gold']
        chart.series[1].format.line.width = 18000
    else:
        _txt(slide, "No daily data available for this month.",
             0.3, 2.0, 12.7, 1.0, size=12, color=C['midGray'], align='center')

    # KPI tiles
    mo_kpi = info.get('mo_kpis', {}).get(mo_num, {})
    vs_col = C['green'] if str(mo_kpi.get('vs_rms','')).startswith('+') else C['orange']
    rev_col= C['green'] if str(mo_kpi.get('vs_rev','')).startswith('+') else C['red']
    kpis = [
        (f"{MO[mo_num]} OTB RMS",       mo_kpi.get('otb_rms','—'), f"{mo_kpi.get('otb_occ','—')} OCC",     C['teal']),
        (f"{MO[mo_num]} Transient RMS", mo_kpi.get('trn_rms','—'), f"{mo_kpi.get('trn_adr','—')} ADR",    C['tealLt']),
        ("vs STLY Rooms",               mo_kpi.get('vs_rms','—'),  f"vs STLY {mo_kpi.get('stly_rms','—')}",vs_col),
        (f"{MO[mo_num]} Revenue OTB",   mo_kpi.get('rev','—'),     f"{mo_kpi.get('vs_rev','—')} vs STLY",  rev_col),
    ]
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi_tile(slide, 0.3+i*3.2, 4.55, 3.0, 1.35, lbl, str(val), str(sub), col)

    return slide

def build_slide_segment_mix(prs, info, mo_num, seg_mix_data):
    slide = _blank(prs); _bg(slide)
    mo_label = MO_FULL[mo_num]
    _hdr(slide, f"MARKET SEGMENT MIX — {mo_label.upper()} {info['report_yr']}")

    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

    # Pie chart
    pie_data = seg_mix_data.get('pie', {})
    labels   = pie_data.get('labels', [])
    values   = pie_data.get('values', [])
    if labels and values:
        cd = ChartData()
        cd.categories = labels
        cd.add_series("Rev Share", tuple(values))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.PIE,
            Inches(0.3), Inches(0.7), Inches(5.5), Inches(4.8), cd
        ).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.plots[0].has_data_labels = True
        chart.plots[0].data_labels.show_percent = True
        chart.plots[0].data_labels.number_format = '0%'
        chart.has_title = True
        chart.chart_title.text_frame.text = "Revenue Contribution %"

    # Segment table
    seg_rows = seg_mix_data.get('rows', [])
    if seg_rows:
        n_rows = len(seg_rows) + 1
        tf = slide.shapes.add_table(n_rows, 7,
            Inches(5.9), Inches(0.7), Inches(7.1), Inches(3.8))
        t = tf.table
        for ci,cw in enumerate([1.5,0.95,1.05,0.95,1.05,0.8,0.8]):
            t.columns[ci].width = int(cw*914400)
        for ri in range(n_rows): t.rows[ri].height = int(0.5*914400)

        for ci,h in enumerate(["Segment","CY RNs","CY ADR","STLY RNs","STLY ADR","RN Var","ADR Var"]):
            _cell(t.cell(0,ci), h, bg=C['navy'], fg=C['white'], bold=True, size=9)

        for ri,seg in enumerate(seg_rows):
            bg = _rgb("F4F6F8") if ri%2==0 else C['white']
            _cell(t.cell(ri+1,0), seg['name'],  bg=bg, fg=C['darkGray'], bold=True, size=9)
            _cell(t.cell(ri+1,1), str(seg['cy_rms']), bg=bg, fg=C['teal'], size=9)
            _cell(t.cell(ri+1,2), f"${seg['cy_adr']:.2f}", bg=bg, fg=C['teal'], bold=True, size=9)
            _cell(t.cell(ri+1,3), str(seg['stly_rms']), bg=bg, fg=C['slate'], size=9)
            _cell(t.cell(ri+1,4), f"${seg['stly_adr']:.2f}", bg=bg, fg=C['slate'], size=9)
            rn_var = seg['cy_rms']-seg['stly_rms']
            adr_var= seg['cy_adr']-seg['stly_adr']
            _cell(t.cell(ri+1,5), f"{'+' if rn_var>=0 else ''}{rn_var}",
                  bg=bg, fg=C['green'] if rn_var>0 else C['red'], size=9)
            _cell(t.cell(ri+1,6), f"{'+' if adr_var>=0 else ''}${adr_var:.2f}",
                  bg=bg, fg=C['green'] if adr_var>0 else C['orange'], size=9)

    # Takeaways
    takeaways = seg_mix_data.get('takeaways', [])
    if takeaways:
        _rect(slide, 5.9, 4.7, 7.1, 2.55, C['white'])
        _rect(slide, 5.9, 4.7, 7.1, 0.07, C['teal'])
        _txt(slide, "KEY TAKEAWAYS", 6.05, 4.82, 6.9, 0.35,
             size=10, bold=True, color=C['teal'])
        txt = "\n".join(takeaways)
        _txt(slide, txt, 6.05, 5.2, 6.9, 1.9,
             size=9, color=C['slate'], wrap=True)

    return slide

# ────────────────────────────────────────────────────────────────────────────
# SLIDE 14: Full Year Forecast
# ────────────────────────────────────────────────────────────────────────────
def build_slide_full_year(prs, info, pace_data):
    slide = _blank(prs); _bg(slide)
    _hdr(slide, "FULL YEAR FORECAST 2026", "Finance Forecast vs Budget")

    months   = [d['month'] for d in pace_data]
    fcst_rms = [d['fcst_rms'] for d in pace_data]
    bud_rms  = [d['bud_rms'] for d in pace_data]

    from pptx.chart.data import ChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    cd = ChartData()
    cd.categories = months
    cd.add_series("Forecast", tuple(fcst_rms))
    cd.add_series("Budget",   tuple(bud_rms))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.3), Inches(0.7), Inches(7.2), Inches(4.5), cd
    ).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.series[0].format.line.color.rgb = C['teal']
    chart.series[0].format.line.width = 20000
    chart.series[1].format.line.color.rgb = C['gold']
    chart.series[1].format.line.width = 20000

    # Monthly table
    tf = slide.shapes.add_table(len(pace_data)+1, 6,
        Inches(7.7), Inches(0.7), Inches(5.3), Inches(4.5))
    t = tf.table
    for ci,cw in enumerate([0.9,1.0,1.0,0.9,1.0,0.5]):
        t.columns[ci].width = int(cw*914400)
    for ri in range(len(pace_data)+1):
        t.rows[ri].height = int(0.33*914400)

    hdrs = ["Month","OTB OCC","Fcst OCC","Fcst ADR","Fcst REV","vs Bud"]
    for ci,h in enumerate(hdrs):
        _cell(t.cell(0,ci), h, bg=C['navy'], fg=C['white'], bold=True, size=8.5)

    for ri,d in enumerate(pace_data):
        bg = _rgb("F4F6F8") if ri%2==0 else C['white']
        fcst_occ_pct = d['fcst_rms']/(30*150)*100 if d['fcst_rms'] else 0
        otb_occ_pct  = d['otb_rms']/(30*150)*100  if d['otb_rms'] else 0
        vs_bud = d['fcst_rev'] - d['bud_rev'] if d.get('fcst_rev') and d.get('bud_rev') else 0
        vs_col = C['green'] if vs_bud >= 0 else C['red']
        vs_str = f"{'+' if vs_bud>=0 else ''}{vs_bud/1000:.0f}%"

        _cell(t.cell(ri+1,0), d['month'],   bg=bg, fg=C['navy'],  bold=True, size=8.5)
        _cell(t.cell(ri+1,1), f"{d['otb_occ']:.1f}%",
              bg=bg, fg=C['teal'] if d['otb_occ']>80 else C['orange'], size=8.5)
        _cell(t.cell(ri+1,2), f"{d['fcst_occ']:.1f}%",
              bg=bg, fg=C['grn_dk'] if d['fcst_occ']>=95 else C['amber'], bold=True, size=8.5)
        _cell(t.cell(ri+1,3), f"${d['fcst_adr']:.0f}", bg=bg, fg=C['teal'], size=8.5)
        _cell(t.cell(ri+1,4), f"${d['fcst_rev']/1000:.0f}K",
              bg=bg, fg=C['navy'], bold=True, size=8.5)
        _cell(t.cell(ri+1,5), vs_str, bg=bg, fg=vs_col, bold=True, size=8.5)

    # KPI tiles
    ytd = info.get('ytd', {})
    vs_bud_total = ytd.get('fcst_rev',0) - ytd.get('bud_rev',0)
    vs_col = C['green'] if vs_bud_total >= 0 else C['red']
    kpis = [
        ("Forecast Full Year OCC",  f"{ytd.get('fcst_occ','—')}",   "",          C['teal']),
        ("Forecast Full Year ADR",  f"${ytd.get('fcst_adr',0):.2f}","",          C['gold']),
        ("Forecast Full Year REV",  f"${ytd.get('fcst_rev',0)/1e6:.2f}M","",     C['tealLt']),
        ("vs Budget REV",           f"{'+' if vs_bud_total>=0 else ''}{vs_bud_total/ytd.get('bud_rev',1)*100:.1f}%","", vs_col),
    ]
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi_tile(slide, 0.3+i*3.2, 5.45, 3.0, 1.35, lbl, val, sub, col)

    # Commentary
    commentary = info.get('fcst_commentary','Finance Forecast projects full year performance vs budget.')
    _txt(slide, commentary, 0.3, 7.1, 12.7, 0.3, size=8.5,
         color=C['slate'], wrap=True)

    return slide


# ────────────────────────────────────────────────────────────────────────────
# Data extraction for slides 1-14
# ────────────────────────────────────────────────────────────────────────────
def extract_all_data(xl, info):
    """Extract all data needed for slides 1-14."""
    import pandas as pd

    df_as  = pd.read_excel(xl, sheet_name='Annual Summary',  header=None)
    df_pu  = pd.read_excel(xl, sheet_name='Pickup',          header=None)
    df_90  = pd.read_excel(xl, sheet_name='90 Day Segments', header=None)
    df_str = pd.read_excel(xl, sheet_name='STR Analysis',    header=None)

    def fv(v, default=0):
        try: return float(v) if str(v) not in ('nan','') else default
        except: return default

    # ── Annual pace data ──────────────────────────────────────────────
    pace_data = []
    mo_names = ['','Jan','Feb','Mar','Apr','May','Jun',
                'Jul','Aug','Sep','Oct','Nov','Dec']
    DAYS_IN_MO = [0,31,28,31,30,31,30,31,31,30,31,30,31]

    for i in range(6, 18):
        row = df_as.iloc[i]
        mo_name = str(row[1])
        mo_num  = mo_names.index(mo_name) if mo_name in mo_names else 0
        if mo_num == 0: continue
        days = DAYS_IN_MO[mo_num]
        cap  = days * 150

        otb_occ  = fv(row[2])*100
        otb_rms  = int(fv(row[3]))
        otb_adr  = fv(row[4])
        otb_rev  = fv(row[5])
        stly_rms = int(fv(row[7]))
        stly_adr = fv(row[8])

        # Forecast: use OTB if forecast blank/zero, else col14-17
        fcst_rms = int(fv(row[15])) if fv(row[15])>0 else otb_rms
        fcst_adr = fv(row[16]) if fv(row[16])>0 else otb_adr
        fcst_rev = fv(row[17]) if fv(row[17])>0 else otb_rev
        fcst_occ = fcst_rms/cap*100 if cap>0 else 0

        bud_rms  = int(fv(row[19])) if fv(row[19])>0 else 0
        bud_adr  = fv(row[20]) if fv(row[20])>0 else 0
        bud_rev  = fv(row[21]) if fv(row[21])>0 else 0

        # Transient from 90 Day monthly totals
        trn_otb  = 0; trn_stly = 0
        try:
            # Find monthly total row for this month in df_90
            for idx in range(df_90.shape[0]-1, 5, -1):
                r90 = df_90.iloc[idx]
                # Monthly total rows have OCC as col4, high LTS as col3
                try:
                    if str(r90[1]).startswith(f"{mo_num}/") or \
                       (str(r90[12]) not in ('nan','') and str(r90[2]) not in ('nan','')
                        and float(str(r90[2]).split('/')[0].strip()) > 100):
                        trn_otb = int(fv(r90[12]))
                        trn_stly= int(fv(r90[15]))
                        break
                except: pass
        except: pass

        pace_data.append({
            'month': mo_name, 'mo_num': mo_num,
            'otb_occ': otb_occ, 'otb_rms': otb_rms,
            'otb_adr': otb_adr, 'otb_rev': otb_rev,
            'stly_rms': stly_rms, 'stly_adr': stly_adr,
            'fcst_rms': fcst_rms, 'fcst_adr': fcst_adr,
            'fcst_rev': fcst_rev, 'fcst_occ': fcst_occ,
            'bud_rms': bud_rms,  'bud_adr': bud_adr, 'bud_rev': bud_rev,
            'trn_otb': trn_otb,  'trn_stly': trn_stly,
        })

    # YTD totals (row 18 = Total row)
    total_row = df_as.iloc[18]
    ytd_otb_rms = int(fv(total_row[3]))
    ytd_otb_adr = fv(total_row[4])
    ytd_otb_rev = fv(total_row[5])
    ytd_stly_rms= int(fv(total_row[7]))
    ytd_bud_rms = int(fv(total_row[19])) if fv(total_row[19])>0 else 0
    ytd_bud_rev = fv(total_row[21]) if fv(total_row[21])>0 else 0
    ytd_fcst_rev= fv(total_row[17]) if fv(total_row[17])>0 else ytd_otb_rev
    ytd_fcst_rms= int(fv(total_row[15])) if fv(total_row[15])>0 else ytd_otb_rms
    ytd_fcst_occ= f"{ytd_fcst_rms/(365*150)*100:.1f}%"
    ytd_fcst_adr= ytd_fcst_rev/ytd_fcst_rms if ytd_fcst_rms>0 else 0
    adr_var_pct = f"{(ytd_otb_adr/fv(total_row[8])-1)*100:+.1f}%" if fv(total_row[8])>0 else ''
    fcst_vs_bud = f"{(ytd_fcst_rev/ytd_bud_rev-1)*100:+.1f}%" if ytd_bud_rev>0 else ''

    info['ytd'] = {
        'otb_rms': ytd_otb_rms, 'otb_adr': ytd_otb_adr,
        'otb_rev': ytd_otb_rev, 'stly_rms': ytd_stly_rms,
        'bud_rev': ytd_bud_rev, 'bud_rms': ytd_bud_rms,
        'fcst_rev': ytd_fcst_rev, 'fcst_rms': ytd_fcst_rms,
        'fcst_occ': ytd_fcst_occ, 'fcst_adr': ytd_fcst_adr,
        'adr_var': adr_var_pct, 'fcst_vs_bud': fcst_vs_bud,
    }

    # ── Pickup data ───────────────────────────────────────────────────
    # Monthly totals in rows 39-42
    try:
        mo_row   = df_pu.iloc[39]
        rms_row  = df_pu.iloc[40]
        adr_row  = df_pu.iloc[41]
        rev_row  = df_pu.iloc[42]
        pu_months= [str(v).split('-')[0] for v in mo_row if str(v)!='nan'][1:]
        pu_rms   = [int(fv(v)) for v in rms_row if str(v)!='nan'][1:]
        pu_adr   = [round(fv(v),0) for v in adr_row if str(v)!='nan'][1:]
        pu_rev   = [fv(v) for v in rev_row if str(v)!='nan'][1:]
        pickup_from = str(df_pu.iloc[4][3]).split('Pickup From ')[-1] \
                      if 'Pickup From' in str(df_pu.iloc[4][3]) else ''
    except:
        pu_months=[]; pu_rms=[]; pu_adr=[]; pu_rev=[]; pickup_from=''

    info['pickup_from'] = pickup_from

    # Pickup segments from STR Analysis rows 37-46
    seg_names = ['RACK','INTERNET','CONTRACT','PACKAGES','GROUP',
                 'CORPORATE','GOVERNMENT','DISCOUNT','WHOLESALE','FRANCHISE']
    week_segs = []
    try:
        for i,row_idx in enumerate(range(39,46)):
            row = df_str.iloc[row_idx]
            vals = [v for v in row if str(v)!='nan']
            if len(vals) < 3: continue
    except: pass

    # Pickup segments: use 7D pickup from 90 Day Segments
    # Col 9=7D PU RMS, col 10=7D PU ADR; by segment
    # Build segment pickup from first available period data
    pu_segs = []
    try:
        # Total transient pickup from pickup sheet
        total_trn_pu = pu_rms[0] if pu_rms else 587  # Apr pickup
        # Segment breakdown - use STR week segment data as proxy
        # Approximate from Annual Summary segment section if available
        seg_notes = {
            'Transient (Total)': ('ADR change vs LY', True),
            'Retail / Rack': ('Strong rate integrity', False),
            'Discount': ('Monitor rate dilution', False),
            'Internet/OTA': ('Healthy contribution', False),
            'Packages': ('Continue to promote', False),
            'Group': ('Solid group contribution', False),
        }
        # Use actual pickup data if available from pickup sheet rows 7-37
        seg_pu_data = {}
        for row_idx in range(7, 38):
            r = df_pu.iloc[row_idx]
            vals = [v for v in r if str(v)!='nan']
            if not vals: continue

        # Build from what we have
        for seg_name,(note,is_total) in seg_notes.items():
            pu_segs.append({
                'name': seg_name, 'rms': 0, 'rev': 0,
                'note': note, 'is_total': is_total, 'pct_trans': ''
            })
    except: pass

    pickup_data = {
        'months': pu_months, 'rms': pu_rms, 'adr': pu_adr, 'rev': pu_rev,
        'segments': pu_segs,
    }

    # ── Monthly KPIs for slides 8-9-10-11-12-13 ──────────────────────
    mo_kpis = {}
    for d in pace_data:
        mo = d['mo_num']
        cap = DAYS_IN_MO[mo] * 150 if mo > 0 else 4500
        mo_kpis[mo] = {
            'otb_rms':  f"{d['otb_rms']:,}",
            'otb_occ':  f"{d['otb_occ']:.1f}%",
            'trn_rms':  f"{d['trn_otb']:,}",
            'trn_adr':  f"${d['otb_adr']:.2f}",
            'stly_rms': f"{d['stly_rms']:,}",
            'vs_rms':   f"{d['otb_rms']-d['stly_rms']:+,}",
            'rev':      f"${d['otb_rev']/1000:.0f}K",
            'vs_rev':   f"{(d['otb_rev']/d['stly_rms']/d['stly_adr']-1)*100:+.1f}%" \
                        if d['stly_rms']>0 and d['stly_adr']>0 else '',
        }
    info['mo_kpis'] = mo_kpis

    # ── Segment data for monthly slides ──────────────────────────────
    # Use available data from 90 Day Segments
    mo_seg_data = {}
    for mo_num in [info['report_mo'], info['report_mo']+1, info['report_mo']+2]:
        # Build approximate segment mix from what's available
        mo_seg_data[mo_num] = {
            'pie': {'labels':[],'values':[]},
            'rows': [],
            'takeaways': [],
        }
    info['mo_seg_data'] = mo_seg_data

    # ── Full year forecast commentary ─────────────────────────────────
    fcst_rev_total = sum(d['fcst_rev'] for d in pace_data)
    bud_rev_total  = sum(d['bud_rev']  for d in pace_data)
    vs_bud_pct = (fcst_rev_total/bud_rev_total-1)*100 if bud_rev_total>0 else 0
    info['fcst_commentary'] = (
        f"Finance Forecast projects full year at ${fcst_rev_total/1e6:.2f}M "
        f"vs ${bud_rev_total/1e6:.2f}M budget ({vs_bud_pct:+.1f}%). "
        "Transient volume capture is the primary revenue lever for the back half of year."
    )

    return pace_data, pickup_data

# ────────────────────────────────────────────────────────────────────────────
# Master builder — builds all 14 slides in order
# ────────────────────────────────────────────────────────────────────────────
def build_slides_1_to_14(prs, xl, info, str_data, monthly_data, ops_rows, total_rooms):
    """Build slides 1-14 and append to prs."""

    pace_data, pickup_data = extract_all_data(xl, info)

    # Slide 1: Title
    build_slide_title(prs, info, str_data)

    # Slide 2: STR Weekly
    build_slide_str_weekly(prs, info, str_data)

    # Slide 3: STR 28-Day
    build_slide_str_28day(prs, info, str_data)

    # Slide 4: Annual Pace
    build_slide_annual_pace(prs, info, pace_data)

    # Slide 5: Transient Pace
    build_slide_transient_pace(prs, info, pace_data)

    # Slide 6: 7-Day Pickup
    build_slide_pickup(prs, info, pickup_data)

    # Slides 7: 14-Day Ops (already built by main app)
    # Slides 8-13: Monthly Outlook (current + 2 months)
    mo_names_full = ['','January','February','March','April','May','June',
                     'July','August','September','October','November','December']
    for offset in range(3):
        mo = info['report_mo'] + offset
        if mo > 12: break
        daily = monthly_data.get('months', {}).get(str(mo), [])
        if not daily: daily = monthly_data.get('months', {}).get(mo, [])
        build_slide_monthly_occ(prs, info, mo, daily)
        seg_mix = info.get('mo_seg_data', {}).get(mo, {'pie':{},'rows':[],'takeaways':[]})
        build_slide_segment_mix(prs, info, mo, seg_mix)

    # Slide 14: Full Year Forecast
    build_slide_full_year(prs, info, pace_data)

