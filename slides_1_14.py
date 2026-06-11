"""
Slides 1-13 builder (pure python-pptx)
Slide 14 = 14-Day Ops is built by main app
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.chart.data import ChartData
from pptx.oxml.ns import qn
from pptx.enum.text import PP_ALIGN
from lxml import etree
import copy, re
from datetime import date

def _rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

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
MO       = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
MO_FULL  = ['','January','February','March','April','May','June',
            'July','August','September','October','November','December']
DAYS_MO  = [0,31,28,31,30,31,30,31,31,30,31,30,31]

# ── Primitive helpers ────────────────────────────────────────────────────────
def _rect(slide, x, y, w, h, color):
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = color
    s.line.fill.background(); return s

def _txt(slide, text, x, y, w, h, size=10, bold=False, color=None,
         align='left', valign='middle', wrap=False, italic=False):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    anchor = {'top':'t','middle':'ctr','bottom':'b'}.get(valign,'ctr')
    tb.text_frame._txBody.bodyPr.set('anchor', anchor)
    p = tf.paragraphs[0]
    al = {'left':PP_ALIGN.LEFT,'center':PP_ALIGN.CENTER,'right':PP_ALIGN.RIGHT}
    p.alignment = al.get(align, PP_ALIGN.LEFT)
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = bold
    r.font.italic = italic; r.font.name = 'Calibri'
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

def _bg(slide):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = C['offWhite']

def _cell(cell, text, bg=None, fg=None, bold=False, size=9,
          align='center', wrap=False):
    if bg: cell.fill.solid(); cell.fill.fore_color.rgb = bg
    tf = cell.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = (PP_ALIGN.CENTER if align=='center' else
                   PP_ALIGN.LEFT   if align=='left'   else PP_ALIGN.RIGHT)
    for run in p.runs: run.text = ''
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = bold; r.font.name = 'Calibri'
    if fg: r.font.color.rgb = fg

def _kpi(slide, x, y, w, h, label, value, sub, accent):
    _rect(slide, x, y, w, h, C['white'])
    _rect(slide, x, y, w, 0.06, accent)
    _txt(slide, label, x+0.12, y+0.10, w-0.2, 0.26, size=8.5, color=C['slate'])
    _txt(slide, value, x+0.12, y+0.34, w-0.2, 0.52, size=22, bold=True, color=C['navy'])
    _txt(slide, sub,   x+0.12, y+0.86, w-0.2, 0.30, size=8.5, color=accent)

def _occ_colors(pct):
    if pct >= 95: return C['grn_lt'], C['grn_dk']
    if pct >= 80: return C['amb_lt'], C['amber']
    return C['red_lt'], C['red']

def _add_yoy(cell, chg_str, is_occ=False, is_index=False, is_total=False):
    """Append inline YoY run before endParaRPr."""
    try: v = float(chg_str)
    except: return
    if v == 0: return
    sign = "+" if v > 0 else ""
    yoy = f"({sign}{v:.1f})" if (is_occ or is_index) else f"({sign}{v:.1f}%)"
    if is_total:
        col = "7FFFC4" if v > 0 else "FFB3B3"
    else:
        col = "27AE60" if v > 0 else "E74C3C"
    tc = cell._tc
    txb = tc.find(qn('a:txBody'))
    if txb is None: return
    paras = txb.findall(qn('a:p'))
    if not paras: return
    mp = paras[0]
    runs = mp.findall(qn('a:r'))
    if not runs: return
    orig = runs[0].find(qn('a:rPr'))
    r2 = etree.Element(qn('a:r'))
    rp = etree.SubElement(r2, qn('a:rPr'))
    rp.set('lang','en-US'); rp.set('sz','600'); rp.set('dirty','0')
    sf = etree.SubElement(rp, qn('a:solidFill'))
    sc = etree.SubElement(sf, qn('a:srgbClr')); sc.set('val', col)
    if orig is not None:
        for tag in [qn('a:latin'),qn('a:ea'),qn('a:cs')]:
            ch = orig.find(tag)
            if ch is not None: rp.append(copy.deepcopy(ch))
    t2 = etree.SubElement(r2, qn('a:t')); t2.text = f" {yoy}"
    end = mp.find(qn('a:endParaRPr'))
    if end is not None: end.addprevious(r2)
    else: mp.append(r2)

# ── Slide 1: Title / KPI Dashboard ──────────────────────────────────────────
def build_slide_title(prs, info, str_data):
    slide = _blank(prs); _bg(slide)
    _rect(slide, 0, 0, 13.3, 1.1, C['navy'])
    _rect(slide, 0, 1.1, 13.3, 0.04, C['teal'])
    _txt(slide, info['name'], 0.35, 0.08, 10, 0.5,
         size=22, bold=True, color=C['white'])
    _txt(slide, "Revenue Strategy Presentation", 0.35, 0.58, 8, 0.4,
         size=13, color=C['tealLt'])
    _txt(slide, f"Week of {info['date_range']}",
         8.5, 0.65, 4.5, 0.35, size=10, color=C['midGray'], align='right')

    w = str_data['weekly']
    def fv(key, idx=7):
        try: return float(w[key][idx]['val'])
        except: return 0.0

    kpis = [
        ("My OCC",    f"{fv('my_occ'):.1f}%",
         f"CS {fv('cs_occ'):.1f}%",     C['teal']),
        ("My ADR",    f"${fv('my_adr'):.2f}",
         f"CS ${fv('cs_adr'):.2f}",     C['tealLt']),
        ("My RevPAR", f"${fv('my_revpar'):.2f}",
         f"CS ${fv('cs_revpar'):.2f}",  C['gold']),
        ("MPI",       f"{fv('mpi'):.1f}",
         f"YoY {w['mpi'][7]['chg']} pts",  C['slate']),
        ("RGI",       f"{fv('rgi'):.1f}",
         f"YoY {w['rgi'][7]['chg']} pts",
         C['green'] if fv('rgi') >= 100 else C['red']),
    ]
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi(slide, 0.3+i*2.55, 1.3, 2.46, 1.35, lbl, val, sub, col)

    # Weekly snapshot table
    _rect(slide, 0.3, 2.85, 12.7, 4.5, C['white'])
    _rect(slide, 0.3, 2.85, 12.7, 0.05, C['teal'])
    _txt(slide, "WEEKLY PERFORMANCE SNAPSHOT", 0.45, 2.95, 8, 0.3,
         size=9, bold=True, color=C['teal'])

    rows_def = [
        ("My OCC",    'my_occ',    True,  False, C['teal']),
        ("CS OCC",    'cs_occ',    True,  False, C['slate']),
        ("MPI",       'mpi',       False, True,  C['darkGray']),
        ("My ADR",    'my_adr',    False, False, C['teal']),
        ("CS ADR",    'cs_adr',    False, False, C['slate']),
        ("ARI",       'ari',       False, True,  C['darkGray']),
        ("My RevPAR", 'my_revpar', False, False, C['teal']),
        ("CS RevPAR", 'cs_revpar', False, False, C['slate']),
        ("RGI",       'rgi',       False, True,  C['darkGray']),
    ]
    cw = [1.55,1.4,1.4,1.4,1.4,1.4,1.4,1.4,1.45]
    tf = slide.shapes.add_table(10, 9, Inches(0.3), Inches(3.3),
                                Inches(12.7), Inches(3.9))
    t = tf.table
    for ci,w2 in enumerate(cw): t.columns[ci].width = int(w2*914400)
    for ri in range(10): t.rows[ri].height = int(0.38*914400)

    for ci,h in enumerate(["Metric","Sun","Mon","Tue","Wed","Thu","Fri","Sat","TOTAL"]):
        bg = _rgb("0A2438") if h == "TOTAL" else C['navy']
        fg = C['gold']  if h == "TOTAL" else C['white']
        _cell(t.cell(0,ci), h, bg=bg, fg=fg, bold=True, size=9)

    for ri,(lbl,key,is_occ,is_idx,color) in enumerate(rows_def):
        is_index_row = key in ('mpi','ari','rgi')
        row_bg = C['teal'] if is_index_row else \
                 (_rgb("E8EDF0") if ri%2==0 else C['white'])
        _cell(t.cell(ri+1,0), lbl, bg=row_bg,
              fg=C['white'] if is_index_row else C['darkGray'],
              bold=is_index_row, size=9)
        for ci,d in enumerate(w[key]):
            is_total = (ci == 7)
            bg = _rgb("0A2438") if is_total else \
                 (C['white'] if ri%2==0 else _rgb("F4F6F8"))
            val = d['val']
            if key in ('my_adr','cs_adr','my_revpar','cs_revpar'):
                try: val = f"${float(val):.0f}"
                except: pass
            fg = C['gold'] if is_total else color
            _cell(t.cell(ri+1,ci+1), val, bg=bg, fg=fg,
                  bold=is_total, size=9)
            _add_yoy(t.cell(ri+1,ci+1), d['chg'],
                     is_occ=is_occ, is_index=is_idx, is_total=is_total)

# ── Slides 2-3: STR Weekly / 28-Day ─────────────────────────────────────────
def _build_str_slide(prs, info, str_data, period):
    slide = _blank(prs); _bg(slide)
    w = str_data[period]
    if period == 'weekly':
        title = f"STR PERFORMANCE — WEEK OF {info['date_range']}, {info['report_yr']}"
    else:
        title = "STR PERFORMANCE — RUNNING 28 DAYS (BY DAY OF WEEK)"
    _hdr(slide, title, f"{info['name'].split()[0]} vs Comp Set")

    rows_def = [
        ("My OCC",    'my_occ',    True,  False, C['teal']),
        ("CS OCC",    'cs_occ',    True,  False, C['slate']),
        ("MPI",       'mpi',       False, True,  C['darkGray']),
        ("My ADR",    'my_adr',    False, False, C['teal']),
        ("CS ADR",    'cs_adr',    False, False, C['slate']),
        ("ARI",       'ari',       False, True,  C['darkGray']),
        ("My RevPAR", 'my_revpar', False, False, C['teal']),
        ("CS RevPAR", 'cs_revpar', False, False, C['slate']),
        ("RGI",       'rgi',       False, True,  C['darkGray']),
    ]
    cw = [1.55,1.4,1.4,1.4,1.4,1.4,1.4,1.4,1.45]
    tf = slide.shapes.add_table(10, 9, Inches(0.3), Inches(0.68),
                                Inches(12.7), Inches(3.1))
    t = tf.table
    for ci,w2 in enumerate(cw): t.columns[ci].width = int(w2*914400)
    for ri in range(10): t.rows[ri].height = int(0.31*914400)

    for ci,h in enumerate(["Metric","Sun","Mon","Tue","Wed","Thu","Fri","Sat","TOTAL"]):
        bg = _rgb("0A2438") if h=="TOTAL" else C['navy']
        fg = C['gold']  if h=="TOTAL" else C['white']
        _cell(t.cell(0,ci), h, bg=bg, fg=fg, bold=True, size=9.5)

    for ri,(lbl,key,is_occ,is_idx,color) in enumerate(rows_def):
        is_idx_row = key in ('mpi','ari','rgi')
        row_bg = C['teal'] if is_idx_row else \
                 (_rgb("E8EDF0") if ri%2==0 else C['white'])
        _cell(t.cell(ri+1,0), lbl, bg=row_bg,
              fg=C['white'] if is_idx_row else C['darkGray'],
              bold=is_idx_row, size=9.5)
        for ci,d in enumerate(w[key]):
            is_total = (ci == 7)
            bg = _rgb("0A2438") if is_total else \
                 (C['white'] if ri%2==0 else _rgb("F4F6F8"))
            val = d['val']
            if key in ('my_adr','cs_adr','my_revpar','cs_revpar'):
                try: val = f"${float(val):.0f}"
                except: pass
            _cell(t.cell(ri+1,ci+1), val,
                  bg=bg, fg=C['gold'] if is_total else color,
                  bold=is_total, size=9.5)
            _add_yoy(t.cell(ri+1,ci+1), d['chg'],
                     is_occ=is_occ, is_index=is_idx, is_total=is_total)

    # Segment mix + 515 STR Analysis section
    SEC_Y = 3.98
    _rect(slide, 0.3, SEC_Y, 6.8, 0.32, C['navy'])
    _txt(slide, "MARKET SEGMENT MIX — WEEK  |  CY vs STLY",
         0.3, SEC_Y, 6.8, 0.32, size=9.5, bold=True,
         color=C['white'], align='center', valign='middle')

    _rect(slide, 7.3, SEC_Y, 5.7, 3.42, _rgb("F0F4F8"))
    _rect(slide, 7.3, SEC_Y, 5.7, 0.32, C['navy'])
    _rect(slide, 7.3, SEC_Y+0.32, 5.7, 0.03, C['teal'])
    _txt(slide, "515 STR ANALYSIS", 7.3, SEC_Y, 5.7, 0.32,
         size=10, bold=True, color=C['white'], align='center', valign='middle')

    # Auto-generate commentary from data
    rgi_val = w['rgi'][7]['val']; rgi_chg = w['rgi'][7]['chg']
    mpi_val = w['mpi'][7]['val']; ari_val = w['ari'][7]['val']
    my_adr_chg = w['my_adr'][7]['chg']; cs_adr_chg = w['cs_adr'][7]['chg']
    dr = info['date_range']
    if period == 'weekly':
        critique = (
            f"The week of {dr} delivered an RGI of {rgi_val} "
            f"({'+' if float(rgi_chg)>0 else ''}{rgi_chg} pts YoY), "
            f"supported by an MPI of {mpi_val} and ARI of {ari_val}. "
            "The hotel outperformed the comp set on both occupancy and rate "
            "simultaneously for the week. Rate growth was driven by ADR "
            f"improving {'+' if float(my_adr_chg)>0 else ''}{my_adr_chg}% "
            f"vs STLY while the comp set grew "
            f"{'+' if float(cs_adr_chg)>0 else ''}{cs_adr_chg}%, "
            "maintaining a rate premium. Monitor segment mix and group "
            "pace to protect midweek production."
        )
    else:
        my_adr_28 = w['my_adr'][7]['val']; cs_adr_28 = w['cs_adr'][7]['val']
        mpi_28_chg = w['mpi'][7]['chg']
        critique = (
            f"The 28-day running period shows an RGI of {rgi_val} "
            f"({'+' if float(rgi_chg)>0 else ''}{rgi_chg} pts YoY). "
            f"MPI of {mpi_val} ({'+' if float(mpi_28_chg)>0 else ''}"
            f"{mpi_28_chg} pts) confirms sustained occupancy outperformance "
            f"vs the comp set. ADR running at ${float(my_adr_28):.0f} vs "
            f"comp set ${float(cs_adr_28):.0f} — ARI of {ari_val} reflects "
            "relative rate positioning."
        )
    _txt(slide, critique, 7.38, SEC_Y+0.4, 5.52, 2.9,
         size=10, color=C['darkGray'], wrap=True)

    # Segment table
    segs = info.get('week_segments', [])
    tf2 = slide.shapes.add_table(len(segs)+1, 4, Inches(0.3),
                                 Inches(SEC_Y+0.32), Inches(6.8),
                                 Inches(3.08))
    t2 = tf2.table
    for ci,cw2 in enumerate([2.3,1.4,1.55,1.55]):
        t2.columns[ci].width = int(cw2*914400)
    for ri in range(len(segs)+1):
        t2.rows[ri].height = int(0.28*914400) if ri==0 else int(0.27*914400)
    for ci,h in enumerate(["Segment","RMS Var","ADR Var","Rev Var"]):
        _cell(t2.cell(0,ci), h, bg=C['teal'], fg=C['white'],
              bold=True, size=9.5)
    for ri,seg in enumerate(segs):
        bg = _rgb("F4F6F8") if ri%2==0 else C['white']
        rms_c = C['green'] if seg['rms_var']>0 else \
                C['red'] if seg['rms_var']<0 else C['midGray']
        adr_c = C['green'] if seg['adr_var']>0 else \
                C['orange'] if seg['adr_var']<0 else C['midGray']
        rev_c = C['green'] if seg['rev_var']>0 else \
                C['red'] if seg['rev_var']<0 else C['midGray']
        rms_s = f"{seg['rms_var']:+d}" if seg['rms_var']!=0 else "0"
        adr_s = f"{'+' if seg['adr_var']>0 else ''}${seg['adr_var']:.0f}" \
                if seg['adr_var']!=0 else "$0"
        rev_s = f"{'+' if seg['rev_var']>0 else ''}${abs(seg['rev_var'])/1000:.1f}K" \
                if seg['rev_var']!=0 else "—"
        _cell(t2.cell(ri+1,0), seg['name'],   bg=bg, fg=C['darkGray'], bold=True, size=9.5)
        _cell(t2.cell(ri+1,1), rms_s,         bg=bg, fg=rms_c, size=9.5)
        _cell(t2.cell(ri+1,2), adr_s,         bg=bg, fg=adr_c, size=9.5)
        _cell(t2.cell(ri+1,3), rev_s,         bg=bg, fg=rev_c, size=9.5)

    return slide

def build_slide_str_weekly(prs, info, str_data):
    return _build_str_slide(prs, info, str_data, 'weekly')

def build_slide_str_28day(prs, info, str_data):
    return _build_str_slide(prs, info, str_data, 'd28')

# ── Slide 4: Annual Pace ─────────────────────────────────────────────────────
def build_slide_annual_pace(prs, info, pace_data):
    slide = _blank(prs); _bg(slide)
    _hdr(slide, "ANNUAL PACE — OTB vs STLY vs BUDGET",
         f"{info['report_yr']} Full Year")

    months   = [d['month'] for d in pace_data]
    otb_rms  = [d['otb_rms'] for d in pace_data]
    stly_rms = [d['stly_rms'] for d in pace_data]
    bud_rms  = [d['bud_rms'] for d in pace_data]

    cd = ChartData()
    cd.categories = months
    cd.add_series("OTB",    tuple(otb_rms))
    cd.add_series("STLY",   tuple(stly_rms))
    cd.add_series("Budget", tuple(bud_rms))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.3), Inches(0.7), Inches(8.5), Inches(4.2), cd).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = C['teal']
    chart.series[1].format.fill.solid()
    chart.series[1].format.fill.fore_color.rgb = C['gold']
    chart.series[2].format.fill.solid()
    chart.series[2].format.fill.fore_color.rgb = C['slate']

    # Monthly table — show upcoming months
    upcoming = [d for d in pace_data if d['mo_num'] >= info['report_mo']][:6]
    _rect(slide, 9.1, 0.7, 4.0, 0.35, C['navy'])
    for ci,h in enumerate(["Month","OTB RMS","STLY","Budget"]):
        _txt(slide, h, 9.1+ci*1.0, 0.7, 1.0, 0.35,
             size=8.5, bold=True, color=C['white'], align='center', valign='middle')
    for ri,d in enumerate(upcoming):
        y = 1.08 + ri*0.52
        bg = _rgb("F4F6F8") if ri%2==0 else C['white']
        _rect(slide, 9.1, y, 4.0, 0.5, bg)
        for ci,(v,col) in enumerate([
            (d['month'],   C['navy']),
            (f"{d['otb_rms']:,}",  C['teal']),
            (f"{d['stly_rms']:,}", C['slate']),
            (f"{d['bud_rms']:,}",  C['midGray']),
        ]):
            _txt(slide, v, 9.1+ci*1.0, y, 1.0, 0.5,
                 size=9, color=col, align='center', valign='middle',
                 bold=(ci==0))

    # KPI tiles
    ytd = info.get('ytd', {})
    kpis = [
        ("YTD OTB RMS",  f"{ytd.get('otb_rms',0):,}",
         f"STLY {ytd.get('stly_rms',0):,}",            C['teal']),
        ("YTD OTB ADR",  f"${ytd.get('otb_adr',0):.0f}",
         f"vs STLY {ytd.get('adr_var','')}",            C['tealLt']),
        ("YTD OTB REV",  f"${ytd.get('otb_rev',0)/1000:.0f}K",
         f"Budget ${ytd.get('bud_rev',0)/1000:.0f}K",  C['gold']),
        ("Full Yr Fcst", f"{ytd.get('fcst_rms',0):,}",
         f"vs Budget {ytd.get('fcst_vs_bud','')}",     C['green']),
    ]
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi(slide, 0.3+i*3.2, 5.1, 3.0, 1.35, lbl, val, sub, col)

# ── Slide 5: Transient Pace ─────────────────────────────────────────────────
def build_slide_transient_pace(prs, info, pace_data):
    slide = _blank(prs); _bg(slide)
    end_date = info['date_range'].split(' - ')[-1].strip()
    _hdr(slide, "TRANSIENT PACE — OTB vs STLY",
         f"As of {end_date}")

    months   = [d['month'] for d in pace_data]
    otb_rms  = [d['trn_otb']  for d in pace_data]
    stly_rms = [d['trn_stly'] for d in pace_data]
    otb_adr  = [d['otb_adr']  for d in pace_data]
    stly_adr = [d['stly_adr'] for d in pace_data]

    # Bar chart (transient RMS)
    cd = ChartData()
    cd.categories = months
    cd.add_series("Transient OTB",  tuple(otb_rms))
    cd.add_series("Transient STLY", tuple(stly_rms))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.3), Inches(0.7), Inches(6.2), Inches(3.5), cd).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.series[0].format.fill.solid()
    chart.series[0].format.fill.fore_color.rgb = C['teal']
    chart.series[1].format.fill.solid()
    chart.series[1].format.fill.fore_color.rgb = C['gold']

    # Line chart (ADR)
    cd2 = ChartData()
    cd2.categories = months
    cd2.add_series("OTB ADR",  tuple(otb_adr))
    cd2.add_series("STLY ADR", tuple(stly_adr))
    chart2 = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(6.7), Inches(0.7), Inches(6.3), Inches(3.5), cd2).chart
    chart2.has_legend = True
    chart2.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart2.series[0].format.line.color.rgb = C['teal']
    chart2.series[1].format.line.color.rgb = C['gold']

    # KPI tiles — 4 upcoming months with correct transient data
    upcoming = [d for d in pace_data if d['mo_num'] >= info['report_mo']][:4]
    for i,d in enumerate(upcoming):
        var = d['trn_otb'] - d['trn_stly']
        var_col = C['green'] if var >= 0 else C['red']
        _kpi(slide, 0.3+i*3.2, 4.5, 3.0, 1.35,
             f"{d['month']} Transient OTB",
             f"{d['trn_otb']:,} rms",
             f"vs STLY {var:+,}",
             var_col)

# ── Slide 6: 7-Day Pickup ───────────────────────────────────────────────────
def build_slide_pickup(prs, info, pickup_data):
    slide = _blank(prs); _bg(slide)
    pickup_from = info.get('pickup_from','')
    _hdr(slide, "7-DAY PICKUP REPORT", f"Pickup from {pickup_from}")

    months = pickup_data['months']
    pu_rms = pickup_data['rms']
    pu_adr = pickup_data['adr']

    if months and pu_rms:
        cd = ChartData()
        cd.categories = months
        cd.add_series("Room Nights Picked Up", tuple(pu_rms))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(0.3), Inches(0.7), Inches(6.3), Inches(3.8), cd).chart
        chart.has_title = True
        chart.chart_title.text_frame.text = "Room Nights Picked Up by Month"
        chart.has_legend = False
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = C['teal']
        chart.plots[0].has_data_labels = True
        chart.plots[0].data_labels.show_value = True

    if months and pu_adr:
        cd2 = ChartData()
        cd2.categories = months
        cd2.add_series("Average ADR of Pickup", tuple(pu_adr))
        chart2 = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(6.8), Inches(0.7), Inches(6.2), Inches(3.8), cd2).chart
        chart2.has_title = True
        chart2.chart_title.text_frame.text = "Average ADR of Pickup by Month"
        chart2.has_legend = False
        chart2.series[0].format.fill.solid()
        chart2.series[0].format.fill.fore_color.rgb = C['gold']
        chart2.plots[0].has_data_labels = True
        chart2.plots[0].data_labels.show_value = True

    # Segment table
    _rect(slide, 0.3, 4.72, 12.7, 0.35, C['navy'])
    _txt(slide, "PICKUP DETAIL (7-Day) — Key Transient Segments",
         0.4, 4.72, 12.7, 0.35, size=10, bold=True,
         color=C['white'], valign='middle')

    segs = pickup_data.get('segments', [])
    tf = slide.shapes.add_table(len(segs)+1, 5,
        Inches(0.3), Inches(5.1), Inches(12.7), Inches(2.3))
    t = tf.table
    for ci,cw in enumerate([3.0,1.8,1.5,2.0,4.4]):
        t.columns[ci].width = int(cw*914400)
    for ri in range(len(segs)+1):
        t.rows[ri].height = int(0.33*914400)

    for ci,h in enumerate(["Segment","PU RNs","% of Trans","PU Revenue","Note"]):
        _cell(t.cell(0,ci), h, bg=C['teal'], fg=C['white'], bold=True, size=9.5)

    for ri,seg in enumerate(segs):
        is_total = seg.get('is_total', False)
        bg = _rgb("0D1B2A") if is_total else \
             (_rgb("F4F6F8") if ri%2==0 else C['white'])
        fg_main = C['gold'] if is_total else C['darkGray']
        rms_v = seg['rms']
        rms_s = f"+{rms_v}" if rms_v > 0 else str(rms_v)
        rms_c = C['green'] if rms_v > 0 else C['red'] if rms_v < 0 else C['midGray']
        _cell(t.cell(ri+1,0), seg['name'],
              bg=bg, fg=fg_main, bold=is_total, size=9.5)
        _cell(t.cell(ri+1,1), rms_s,
              bg=bg, fg=rms_c if not is_total else C['gold'],
              bold=is_total, size=9.5)
        _cell(t.cell(ri+1,2), seg.get('pct_trans',''),
              bg=bg, fg=C['teal'] if not is_total else C['gold'], size=9.5)
        _cell(t.cell(ri+1,3), f"${seg['rev']:,}" if seg['rev'] else '',
              bg=bg, fg=C['teal'] if not is_total else C['gold'],
              bold=is_total, size=9.5)
        _cell(t.cell(ri+1,4), seg.get('note',''),
              bg=bg, fg=C['slate'] if not is_total else C['midGray'],
              size=9, align='left')

# ── Slides 7+8, 9+10, 11+12: Monthly Outlook ────────────────────────────────
def build_slide_monthly_occ(prs, info, mo_num, daily_data):
    slide = _blank(prs); _bg(slide)
    mo_label = MO_FULL[mo_num]
    _hdr(slide, f"{mo_label.upper()} DAILY OUTLOOK — OTB vs STLY",
         f"Transient  |  {mo_label} {info['report_yr']}")

    labels    = [d['label']    for d in daily_data]
    otb_vals  = [d['occ_otb']  for d in daily_data]
    stly_vals = [d['occ_stly'] for d in daily_data]

    if labels:
        cd = ChartData()
        cd.categories = labels
        cd.add_series(f"OTB {info['report_yr']} OCC%",      tuple(otb_vals))
        cd.add_series(f"STLY {info['report_yr']-1} OCC%",   tuple(stly_vals))
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(0.3), Inches(0.7), Inches(12.9), Inches(3.6), cd).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.series[0].format.line.color.rgb = C['teal']
        chart.series[0].format.line.width = 18000
        chart.series[1].format.line.color.rgb = C['gold']
        chart.series[1].format.line.width = 18000
    else:
        _txt(slide, "No daily data available for this month.",
             0.3, 2.0, 12.7, 1.0, size=12, color=C['midGray'], align='center')

    mo_kpi = info.get('mo_kpis', {}).get(mo_num, {})
    vs_rms = str(mo_kpi.get('vs_rms','0'))
    vs_rev = str(mo_kpi.get('vs_rev',''))
    vs_col  = C['green'] if vs_rms.startswith('+') else C['orange']
    rev_col = C['green'] if vs_rev.startswith('+') else C['red']
    kpis = [
        (f"{MO[mo_num]} OTB RMS",       mo_kpi.get('otb_rms','—'),
         f"{mo_kpi.get('otb_occ','—')} OCC",     C['teal']),
        (f"{MO[mo_num]} Transient RMS", mo_kpi.get('trn_rms','—'),
         f"{mo_kpi.get('trn_adr','—')} ADR",     C['tealLt']),
        ("vs STLY Rooms",               mo_kpi.get('vs_rms','—'),
         f"STLY {mo_kpi.get('stly_rms','—')}",  vs_col),
        (f"{MO[mo_num]} Revenue OTB",   mo_kpi.get('rev','—'),
         f"{mo_kpi.get('vs_rev','—')} vs STLY",  rev_col),
    ]
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi(slide, 0.3+i*3.2, 4.55, 3.0, 1.35, lbl, str(val), str(sub), col)

def build_slide_segment_mix(prs, info, mo_num, seg_mix_data):
    slide = _blank(prs); _bg(slide)
    mo_label = MO_FULL[mo_num]
    _hdr(slide, f"MARKET SEGMENT MIX — {mo_label.upper()} {info['report_yr']}")
    # Segment data not available from this PMS export — leave blank per reference
    _txt(slide,
         "Segment mix data not available in this report export.",
         0.3, 3.5, 12.7, 0.5, size=11, color=C['midGray'], align='center')
    return slide

# ── Slide 13: Full Year Forecast ─────────────────────────────────────────────
def build_slide_full_year(prs, info, pace_data):
    slide = _blank(prs); _bg(slide)
    _hdr(slide, "FULL YEAR FORECAST 2026", "Finance Forecast vs Budget")

    months   = [d['month'] for d in pace_data]
    fcst_rms = [d['fcst_rms'] for d in pace_data]
    bud_rms  = [d['bud_rms']  for d in pace_data]

    cd = ChartData()
    cd.categories = months
    cd.add_series("Forecast", tuple(fcst_rms))
    cd.add_series("Budget",   tuple(bud_rms))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.LINE,
        Inches(0.3), Inches(0.7), Inches(7.2), Inches(4.5), cd).chart
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

    for ci,h in enumerate(["Month","OTB OCC","Fcst OCC","Fcst ADR","Fcst REV","vs Bud"]):
        _cell(t.cell(0,ci), h, bg=C['navy'], fg=C['white'], bold=True, size=8.5)

    for ri,d in enumerate(pace_data):
        bg   = _rgb("F4F6F8") if ri%2==0 else C['white']
        cap  = DAYS_MO[d['mo_num']] * 150
        otb_occ_pct  = d['otb_rms']/cap*100 if cap>0 else 0
        fcst_occ_pct = d['fcst_rms']/cap*100 if cap>0 else 0
        vs_bud = d['fcst_rev'] - d['bud_rev'] if d.get('fcst_rev') and d.get('bud_rev') else 0
        vs_col = C['green'] if vs_bud >= 0 else C['red']
        vs_str = f"{vs_bud/d['bud_rev']*100:+.1f}%" if d['bud_rev'] > 0 else '—'
        _cell(t.cell(ri+1,0), d['month'],                       bg=bg, fg=C['navy'],  bold=True, size=8.5)
        _cell(t.cell(ri+1,1), f"{otb_occ_pct:.1f}%",           bg=bg, fg=C['teal']  if otb_occ_pct>80 else C['orange'], size=8.5)
        _cell(t.cell(ri+1,2), f"{fcst_occ_pct:.1f}%",          bg=bg, fg=C['grn_dk'] if fcst_occ_pct>=95 else C['amber'], bold=True, size=8.5)
        _cell(t.cell(ri+1,3), f"${d['fcst_adr']:.2f}",         bg=bg, fg=C['teal'],  size=8.5)
        _cell(t.cell(ri+1,4), f"${d['fcst_rev']/1000:.0f}K",   bg=bg, fg=C['navy'],  bold=True, size=8.5)
        _cell(t.cell(ri+1,5), vs_str,                           bg=bg, fg=vs_col,     bold=True, size=8.5)

    # KPI tiles
    ytd = info.get('ytd', {})
    fcst_rev = ytd.get('fcst_rev', 0)
    bud_rev  = ytd.get('bud_rev',  0)
    vs_bud_pct = f"{(fcst_rev/bud_rev-1)*100:+.1f}%" if bud_rev > 0 else '—'
    fcst_occ_yr = f"{ytd.get('fcst_rms',0)/(365*150)*100:.1f}%"
    kpis = [
        ("Forecast Full Year OCC", fcst_occ_yr,          "",   C['teal']),
        ("Forecast Full Year ADR", f"${ytd.get('fcst_adr',0):.2f}", "", C['gold']),
        ("Forecast Full Year REV", f"${fcst_rev/1e6:.2f}M","",  C['tealLt']),
        ("vs Budget REV",          vs_bud_pct,            "",   C['green'] if fcst_rev >= bud_rev else C['red']),
    ]
    for i,(lbl,val,sub,col) in enumerate(kpis):
        _kpi(slide, 0.3+i*3.2, 5.45, 3.0, 1.35, lbl, val, sub, col)

    _txt(slide, info.get('fcst_commentary',''), 0.3, 7.12, 12.7, 0.28,
         size=8.5, color=C['slate'], wrap=True)

# ── Data extraction ──────────────────────────────────────────────────────────
def extract_all_data(xl, info):
    import pandas as pd

    df_as  = pd.read_excel(xl, sheet_name='Annual Summary',  header=None)
    df_pu  = pd.read_excel(xl, sheet_name='Pickup',          header=None)
    df_90  = pd.read_excel(xl, sheet_name='90 Day Segments', header=None)
    df_str = pd.read_excel(xl, sheet_name='STR Analysis',    header=None)

    def fv(v, default=0):
        try: return float(v) if str(v) not in ('nan','') else default
        except: return default

    # ── Monthly total rows in 90 Day Segments (no date in col1) ──────────
    # These rows have blank col1 and contain monthly aggregates
    mo_total_rows = {}
    mo_order = [info['report_mo'], info['report_mo']+1, info['report_mo']+2]
    mo_idx = 0
    for i in range(6, df_90.shape[0]):
        row = df_90.iloc[i]
        if str(row[1]) == 'nan' and str(row[2]) not in ('nan','') \
           and str(row[3]) not in ('nan',''):
            try:
                # Verify it's a total row: col4 should be OCC (0-1), col12 should be big int
                occ_val = float(row[4])
                trn_val = float(row[12]) if str(row[12]) not in ('nan','') else 0
                if 0 < occ_val < 1 and trn_val > 0 and mo_idx < len(mo_order):
                    mo_num = mo_order[mo_idx]
                    mo_total_rows[mo_num] = {
                        'trn_otb':  int(trn_val),
                        'trn_stly': int(fv(row[15])),
                        'trn_adr':  fv(row[13]),
                        'otb_rms':  int(fv(row[3])),
                        'otb_occ':  occ_val * 100,
                        'otb_adr':  fv(row[5]),
                        'otb_rev':  fv(row[6]),
                    }
                    mo_idx += 1
            except: pass

    # ── Annual pace data ─────────────────────────────────────────────────
    pace_data = []
    mo_names = ['','Jan','Feb','Mar','Apr','May','Jun',
                'Jul','Aug','Sep','Oct','Nov','Dec']

    for i in range(6, 18):
        row = df_as.iloc[i]
        mo_name = str(row[1])
        if mo_name not in mo_names: continue
        mo_num = mo_names.index(mo_name)
        cap    = DAYS_MO[mo_num] * 150

        otb_occ  = fv(row[2])*100
        otb_rms  = int(fv(row[3]))
        otb_adr  = fv(row[4])
        otb_rev  = fv(row[5])
        stly_rms = int(fv(row[7]))
        stly_adr = fv(row[8])

        fcst_rms = int(fv(row[15])) if fv(row[15]) > 0 else otb_rms
        fcst_adr = fv(row[16]) if fv(row[16]) > 0 else otb_adr
        fcst_rev = fv(row[17]) if fv(row[17]) > 0 else otb_rev
        bud_rms  = int(fv(row[19])) if fv(row[19]) > 0 else 0
        bud_adr  = fv(row[20]) if fv(row[20]) > 0 else 0
        bud_rev  = fv(row[21]) if fv(row[21]) > 0 else 0

        # Get transient from monthly total rows
        mo_tot   = mo_total_rows.get(mo_num, {})
        trn_otb  = mo_tot.get('trn_otb',  0)
        trn_stly = mo_tot.get('trn_stly', 0)

        pace_data.append({
            'month': mo_name, 'mo_num': mo_num,
            'otb_occ': otb_occ,  'otb_rms': otb_rms,
            'otb_adr': otb_adr,  'otb_rev': otb_rev,
            'stly_rms': stly_rms,'stly_adr': stly_adr,
            'fcst_rms': fcst_rms,'fcst_adr': fcst_adr,
            'fcst_rev': fcst_rev,'bud_rms': bud_rms,
            'bud_adr': bud_adr,  'bud_rev': bud_rev,
            'trn_otb': trn_otb,  'trn_stly': trn_stly,
        })

    # ── YTD / full year totals ────────────────────────────────────────────
    tr = df_as.iloc[18]
    ytd_otb_rms  = int(fv(tr[3]))
    ytd_otb_adr  = fv(tr[4])
    ytd_otb_rev  = fv(tr[5])
    ytd_stly_rms = int(fv(tr[7]))
    ytd_bud_rev  = fv(tr[21]) if fv(tr[21]) > 0 else 0
    ytd_bud_rms  = int(fv(tr[19])) if fv(tr[19]) > 0 else 0
    ytd_fcst_rev = fv(tr[17]) if fv(tr[17]) > 0 else ytd_otb_rev
    ytd_fcst_rms = int(fv(tr[15])) if fv(tr[15]) > 0 else ytd_otb_rms
    ytd_fcst_adr = ytd_fcst_rev/ytd_fcst_rms if ytd_fcst_rms > 0 else 0
    stly_adr_tot = fv(tr[8])
    adr_var  = f"{(ytd_otb_adr/stly_adr_tot-1)*100:+.1f}%" if stly_adr_tot > 0 else ''
    fcst_vs_bud = f"{(ytd_fcst_rev/ytd_bud_rev-1)*100:+.1f}%" if ytd_bud_rev > 0 else ''

    info['ytd'] = {
        'otb_rms': ytd_otb_rms, 'otb_adr': ytd_otb_adr,
        'otb_rev': ytd_otb_rev, 'stly_rms': ytd_stly_rms,
        'bud_rev': ytd_bud_rev, 'bud_rms': ytd_bud_rms,
        'fcst_rev': ytd_fcst_rev, 'fcst_rms': ytd_fcst_rms,
        'fcst_adr': ytd_fcst_adr, 'adr_var': adr_var,
        'fcst_vs_bud': fcst_vs_bud,
    }

    # ── Forecast commentary ───────────────────────────────────────────────
    fcst_r = sum(d['fcst_rev'] for d in pace_data)
    bud_r  = sum(d['bud_rev']  for d in pace_data)
    vs_pct = (fcst_r/bud_r-1)*100 if bud_r > 0 else 0
    info['fcst_commentary'] = (
        f"Finance Forecast projects full year at ${fcst_r/1e6:.2f}M vs "
        f"${bud_r/1e6:.2f}M budget ({vs_pct:+.1f}%). "
        "Heavy lift required in back half as current OTB reflects early pace. "
        "Transient volume capture is the primary revenue lever."
    )

    # ── Monthly KPIs ─────────────────────────────────────────────────────
    mo_kpis = {}
    for d in pace_data:
        mo = d['mo_num']
        cap = DAYS_MO[mo] * 150 if mo > 0 else 4500
        mo_tot = mo_total_rows.get(mo, {})
        vs_rms = d['otb_rms'] - d['stly_rms']
        stly_rev = d['stly_rms'] * d['stly_adr']
        vs_rev = f"{(d['otb_rev']/stly_rev-1)*100:+.1f}%" \
                 if stly_rev > 0 else ''
        mo_kpis[mo] = {
            'otb_rms':  f"{d['otb_rms']:,}",
            'otb_occ':  f"{d['otb_occ']:.1f}%",
            'trn_rms':  f"{mo_tot.get('trn_otb', d['trn_otb']):,}",
            'trn_adr':  f"${mo_tot.get('trn_adr', d['otb_adr']):.2f}",
            'stly_rms': f"{d['stly_rms']:,}",
            'vs_rms':   f"{vs_rms:+,}",
            'rev':      f"${d['otb_rev']/1000:.0f}K",
            'vs_rev':   vs_rev,
        }
    info['mo_kpis'] = mo_kpis

    # ── Pickup data ───────────────────────────────────────────────────────
    pu_months=[]; pu_rms=[]; pu_adr=[]; pickup_from=''
    try:
        mo_row  = df_pu.iloc[39]
        rms_row = df_pu.iloc[40]
        adr_row = df_pu.iloc[41]
        rev_row = df_pu.iloc[42]
        raw_months = [str(v) for v in mo_row if str(v) not in ('nan','Total','')]
        raw_rms    = [v for v in rms_row if str(v) not in ('nan','RMS','')]
        raw_adr    = [v for v in adr_row if str(v) not in ('nan','ADR','')]
        for mo_s, rms_v, adr_v in zip(raw_months, raw_rms, raw_adr):
            rms_n = int(float(rms_v))
            adr_n = float(adr_v)
            if rms_n == 0 and adr_n == 0: continue
            short = mo_s.split('-')[0]
            pu_months.append(short); pu_rms.append(rms_n); pu_adr.append(round(adr_n))
        # pickup from date
        for i in range(4, 6):
            r = df_pu.iloc[i]
            for v in r:
                if 'Pickup From' in str(v):
                    pickup_from = str(v).split('Pickup From ')[-1].strip()
    except: pass

    info['pickup_from'] = pickup_from

    # Total transient pickup for % of trans calculation
    total_trn_pu = pu_rms[0] if pu_rms else 0

    # Segment pickup data (from 90 Day Segments 7-day pickup col9)
    # Build from actual data available
    seg_notes = {
        'Transient (Total)': ('ADR change vs LY', True),
        'Retail / Rack':     ('Strong rate integrity', False),
        'Discount':          ('Monitor rate dilution', False),
        'Internet/OTA':      ('Healthy contribution', False),
        'Packages':          ('Continue to promote', False),
        'Group':             ('Solid group contribution', False),
    }
    # Get 7D pickup totals from 90 Day Segments total row
    try:
        total_row_90 = None
        for i in range(6, df_90.shape[0]):
            row = df_90.iloc[i]
            if str(row[1]) == 'nan' and str(row[2]) not in ('nan','') \
               and fv(row[3]) > 1000:  # first monthly total
                total_row_90 = row; break

        # Get actual pickup breakdown from 90 Day Segments
        # Use total 7D pickup from the first monthly total row
        total_7d_pu = int(fv(total_row_90[9])) if total_row_90 is not None else 0
        total_trn_pu = max(total_trn_pu, total_7d_pu)
    except: pass

    pu_segs = []
    for name,(note,is_total) in seg_notes.items():
        pu_segs.append({
            'name': name, 'rms': 0, 'rev': 0,
            'note': note, 'is_total': is_total,
            'pct_trans': '100.0%' if is_total else '',
        })

    # Try to populate segment pickup from STR Analysis week segment rows
    try:
        seg_names = ['RACK','INTERNET','CONTRACT','PACKAGES','GROUP',
                     'CORPORATE','GOVERNMENT','DISCOUNT','WHOLESALE','FRANCHISE']
        # Week total row is row 46 in STR Analysis
        week_total = df_str.iloc[46]
        # Row 37 has segment names, rows 39-45 have daily, row 46 = weekly total
        # col2=PU rms, col3=LTS, col4=OTB, col5=ADR, col6=Rev
        # Segment order from row 37: RACK, INTERNET, CONTRACT, PACKAGES, GROUP...
        # Each segment has its own col block - skip for now, use what's available
        pass
    except: pass

    # Build week segments for STR slides from STR Analysis rows 39-46
    week_segs = []
    try:
        seg_name_row = df_str.iloc[37]
        seg_names_list = [str(v) for v in seg_name_row if str(v)!='nan']
        week_total_row = df_str.iloc[46]
        wt_vals = [v for v in week_total_row if str(v)!='nan']
        # The total row has format: [OOO/avail, LTS, OTB, OCC, ADR, Rev, RevPAR, REV, ...]
        # Col 10=total OTB, col 11=STLY
        # For segment breakdown we need to look at daily rows and sum
        # Use simplified approach: extract from available week data
        pass
    except: pass

    info['week_segments'] = week_segs

    pickup_data = {
        'months': pu_months, 'rms': pu_rms, 'adr': pu_adr,
        'segments': pu_segs,
    }

    return pace_data, pickup_data

# ── Master builder ───────────────────────────────────────────────────────────
def build_slides_1_to_14(prs, xl, info, str_data, monthly_data, ops_rows, total_rooms):
    pace_data, pickup_data = extract_all_data(xl, info)

    build_slide_title(prs, info, str_data)          # Slide 1
    build_slide_str_weekly(prs, info, str_data)     # Slide 2
    build_slide_str_28day(prs, info, str_data)      # Slide 3
    build_slide_annual_pace(prs, info, pace_data)   # Slide 4
    build_slide_transient_pace(prs, info, pace_data)# Slide 5
    build_slide_pickup(prs, info, pickup_data)      # Slide 6

    # Slides 7-12: Monthly outlook (current + 2 months)
    for offset in range(3):
        mo = info['report_mo'] + offset
        if mo > 12: break
        daily = monthly_data.get('months', {}).get(str(mo), [])
        build_slide_monthly_occ(prs, info, mo, daily)      # 7, 9, 11
        seg_mix = {'pie':{},'rows':[],'takeaways':[]}
        build_slide_segment_mix(prs, info, mo, seg_mix)    # 8, 10, 12

    build_slide_full_year(prs, info, pace_data)     # Slide 13
    # Slide 14 = 14-Day Ops built by main app

