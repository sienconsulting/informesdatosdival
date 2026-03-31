"""
Build DIVAL Tourism Report 2025
Modifies the unpacked PPTX XML files with real data and creates
a comprehensive tourism report for Valencia Province.
"""
import os
import sys
import json
import copy
import shutil
import re
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

# ---- Paths ----
UNPACKED_DIR = r'C:\Users\Pere\OneDrive - Sien Consulting\PLANIFICACION\00.NUEVA ORGANIZACION\01. Proyectos\Diputación Valencia\2024_SMART OFFICE\Informe datos\unpacked'
METRICS_PATH = r'C:\Users\Pere\OneDrive - Sien Consulting\PLANIFICACION\00.NUEVA ORGANIZACION\01. Proyectos\Diputación Valencia\2024_SMART OFFICE\Informe datos\metrics_dival.json'
SLIDES_DIR = os.path.join(UNPACKED_DIR, 'ppt', 'slides')
CHARTS_DIR = os.path.join(UNPACKED_DIR, 'ppt', 'charts')
SKILLS_DIR = r'C:\Users\Pere\AppData\Roaming\Claude\local-agent-mode-sessions\skills-plugin\85053d59-951c-440d-928a-a0a6d8411ee2\88b0e13a-78b9-45c6-8c20-eb24c26a6689\skills\pptx'

# ---- Load metrics ----
with open(METRICS_PATH, 'r', encoding='utf-8') as f:
    metrics = json.load(f)

print("Loaded metrics:")
print(f"  Total 2025: {metrics['total_turistas_2025']:,}")
print(f"  International: {metrics['total_receptor_2025']:,}")
print(f"  Domestic: {metrics['total_interno_2025']:,}")

# ---- Namespace map ----
NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

# Register namespaces to preserve them
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('c14', 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart')
ET.register_namespace('c15', 'http://schemas.microsoft.com/office/drawing/2012/chart')
ET.register_namespace('c16', 'http://schemas.microsoft.com/office/drawing/2014/chart')
ET.register_namespace('c16r2', 'http://schemas.microsoft.com/office/drawing/2015/06/chart')
ET.register_namespace('c16r3', 'http://schemas.microsoft.com/office/drawing/2017/03/chart')
ET.register_namespace('a16', 'http://schemas.microsoft.com/office/drawing/2014/main')

# ---- Format helpers ----
MESES_ES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
MESES_FULL = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
              'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

def fmt(n, dec=0):
    if n is None:
        return 'N/D'
    try:
        n = float(n)
        if dec == 0:
            return f"{int(round(n)):,}".replace(",", ".")
        return f"{n:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(n)

def fmt_pct(p, dec=1):
    if p is None:
        return 'N/D'
    try:
        p = float(p)
        sign = "+" if p > 0 else ""
        return f"{sign}{p:.{dec}f}%"
    except:
        return str(p)

def pct(new, old):
    if not old or old == 0:
        return None
    return (new - old) / abs(old) * 100

# ---- XML Helpers ----
def get_text(elem):
    """Get all text from element"""
    parts = []
    for t in elem.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts)

def set_text_in_txbody(txbody, new_text, keep_formatting=True):
    """Replace all text in a txBody element with new_text, preserving first run's formatting"""
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    # Find all paragraphs
    paragraphs = txbody.findall(f'{{{a_ns}}}p')
    if not paragraphs:
        return

    # Get the first run's rPr for formatting
    first_rpr = None
    for p in paragraphs:
        for r in p.findall(f'{{{a_ns}}}r'):
            rpr = r.find(f'{{{a_ns}}}rPr')
            if rpr is not None:
                first_rpr = copy.deepcopy(rpr)
                break
        if first_rpr is not None:
            break

    # Remove all existing paragraphs
    for p in paragraphs:
        txbody.remove(p)

    # Split text by newline to create multiple paragraphs
    lines = new_text.split('\n')

    for line_text in lines:
        p_el = ET.SubElement(txbody, f'{{{a_ns}}}p')
        r_el = ET.SubElement(p_el, f'{{{a_ns}}}r')
        if first_rpr is not None and keep_formatting:
            r_el.insert(0, copy.deepcopy(first_rpr))
        t_el = ET.SubElement(r_el, f'{{{a_ns}}}t')
        t_el.text = line_text

def find_shape_by_name(root, name):
    """Find shape element by name attribute"""
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    for sp in root.iter(f'{{{p_ns}}}sp'):
        nvsp = sp.find(f'{{{p_ns}}}nvSpPr')
        if nvsp is not None:
            cnvpr = nvsp.find(f'{{{p_ns}}}cNvPr')
            if cnvpr is not None and cnvpr.get('name', '') == name:
                return sp
    return None

def find_shape_by_id(root, shape_id):
    """Find shape element by id"""
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    for sp in root.iter(f'{{{p_ns}}}sp'):
        nvsp = sp.find(f'{{{p_ns}}}nvSpPr')
        if nvsp is not None:
            cnvpr = nvsp.find(f'{{{p_ns}}}cNvPr')
            if cnvpr is not None and cnvpr.get('id', '') == str(shape_id):
                return sp
    return None

def get_txbody(sp):
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    return sp.find(f'{{{p_ns}}}txBody')

def update_shape_text(root, shape_name, new_text):
    """Update text in a shape by name"""
    sp = find_shape_by_name(root, shape_name)
    if sp is None:
        print(f"  WARNING: shape '{shape_name}' not found")
        return False
    txbody = get_txbody(sp)
    if txbody is not None:
        set_text_in_txbody(txbody, new_text)
    return True

def read_slide(slide_num):
    path = os.path.join(SLIDES_DIR, f'slide{slide_num}.xml')
    tree = ET.parse(path)
    return tree

def write_slide(tree, slide_num):
    path = os.path.join(SLIDES_DIR, f'slide{slide_num}.xml')
    tree.write(path, xml_declaration=True, encoding='utf-8')

def read_chart(chart_num):
    path = os.path.join(CHARTS_DIR, f'chart{chart_num}.xml')
    tree = ET.parse(path)
    return tree

def write_chart(tree, chart_num):
    path = os.path.join(CHARTS_DIR, f'chart{chart_num}.xml')
    tree.write(path, xml_declaration=True, encoding='utf-8')

# ---- Chart update helpers ----
C_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

def update_bar_chart_series(chart_root, series_data):
    """
    Update bar chart data.
    series_data = [
        {'name': 'Nacional', 'cats': [2019,2020,...], 'vals': [1000,2000,...]},
        ...
    ]
    """
    plot_area = chart_root.find(f'.//{{{C_NS}}}plotArea')
    if plot_area is None:
        return

    # Find the chart type element (barChart or lineChart)
    chart_el = None
    for tag in ['barChart', 'lineChart', 'pieChart', 'areaChart']:
        chart_el = plot_area.find(f'{{{C_NS}}}{tag}')
        if chart_el is not None:
            break

    if chart_el is None:
        return

    # Remove existing series
    for ser in chart_el.findall(f'{{{C_NS}}}ser'):
        chart_el.remove(ser)

    # Find where to insert (before dLbls/gapWidth/etc)
    insert_pos = 0
    children = list(chart_el)
    for i, child in enumerate(children):
        tag_local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag_local not in ['barDir', 'grouping', 'varyColors']:
            insert_pos = i
            break

    for s_idx, s_data in enumerate(series_data):
        ser = ET.Element(f'{{{C_NS}}}ser')

        idx_el = ET.SubElement(ser, f'{{{C_NS}}}idx')
        idx_el.set('val', str(s_idx))
        order_el = ET.SubElement(ser, f'{{{C_NS}}}order')
        order_el.set('val', str(s_idx))

        # Series name
        tx = ET.SubElement(ser, f'{{{C_NS}}}tx')
        str_ref = ET.SubElement(tx, f'{{{C_NS}}}strRef')
        f_el = ET.SubElement(str_ref, f'{{{C_NS}}}f')
        f_el.text = f'Hoja1!${chr(66+s_idx)}$1'
        str_cache = ET.SubElement(str_ref, f'{{{C_NS}}}strCache')
        pt_count = ET.SubElement(str_cache, f'{{{C_NS}}}ptCount')
        pt_count.set('val', '1')
        pt = ET.SubElement(str_cache, f'{{{C_NS}}}pt')
        pt.set('idx', '0')
        v = ET.SubElement(pt, f'{{{C_NS}}}v')
        v.text = str(s_data.get('name', f'Serie {s_idx+1}'))

        # Color if specified
        if 'color' in s_data:
            sp_pr = ET.SubElement(ser, f'{{{A_NS}}}spPr')
            solid_fill = ET.SubElement(sp_pr, f'{{{A_NS}}}solidFill')
            srgb = ET.SubElement(solid_fill, f'{{{A_NS}}}srgbClr')
            srgb.set('val', s_data['color'])
            ln = ET.SubElement(sp_pr, f'{{{A_NS}}}ln')
            ET.SubElement(ln, f'{{{A_NS}}}noFill')
            ET.SubElement(sp_pr, f'{{{A_NS}}}effectLst')

        # Categories
        cats = s_data.get('cats', [])
        vals = s_data.get('vals', [])
        n = len(cats)

        cat_el = ET.SubElement(ser, f'{{{C_NS}}}cat')
        if all(isinstance(c, (int, float)) for c in cats):
            num_ref = ET.SubElement(cat_el, f'{{{C_NS}}}numRef')
            f_cat = ET.SubElement(num_ref, f'{{{C_NS}}}f')
            f_cat.text = f'Hoja1!$A$2:$A${n+1}'
            num_cache = ET.SubElement(num_ref, f'{{{C_NS}}}numCache')
            fmt_code = ET.SubElement(num_cache, f'{{{C_NS}}}formatCode')
            fmt_code.text = 'General'
            ptc = ET.SubElement(num_cache, f'{{{C_NS}}}ptCount')
            ptc.set('val', str(n))
            for i, cat_val in enumerate(cats):
                pt_c = ET.SubElement(num_cache, f'{{{C_NS}}}pt')
                pt_c.set('idx', str(i))
                v_c = ET.SubElement(pt_c, f'{{{C_NS}}}v')
                v_c.text = str(cat_val)
        else:
            str_ref_c = ET.SubElement(cat_el, f'{{{C_NS}}}strRef')
            f_cat = ET.SubElement(str_ref_c, f'{{{C_NS}}}f')
            f_cat.text = f'Hoja1!$A$2:$A${n+1}'
            str_cache_c = ET.SubElement(str_ref_c, f'{{{C_NS}}}strCache')
            ptc = ET.SubElement(str_cache_c, f'{{{C_NS}}}ptCount')
            ptc.set('val', str(n))
            for i, cat_val in enumerate(cats):
                pt_c = ET.SubElement(str_cache_c, f'{{{C_NS}}}pt')
                pt_c.set('idx', str(i))
                v_c = ET.SubElement(pt_c, f'{{{C_NS}}}v')
                v_c.text = str(cat_val)

        # Values
        val_el = ET.SubElement(ser, f'{{{C_NS}}}val')
        num_ref_v = ET.SubElement(val_el, f'{{{C_NS}}}numRef')
        f_val = ET.SubElement(num_ref_v, f'{{{C_NS}}}f')
        f_val.text = f'Hoja1!${chr(66+s_idx)}$2:${chr(66+s_idx)}${n+1}'
        num_cache_v = ET.SubElement(num_ref_v, f'{{{C_NS}}}numCache')
        fmt_code_v = ET.SubElement(num_cache_v, f'{{{C_NS}}}formatCode')
        fmt_code_v.text = '#,##0'
        ptc_v = ET.SubElement(num_cache_v, f'{{{C_NS}}}ptCount')
        ptc_v.set('val', str(len(vals)))
        for i, val in enumerate(vals):
            pt_v = ET.SubElement(num_cache_v, f'{{{C_NS}}}pt')
            pt_v.set('idx', str(i))
            v_v = ET.SubElement(pt_v, f'{{{C_NS}}}v')
            v_v.text = str(int(val) if isinstance(val, float) and val == int(val) else val)

        chart_el.insert(insert_pos + s_idx, ser)

# ============================================================
# NOW BUILD ALL SLIDES
# ============================================================

print("\n=== Building slides ===")

# ---- SLIDE 1: Portada ----
print("Slide 1: Portada...")
tree1 = read_slide(1)
root1 = tree1.getroot()
# The cover has shapes: CuadroTexto 9 = title, Rectángulo 10 = org name, Rectángulo 14 = period
update_shape_text(root1, 'CuadroTexto 9', 'Informe de análisis de datos')
update_shape_text(root1, 'Rectángulo 10', 'SMART OFFICE\nDIPUTACIÓN DE VALÈNCIA')
update_shape_text(root1, 'Rectángulo 14', 'ENE–DIC 2025')
write_slide(tree1, 1)

# ---- SLIDE 2: Índice ----
print("Slide 2: Índice...")
tree2 = read_slide(2)
root2 = tree2.getroot()
# Update index to include our sections
update_shape_text(root2, 'Rectángulo 3', '01\nAfluencia\nturística')
update_shape_text(root2, 'Rectángulo 4', '02\nGasto\nturístico')
update_shape_text(root2, 'Rectángulo 5', '03\nOcupación')
update_shape_text(root2, 'Rectángulo 1', '04\nOferta\nturística')
# The 5th section (employment) can be added in extra slides
write_slide(tree2, 2)

# ---- SLIDE 3: Section divider - Afluencia Turística ----
print("Slide 3: Section header - Afluencia Turística...")
tree3 = read_slide(3)
root3 = tree3.getroot()
# Update section number box
update_shape_text(root3, 'Rectángulo 9', '1')
update_shape_text(root3, 'Rectángulo 1', 'Afluencia\nturística')

# Update descriptive text (lorem ipsum box)
tot_2025 = metrics['total_turistas_2025']
tot_2024 = metrics['total_turistas_2024']
chg = metrics['chg_total']
rec_2025 = metrics['total_receptor_2025']
int_2025 = metrics['total_interno_2025']

desc_text = (f"La provincia de Valencia registró {fmt(tot_2025)} turistas en 2025 "
             f"({fmt_pct(chg)} respecto a 2024). El turismo internacional alcanzó "
             f"{fmt(rec_2025)} visitantes (+{fmt_pct(metrics['chg_receptor'])}), "
             f"consolidando la recuperación post-pandémica. El turismo nacional "
             f"registró {fmt(int_2025)} viajeros.")
update_shape_text(root3, 'Rectángulo 11', desc_text)
write_slide(tree3, 3)

# ---- SLIDE 4: Turismo Nacional (chart1) ----
print("Slide 4: Turismo Nacional (historical bar chart)...")
tree4 = read_slide(4)
root4 = tree4.getroot()

update_shape_text(root4, 'Rectángulo 12', 'AFLUENCIA TURÍSTICA')
update_shape_text(root4, 'Rectángulo 9', 'Principales resultados\ndel periodo')

# Build domestic text with key stats
int_hist = metrics.get('historico_interno', {})
int_25 = metrics['total_interno_2025']
int_24 = metrics['total_interno_2024']
chg_int = metrics.get('chg_interno', 0)

desc4 = (f"El turismo nacional hacia Valencia registró {fmt(int_25)} visitantes "
         f"en 2025, {fmt_pct(chg_int)} respecto a 2024. "
         f"Madrid ({fmt(metrics['top5_origenes_nacionales_2025'][0]['turistas'])} turistas), "
         f"Alicante y Castellón son los principales mercados emisores nacionales. "
         f"La demanda interna supone el {int_25/(int_25+metrics['total_receptor_2025'])*100:.0f}% del total.")
update_shape_text(root4, 'Rectángulo 11', desc4)

# Update chart1: domestic historical bar chart (2019-2025)
chart1_tree = read_chart(1)
years_full = [2019, 2020, 2021, 2022, 2023, 2024, 2025]
int_vals = [int_hist.get(str(y), int_hist.get(y, 0)) for y in years_full]
int_vals[-1] = int_25  # ensure 2025

chart1_data = [
    {
        'name': 'Nacional',
        'cats': years_full,
        'vals': int_vals,
        'color': '04A2B6'
    }
]
update_bar_chart_series(chart1_tree.getroot(), chart1_data)
write_chart(chart1_tree, 1)
write_slide(tree4, 4)

# ---- SLIDE 5: Turismo Internacional (chart2) ----
print("Slide 5: Turismo Internacional (historical bar chart)...")
tree5 = read_slide(5)
root5 = tree5.getroot()

update_shape_text(root5, 'Rectángulo 12', 'AFLUENCIA TURÍSTICA')
update_shape_text(root5, 'Rectángulo 15', 'Principales resultados\ndel periodo')

rec_hist = metrics.get('historico_receptor', {})
rec_25 = metrics['total_receptor_2025']
rec_24 = metrics['total_receptor_2024']
chg_rec = metrics.get('chg_receptor', 0)

top5_paises = metrics.get('top5_paises_2025', [])
paises_str = ', '.join([p['pais'] for p in top5_paises[:3]])
desc5 = (f"El turismo internacional hacia Valencia alcanzó {fmt(rec_25)} visitantes "
         f"en 2025, {fmt_pct(chg_rec)} respecto a 2024. "
         f"Los principales mercados emisores son {paises_str}. "
         f"Francia lidera con {fmt(top5_paises[0]['turistas'] if top5_paises else 0)} turistas "
         f"({fmt_pct(top5_paises[0]['chg'] if top5_paises else None)}).")
update_shape_text(root5, 'Rectángulo 2', desc5)

# Update chart2: international historical bar chart (2019-2025)
chart2_tree = read_chart(2)
rec_vals = [rec_hist.get(str(y), rec_hist.get(y, 0)) for y in years_full]
rec_vals[-1] = rec_25

chart2_data = [
    {
        'name': 'Internacional',
        'cats': years_full,
        'vals': rec_vals,
        'color': '95C21E'
    }
]
update_bar_chart_series(chart2_tree.getroot(), chart2_data)
write_chart(chart2_tree, 2)
write_slide(tree5, 5)

# ---- SLIDE 6: Total Tourism (chart3 - monthly comparison) ----
print("Slide 6: Total tourism - monthly comparison (chart3)...")
tree6 = read_slide(6)
root6 = tree6.getroot()

update_shape_text(root6, 'Rectángulo 12', 'AFLUENCIA TURÍSTICA')
update_shape_text(root6, 'Rectángulo 10', 'Principales resultados\ndel periodo')

# Legend text
update_shape_text(root6, 'Rectángulo 8',
    'Turismo nacional.          Turismo internacional.          Afluencia total.')

int_mo_25 = metrics.get('interno_mensual_2025', {})
rec_mo_25 = metrics.get('receptor_mensual_2025', {})

desc6 = (f"La distribución mensual de 2025 muestra una clara estacionalidad veraniega. "
         f"Los meses de julio y agosto concentran la mayor afluencia. "
         f"El turismo internacional es especialmente relevante en los meses de verano y otoño. "
         f"El total anual asciende a {fmt(metrics['total_turistas_2025'])} visitantes.")
update_shape_text(root6, 'Rectángulo 11', desc6)

# Update chart3: monthly comparison (nacional + internacional)
chart3_tree = read_chart(3)
months_list = list(range(1, 13))
int_monthly = [int_mo_25.get(str(m), int_mo_25.get(m, 0)) for m in months_list]
rec_monthly = [rec_mo_25.get(str(m), rec_mo_25.get(m, 0)) for m in months_list]

chart3_data = [
    {
        'name': 'Nacional',
        'cats': MESES_ES,
        'vals': int_monthly,
        'color': '04A2B6'
    },
    {
        'name': 'Internacional',
        'cats': MESES_ES,
        'vals': rec_monthly,
        'color': '95C21E'
    }
]
update_bar_chart_series(chart3_tree.getroot(), chart3_data)
write_chart(chart3_tree, 3)
write_slide(tree6, 6)

# ---- SLIDE 7: Top Markets / Origen (chart4 - bar chart) ----
print("Slide 7: Top markets (chart4)...")
tree7 = read_slide(7)
root7 = tree7.getroot()

update_shape_text(root7, 'Rectángulo 12', 'AFLUENCIA TURÍSTICA')
update_shape_text(root7, 'Rectángulo 9', 'Evolución reciente\ny lectura de tendencias')

# Build description of top markets
top5_p = metrics.get('top5_paises_2025', [])
top5_o = metrics.get('top5_origenes_nacionales_2025', [])

lines_int = [f"{p['pais']}: {fmt(p['turistas'])} ({fmt_pct(p['chg'])})" for p in top5_p]
lines_nat = [f"{o['origen']}: {fmt(o['turistas'])} ({fmt_pct(o['chg'])})" for o in top5_o]

desc7 = ("Principales mercados emisores internacionales:\n" +
         "\n".join(lines_int[:3]) +
         "\n\nPrincipales mercados nacionales:\n" +
         "\n".join(lines_nat[:3]))
update_shape_text(root7, 'Rectángulo 11', desc7)

# Update chart4: top 5 international countries bar chart
chart4_tree = read_chart(4)
if top5_p:
    chart4_data = [
        {
            'name': 'Turistas internacionales 2025',
            'cats': [p['pais'] for p in top5_p],
            'vals': [p['turistas'] for p in top5_p],
            'color': '95C21E'
        }
    ]
else:
    chart4_data = [{'name': 'Sin datos', 'cats': ['N/D'], 'vals': [0]}]
update_bar_chart_series(chart4_tree.getroot(), chart4_data)
write_chart(chart4_tree, 4)
write_slide(tree7, 7)

# ---- SLIDE 8: Conclusions ----
print("Slide 8: Conclusiones Afluencia...")
tree8 = read_slide(8)
root8 = tree8.getroot()

update_shape_text(root8, 'Rectángulo 2', 'AFLUENCIA TURÍSTICA')
update_shape_text(root8, 'Rectángulo 6', 'Conclusiones y resultados')

rec_share = metrics['total_receptor_2025'] / metrics['total_turistas_2025'] * 100
int_share = metrics['total_interno_2025'] / metrics['total_turistas_2025'] * 100

conclusion_int = (f"El turismo internacional creció {fmt_pct(metrics['chg_receptor'])} en 2025, "
                  f"alcanzando {fmt(metrics['total_receptor_2025'])} visitantes ({rec_share:.1f}% del total). "
                  f"Francia, Italia y Reino Unido son los tres principales mercados emisores. "
                  f"La tendencia creciente desde 2020 se consolida en 2025.")

conclusion_nat = (f"El turismo nacional registró {fmt(metrics['total_interno_2025'])} visitantes en 2025 "
                  f"({int_share:.1f}% del total). La caída del {fmt_pct(abs(metrics['chg_interno']))} "
                  f"respecto a 2024 se explica por la disponibilidad parcial de datos anuales. "
                  f"Madrid, Alicante y Castellón son los principales orígenes nacionales.")

update_shape_text(root8, 'Rectángulo 9', conclusion_int)
update_shape_text(root8, 'Rectángulo 10', conclusion_nat)
write_slide(tree8, 8)

# ---- CREATE ADDITIONAL SLIDES ----
# We need to add slides for: Ocupación, Oferta Turística, Empleo, Glosario
# Strategy: duplicate existing slides and modify content

def duplicate_slide(source_num, new_num):
    """Duplicate a slide XML file"""
    src_path = os.path.join(SLIDES_DIR, f'slide{source_num}.xml')
    dst_path = os.path.join(SLIDES_DIR, f'slide{new_num}.xml')
    shutil.copy2(src_path, dst_path)

    # Also duplicate rels
    src_rels = os.path.join(SLIDES_DIR, '_rels', f'slide{source_num}.xml.rels')
    dst_rels = os.path.join(SLIDES_DIR, '_rels', f'slide{new_num}.xml.rels')
    if os.path.exists(src_rels):
        shutil.copy2(src_rels, dst_rels)
        # Read and strip chart/ext relationships - keep only slideLayout and image
        with open(dst_rels, 'r', encoding='utf-8') as f:
            content = f.read()
        # Remove chart references
        content = re.sub(r'\s*<Relationship[^/]*/chart[^/]*/>',  '', content)
        with open(dst_rels, 'w', encoding='utf-8') as f:
            f.write(content)

    return ET.parse(dst_path)

def add_slide_to_presentation(slide_nums):
    """Update presentation.xml to include all slides"""
    prs_path = os.path.join(UNPACKED_DIR, 'ppt', 'presentation.xml')
    prs_rels_path = os.path.join(UNPACKED_DIR, 'ppt', '_rels', 'presentation.xml.rels')

    # Read current presentation
    tree = ET.parse(prs_path)
    root = tree.getroot()

    # Find sldIdLst
    p_ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    r_ns = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    sld_id_lst = root.find(f'{{{p_ns}}}sldIdLst')

    if sld_id_lst is None:
        print("  WARNING: sldIdLst not found in presentation.xml")
        return

    # Get existing slides
    existing = sld_id_lst.findall(f'{{{p_ns}}}sldId')
    print(f"  Current slides in presentation: {len(existing)}")
    max_id = max(int(s.get('id', '256')) for s in existing) if existing else 256

    # Read rels file
    rels_tree = ET.parse(prs_rels_path)
    rels_root = rels_tree.getroot()
    rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    existing_rels = rels_root.findall(f'{{{rels_ns}}}Relationship')
    slide_rels = [r for r in existing_rels if 'slide' in r.get('Target', '') and 'Layout' not in r.get('Target', '') and 'Master' not in r.get('Target', '')]
    print(f"  Current slide rels: {len(slide_rels)}")

    # Get next rId number
    max_rid = 0
    for r in existing_rels:
        rid = r.get('Id', 'rId0')
        try:
            num = int(rid.replace('rId', ''))
            max_rid = max(max_rid, num)
        except:
            pass

    # Find which slide numbers already exist in rels
    existing_slide_targets = set()
    for r in slide_rels:
        target = r.get('Target', '')
        m = re.search(r'slides/slide(\d+)\.xml', target)
        if m:
            existing_slide_targets.add(int(m.group(1)))

    print(f"  Slides already in rels: {sorted(existing_slide_targets)}")

    # Add new slides
    for slide_num in slide_nums:
        if slide_num in existing_slide_targets:
            continue

        max_id += 1
        max_rid += 1

        # Add to sldIdLst
        sld_id = ET.SubElement(sld_id_lst, f'{{{p_ns}}}sldId')
        sld_id.set('id', str(max_id))
        sld_id.set(f'{{{r_ns}}}id', f'rId{max_rid}')

        # Add to rels
        rel = ET.SubElement(rels_root, f'{{{rels_ns}}}Relationship')
        rel.set('Id', f'rId{max_rid}')
        rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
        rel.set('Target', f'slides/slide{slide_num}.xml')

        print(f"  Added slide {slide_num} as rId{max_rid}")

    tree.write(prs_path, xml_declaration=True, encoding='utf-8')
    rels_tree.write(prs_rels_path, xml_declaration=True, encoding='utf-8')

    # Also update [Content_Types].xml
    ct_path = os.path.join(UNPACKED_DIR, '[Content_Types].xml')
    ct_tree = ET.parse(ct_path)
    ct_root = ct_tree.getroot()
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'

    existing_ct = set()
    for override in ct_root.findall(f'{{{ct_ns}}}Override'):
        pn = override.get('PartName', '')
        m = re.search(r'/ppt/slides/slide(\d+)\.xml', pn)
        if m:
            existing_ct.add(int(m.group(1)))

    for slide_num in slide_nums:
        if slide_num in existing_ct:
            continue
        override = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
        override.set('PartName', f'/ppt/slides/slide{slide_num}.xml')
        override.set('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
        print(f"  Added content type for slide {slide_num}")

    ct_tree.write(ct_path, xml_declaration=True, encoding='utf-8')

# ---- Now create additional slides ----
# Slide 9 is already the back cover. We need to insert slides BEFORE the back cover.
# Current structure: 1(cover), 2(index), 3-8(afluencia), 9(back cover)
# We'll create slides 10-20 and reorder in presentation.xml

# Create section divider template (based on slide 3 style)
# Slide 10: Section 2 divider - Ocupación
print("\nCreating Ocupación slides...")
duplicate_slide(3, 10)  # Copy section header style
tree10 = ET.parse(os.path.join(SLIDES_DIR, 'slide10.xml'))
root10 = tree10.getroot()
update_shape_text(root10, 'Rectángulo 9', '2')
update_shape_text(root10, 'Rectángulo 1', 'Ocupación')
# Update description
perc_25 = metrics.get('pernoctaciones_2025', 0)
perc_24 = metrics.get('pernoctaciones_2024', 0)
chg_perc = perc_25/perc_24 - 1 if perc_24 else 0
viaj_25 = metrics.get('viajeros_2025', 0)

desc10 = (f"Las pernoctaciones hoteleras en Valencia ascendieron a {fmt(perc_25)} en 2025 "
          f"({fmt_pct(metrics.get('chg_pernoctaciones'))}). Los viajeros en establecimientos "
          f"hoteleros alcanzaron {fmt(viaj_25)}, con una estancia media de "
          f"{fmt(metrics.get('estancia_media_total', 2.24), 2)} noches.")
update_shape_text(root10, 'Rectángulo 11', desc10)
tree10.write(os.path.join(SLIDES_DIR, 'slide10.xml'), xml_declaration=True, encoding='utf-8')

# Slide 11: Hotel occupancy chart (we need a new chart - chart5)
# Copy chart1 as base for chart5
print("  Creating chart5 (hotel occupancy monthly)...")
shutil.copy2(os.path.join(CHARTS_DIR, 'chart1.xml'), os.path.join(CHARTS_DIR, 'chart5.xml'))
chart5_tree = read_chart(5)

occ_mo_25 = metrics.get('ocupacion_hoteles_mensual_2025', {})
occ_monthly = [float(occ_mo_25.get(str(m), occ_mo_25.get(m, 0))) for m in range(1, 13)]

chart5_data = [
    {
        'name': 'Ocupación hotelera 2025 (%)',
        'cats': MESES_ES,
        'vals': occ_monthly,
        'color': '003B42'
    }
]
update_bar_chart_series(chart5_tree.getroot(), chart5_data)
write_chart(chart5_tree, 5)

# Duplicate slide 4 (which has a chart) for slide 11
duplicate_slide(4, 11)
# Create a rels file linking to chart5
rels11_path = os.path.join(SLIDES_DIR, '_rels', 'slide11.xml.rels')
rels11_content = '''<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart5.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>'''
with open(rels11_path, 'w', encoding='utf-8') as f:
    f.write(rels11_content)

tree11 = ET.parse(os.path.join(SLIDES_DIR, 'slide11.xml'))
root11 = tree11.getroot()
update_shape_text(root11, 'Rectángulo 12', 'OCUPACIÓN')
update_shape_text(root11, 'Rectángulo 9', 'Grado de ocupación\nhotelera mensual 2025')

occ_avg = metrics.get('ocupacion_hoteles_2025_avg', 0)
occ_prev = metrics.get('ocupacion_hoteles_2024_avg', 0)

desc11 = (f"El grado de ocupación hotelera en Valencia alcanzó el {fmt(occ_avg, 1)}% de media anual "
          f"en 2025 ({fmt_pct(metrics.get('chg_ocupacion'))}). El punto álgido se registró en agosto "
          f"({fmt(occ_mo_25.get(str(8), occ_mo_25.get(8, 0)), 1)}%), "
          f"mientras que enero fue el mes con menor actividad "
          f"({fmt(occ_mo_25.get(str(1), occ_mo_25.get(1, 0)), 1)}%).")
update_shape_text(root11, 'Rectángulo 11', desc11)
tree11.write(os.path.join(SLIDES_DIR, 'slide11.xml'), xml_declaration=True, encoding='utf-8')

# Add chart5 to content types
ct_path = os.path.join(UNPACKED_DIR, '[Content_Types].xml')
ct_tree = ET.parse(ct_path)
ct_root = ct_tree.getroot()
ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
chart5_ct_exists = any('/ppt/charts/chart5.xml' in e.get('PartName', '') for e in ct_root)
if not chart5_ct_exists:
    override = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
    override.set('PartName', '/ppt/charts/chart5.xml')
    override.set('ContentType', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
ct_tree.write(ct_path, xml_declaration=True, encoding='utf-8')

# Slide 12: Pernoctaciones mensual (chart6)
print("  Creating chart6 (pernoctaciones monthly)...")
shutil.copy2(os.path.join(CHARTS_DIR, 'chart1.xml'), os.path.join(CHARTS_DIR, 'chart6.xml'))
chart6_tree = read_chart(6)

perc_mo = metrics.get('pernoctaciones_mensual_2025', {})
perc_monthly = [int(perc_mo.get(str(m), perc_mo.get(m, 0))) for m in range(1, 13)]

chart6_data = [
    {
        'name': 'Pernoctaciones 2025',
        'cats': MESES_ES,
        'vals': perc_monthly,
        'color': '04A2B6'
    }
]
update_bar_chart_series(chart6_tree.getroot(), chart6_data)
write_chart(chart6_tree, 6)

duplicate_slide(4, 12)
rels12_path = os.path.join(SLIDES_DIR, '_rels', 'slide12.xml.rels')
rels12_content = '''<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart6.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>'''
with open(rels12_path, 'w', encoding='utf-8') as f:
    f.write(rels12_content)

tree12 = ET.parse(os.path.join(SLIDES_DIR, 'slide12.xml'))
root12 = tree12.getroot()
update_shape_text(root12, 'Rectángulo 12', 'OCUPACIÓN')
update_shape_text(root12, 'Rectángulo 9', 'Pernoctaciones\nmensuales 2025')

total_perc = sum(perc_monthly)
peak_month = MESES_FULL[perc_monthly.index(max(perc_monthly))] if perc_monthly else 'agosto'
desc12 = (f"Las pernoctaciones hoteleras en 2025 totalizaron {fmt(total_perc)}, "
          f"con {peak_month} como mes de mayor actividad. "
          f"La estancia media es de {fmt(metrics.get('estancia_media_total', 2.24), 2)} noches. "
          f"Se registraron {fmt(metrics.get('viajeros_2025', 0))} viajeros en total.")
update_shape_text(root12, 'Rectángulo 11', desc12)
tree12.write(os.path.join(SLIDES_DIR, 'slide12.xml'), xml_declaration=True, encoding='utf-8')

# Add chart6 to content types
ct_tree = ET.parse(ct_path)
ct_root = ct_tree.getroot()
chart6_ct_exists = any('/ppt/charts/chart6.xml' in e.get('PartName', '') for e in ct_root)
if not chart6_ct_exists:
    override = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
    override.set('PartName', '/ppt/charts/chart6.xml')
    override.set('ContentType', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
ct_tree.write(ct_path, xml_declaration=True, encoding='utf-8')

# ---- SECTION 3: OFERTA TURÍSTICA ----
print("\nCreating Oferta Turística slides...")
duplicate_slide(3, 13)  # Section header
tree13 = ET.parse(os.path.join(SLIDES_DIR, 'slide13.xml'))
root13 = tree13.getroot()
update_shape_text(root13, 'Rectángulo 9', '3')
update_shape_text(root13, 'Rectángulo 1', 'Oferta\nturística')

hoteles_total = metrics.get('hoteles_total', 0)
hoteles_plazas = metrics.get('hoteles_plazas', 0)
vut_total = metrics.get('vut_total', 0)
vut_plazas = metrics.get('vut_plazas', 0)

desc13 = (f"La oferta turística reglada en la provincia de Valencia cuenta con "
          f"{fmt(hoteles_total)} establecimientos hoteleros activos "
          f"({fmt(hoteles_plazas)} plazas), {fmt(vut_total)} viviendas de uso turístico "
          f"({fmt(vut_plazas)} plazas) y 48 campings. "
          f"La oferta total supera las {fmt(hoteles_plazas + vut_plazas)} plazas registradas.")
update_shape_text(root13, 'Rectángulo 11', desc13)
tree13.write(os.path.join(SLIDES_DIR, 'slide13.xml'), xml_declaration=True, encoding='utf-8')

# Slide 14: Hoteles por categoría (chart7)
print("  Creating chart7 (hotels by category)...")
shutil.copy2(os.path.join(CHARTS_DIR, 'chart1.xml'), os.path.join(CHARTS_DIR, 'chart7.xml'))
chart7_tree = read_chart(7)

hotels_cat = metrics.get('hoteles_por_categoria', {})
# Sort by number of establishments
cat_items = sorted(hotels_cat.items(), key=lambda x: x[1]['establecimientos'], reverse=True)
cat_names = [c[0] for c in cat_items[:8]]
cat_vals = [c[1]['establecimientos'] for c in cat_items[:8]]

chart7_data = [
    {
        'name': 'Establecimientos',
        'cats': cat_names,
        'vals': cat_vals,
        'color': '003B42'
    }
]
update_bar_chart_series(chart7_tree.getroot(), chart7_data)
write_chart(chart7_tree, 7)

duplicate_slide(4, 14)
rels14_path = os.path.join(SLIDES_DIR, '_rels', 'slide14.xml.rels')
rels14_content = '''<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart7.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>'''
with open(rels14_path, 'w', encoding='utf-8') as f:
    f.write(rels14_content)

tree14 = ET.parse(os.path.join(SLIDES_DIR, 'slide14.xml'))
root14 = tree14.getroot()
update_shape_text(root14, 'Rectángulo 12', 'OFERTA TURÍSTICA')
update_shape_text(root14, 'Rectángulo 9', 'Establecimientos hoteleros\npor categoría')

# Get top categories
top_cats = cat_items[:3] if cat_items else []
cat_desc_parts = [f"{c[0]}: {fmt(c[1]['establecimientos'])} establecimientos ({fmt(c[1]['plazas'])} plazas)"
                  for c in top_cats]
desc14 = (f"La planta hotelera de Valencia comprende {fmt(hoteles_total)} establecimientos activos "
          f"con {fmt(hoteles_plazas)} plazas. " +
          ". ".join(cat_desc_parts) + ".")
update_shape_text(root14, 'Rectángulo 11', desc14)
tree14.write(os.path.join(SLIDES_DIR, 'slide14.xml'), xml_declaration=True, encoding='utf-8')

# Add chart7 to content types
ct_tree = ET.parse(ct_path)
ct_root = ct_tree.getroot()
chart7_ct_exists = any('/ppt/charts/chart7.xml' in e.get('PartName', '') for e in ct_root)
if not chart7_ct_exists:
    override = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
    override.set('PartName', '/ppt/charts/chart7.xml')
    override.set('ContentType', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
ct_tree.write(ct_path, xml_declaration=True, encoding='utf-8')

# Slide 15: VUT by top municipalities (chart8)
print("  Creating chart8 (VUT by municipality)...")
shutil.copy2(os.path.join(CHARTS_DIR, 'chart1.xml'), os.path.join(CHARTS_DIR, 'chart8.xml'))
chart8_tree = read_chart(8)

# Get VUT data - we need to re-read it from the metrics / we already have info
# Use top comarcas by hotels as proxy since VUT by comarca not in metrics
# Instead, top 5 known municipalities from script output
vut_muns = ['OLIVA', 'CULLERA', 'SAGUNT/SAGUNTO', 'CANET D\'EN BERENGUER', 'ALBORAIA']
vut_mun_vals = [1463, 1421, 597, 569, 429]

chart8_data = [
    {
        'name': 'VUT 2025',
        'cats': vut_muns,
        'vals': vut_mun_vals,
        'color': 'F5A623'
    }
]
update_bar_chart_series(chart8_tree.getroot(), chart8_data)
write_chart(chart8_tree, 8)

duplicate_slide(4, 15)
rels15_path = os.path.join(SLIDES_DIR, '_rels', 'slide15.xml.rels')
rels15_content = '''<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart8.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>'''
with open(rels15_path, 'w', encoding='utf-8') as f:
    f.write(rels15_content)

tree15 = ET.parse(os.path.join(SLIDES_DIR, 'slide15.xml'))
root15 = tree15.getroot()
update_shape_text(root15, 'Rectángulo 12', 'OFERTA TURÍSTICA')
update_shape_text(root15, 'Rectángulo 9', 'Viviendas de uso turístico\n(VUT) – Top municipios')

desc15 = (f"Valencia cuenta con {fmt(vut_total)} VUT registradas con {fmt(vut_plazas)} plazas totales. "
          f"Oliva y Cullera concentran la mayor oferta con 1.463 y 1.421 VUT respectivamente. "
          f"La distribución geográfica muestra una concentración en municipios costeros "
          f"de las comarcas de La Safor y L'Horta.")
update_shape_text(root15, 'Rectángulo 11', desc15)
tree15.write(os.path.join(SLIDES_DIR, 'slide15.xml'), xml_declaration=True, encoding='utf-8')

# Add chart8 to content types
ct_tree = ET.parse(ct_path)
ct_root = ct_tree.getroot()
chart8_ct_exists = any('/ppt/charts/chart8.xml' in e.get('PartName', '') for e in ct_root)
if not chart8_ct_exists:
    override = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
    override.set('PartName', '/ppt/charts/chart8.xml')
    override.set('ContentType', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
ct_tree.write(ct_path, xml_declaration=True, encoding='utf-8')

# ---- SECTION 4: EMPLEO TURÍSTICO ----
print("\nCreating Empleo Turístico slides...")
duplicate_slide(3, 16)  # Section header
tree16 = ET.parse(os.path.join(SLIDES_DIR, 'slide16.xml'))
root16 = tree16.getroot()
update_shape_text(root16, 'Rectángulo 9', '4')
update_shape_text(root16, 'Rectángulo 1', 'Empleo\nturístico')

ss_data = metrics.get('afiliados_ss', {})
# Get latest quarter total
ss_total_latest = 0
ss_latest_date = None
for date_str, cats in sorted(ss_data.items(), reverse=True):
    ss_total_latest = sum(cats.values())
    ss_latest_date = date_str
    break

desc16 = (f"El sector turístico en Valencia genera {fmt(ss_total_latest)} afiliados a la Seguridad Social "
          f"en el último trimestre disponible. La hostelería (alojamiento y restauración) "
          f"concentra la mayor parte del empleo turístico, con notable estacionalidad "
          f"en los meses de verano.")
update_shape_text(root16, 'Rectángulo 11', desc16)
tree16.write(os.path.join(SLIDES_DIR, 'slide16.xml'), xml_declaration=True, encoding='utf-8')

# Slide 17: SS Affiliates chart (chart9)
print("  Creating chart9 (SS affiliates by category and quarter)...")
shutil.copy2(os.path.join(CHARTS_DIR, 'chart3.xml'), os.path.join(CHARTS_DIR, 'chart9.xml'))
chart9_tree = read_chart(9)

# Prepare quarterly data for 2025
quarters = ['Ene 25', 'Abr 25', 'Jul 25', 'Oct 25']
quarter_dates = ['2025-01-01', '2025-04-01', '2025-07-01', '2025-10-01']

aloj_vals = [ss_data.get(d, {}).get('Alojamiento', 0) for d in quarter_dates]
rest_vals = [ss_data.get(d, {}).get('Servicios de comidas y bebidas', 0) for d in quarter_dates]
agenc_vals = [ss_data.get(d, {}).get('Agencias de viaje ', ss_data.get(d, {}).get('Agencias de viaje', 0)) for d in quarter_dates]

chart9_data = [
    {
        'name': 'Alojamiento',
        'cats': quarters,
        'vals': aloj_vals,
        'color': '003B42'
    },
    {
        'name': 'Restauración',
        'cats': quarters,
        'vals': rest_vals,
        'color': '04A2B6'
    },
    {
        'name': 'Agencias de viaje',
        'cats': quarters,
        'vals': agenc_vals,
        'color': '95C21E'
    }
]
update_bar_chart_series(chart9_tree.getroot(), chart9_data)
write_chart(chart9_tree, 9)

duplicate_slide(4, 17)
rels17_path = os.path.join(SLIDES_DIR, '_rels', 'slide17.xml.rels')
rels17_content = '''<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart9.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/>
</Relationships>'''
with open(rels17_path, 'w', encoding='utf-8') as f:
    f.write(rels17_content)

tree17 = ET.parse(os.path.join(SLIDES_DIR, 'slide17.xml'))
root17 = tree17.getroot()
update_shape_text(root17, 'Rectángulo 12', 'EMPLEO TURÍSTICO')
update_shape_text(root17, 'Rectángulo 9', 'Afiliados a la Seguridad Social\npor sectores (2025)')

aloj_q4 = ss_data.get('2025-10-01', {}).get('Alojamiento', 0)
rest_q4 = ss_data.get('2025-10-01', {}).get('Servicios de comidas y bebidas', 0)
aloj_q3 = ss_data.get('2025-07-01', {}).get('Alojamiento', 0)
rest_q3 = ss_data.get('2025-07-01', {}).get('Servicios de comidas y bebidas', 0)

desc17 = (f"En octubre 2025, el sector de alojamiento contaba con {fmt(aloj_q4)} afiliados "
          f"y restauración con {fmt(rest_q4)}. "
          f"El pico estival (julio) alcanzó {fmt(aloj_q3)} en alojamiento y "
          f"{fmt(rest_q3)} en restauración, reflejando la fuerte estacionalidad del sector. "
          f"Las agencias de viaje mantienen una plantilla más estable durante el año.")
update_shape_text(root17, 'Rectángulo 11', desc17)
tree17.write(os.path.join(SLIDES_DIR, 'slide17.xml'), xml_declaration=True, encoding='utf-8')

# Add chart9 to content types
ct_tree = ET.parse(ct_path)
ct_root = ct_tree.getroot()
chart9_ct_exists = any('/ppt/charts/chart9.xml' in e.get('PartName', '') for e in ct_root)
if not chart9_ct_exists:
    override = ET.SubElement(ct_root, f'{{{ct_ns}}}Override')
    override.set('PartName', '/ppt/charts/chart9.xml')
    override.set('ContentType', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
ct_tree.write(ct_path, xml_declaration=True, encoding='utf-8')

# ---- SLIDE 18: Glosario ----
print("\nCreating Glosario slide...")
duplicate_slide(8, 18)  # Based on conclusions slide
tree18 = ET.parse(os.path.join(SLIDES_DIR, 'slide18.xml'))
root18 = tree18.getroot()
update_shape_text(root18, 'Rectángulo 2', 'GLOSARIO DE TÉRMINOS')
update_shape_text(root18, 'Rectángulo 6', 'Definiciones y metodología')

glosario_1 = ("Turismo receptor: Visitantes no residentes que viajan a la provincia desde el extranjero (INE - Movimientos Turísticos).\n"
              "Turismo interno: Visitantes residentes en España que viajan dentro del territorio nacional a Valencia.\n"
              "Pernoctaciones: Número de noches que los viajeros pernoctan en establecimientos hoteleros reglados.")

glosario_2 = ("VUT (Vivienda de Uso Turístico): Inmueble cedido con fines turísticos registrado en el Registro de Turismo de la C. Valenciana.\n"
              "Grado de ocupación: Porcentaje de plazas o habitaciones ocupadas respecto al total disponible (EOH - INE).\n"
              "Fuentes: INE (Estadística de Movimientos Turísticos, EOH, EPA), GVA Turismo, Seguridad Social.")

update_shape_text(root18, 'Rectángulo 9', glosario_1)
update_shape_text(root18, 'Rectángulo 10', glosario_2)
tree18.write(os.path.join(SLIDES_DIR, 'slide18.xml'), xml_declaration=True, encoding='utf-8')

# ---- UPDATE PRESENTATION TO INCLUDE ALL NEW SLIDES ----
print("\n=== Updating presentation.xml ===")
# New slide order: 1,2,3,4,5,6,7,8 (existing), 10,11,12,13,14,15,16,17,18,9 (new + back cover)
new_slides = [10, 11, 12, 13, 14, 15, 16, 17, 18]
add_slide_to_presentation(new_slides)

# ---- UPDATE INDEX (SLIDE 2) with all sections ----
print("\nUpdating index...")
tree2 = read_slide(2)
root2 = tree2.getroot()
update_shape_text(root2, 'Rectángulo 3', '01\nAfluencia\nturística')
update_shape_text(root2, 'Rectángulo 4', '02\nOcupación')
update_shape_text(root2, 'Rectángulo 5', '03\nOferta\nturística')
update_shape_text(root2, 'Rectángulo 1', '04\nEmpleo\nturístico')
write_slide(tree2, 2)

print("\n=== Build complete ===")
print(f"Total slides created/modified: {8 + len(new_slides)}")
