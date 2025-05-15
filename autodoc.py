import json
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Configuration: ajustez ces chemins si besoin
INPUT_FILE = r'C:\Users\GLT\eiffage.com\OneDrive - eiffageenergie.be\Documents\Guillaume\Programmation\Auto. Doc. PBI\fichier_converti.json'
OUTPUT_FILE = r'C:\Users\GLT\eiffage.com\OneDrive - eiffageenergie.be\Documents\Guillaume\Programmation\Auto. Doc. PBI\documentation.docx'
EXCLUDE_KEYWORDS = ['LocalDateTable', 'DateTableTemplate']


def load_report(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def extract_metadata(report):
    tables = []
    calc_tables = []
    for tbl in report.get('model', {}).get('tables', []):
        name = tbl.get('name', '')
        # Exclusion
        if any(key in name for key in EXCLUDE_KEYWORDS):
            continue
        # Table calculée
        if tbl.get('type') == 'calculatedTable':
            expr = tbl.get('expression', '')
            if isinstance(expr, list):
                expr = '\n'.join(expr)
            calc_tables.append({'name': name, 'expression': expr})
            continue
        # Table standard
        info = {
            'name': name,
            'measures': [],
            'calculated_columns': [],
            'partitions': [],
            'hierarchies': []
        }
        # Mesures
        for m in tbl.get('measures', []):
            expr = m.get('expression', '')
            if isinstance(expr, list): expr = '\n'.join(expr)
            info['measures'].append({'name': m.get('name'), 'expression': expr})
        # Colonnes calculées
        for col in tbl.get('columns', []):
            if col.get('type') in ['calculated', 'calculatedTableColumn']:
                expr = col.get('expression', '')
                if isinstance(expr, list): expr = '\n'.join(expr)
                info['calculated_columns'].append({'name': col.get('name'), 'expression': expr})
        # Partitions (incluant type et expression)
        for part in tbl.get('partitions', []):
            src = part.get('source', {})
            src_type = src.get('type', '')
            src_expr = src.get('expression', '') if src_type == 'calculated' else ''
            info['partitions'].append({
                'name': part.get('name'),
                'mode': part.get('mode'),
                'source_type': src_type,
                'source_expression': src_expr
            })
        # Hiérarchies
        for hier in tbl.get('hierarchies', []):
            levels = [lvl.get('name') for lvl in hier.get('levels', [])]
            info['hierarchies'].append({'name': hier.get('name'), 'levels': levels})
        tables.append(info)
    return tables, calc_tables


def style_table(table, headers, rows):
    # En-têtes
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        cell = hdr[i]
        cell.text = h
        run = cell.paragraphs[0].runs[0]
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(255, 255, 255)
        tcPr = cell._element.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), 'FF0000'); tcPr.append(shd)
        borders = OxmlElement('w:tcBorders')
        for n in ['top','left','bottom','right','insideH','insideV']:
            bd = OxmlElement(f'w:{n}'); bd.set(qn('w:val'),'single'); bd.set(qn('w:sz'),'4'); bd.set(qn('w:space'),'0'); bd.set(qn('w:color'),'000000'); borders.append(bd)
        tcPr.append(borders)
    # Données
    for idx, row in enumerate(rows):
        cells = table.add_row().cells
        for j, val in enumerate(row):
            cell = cells[j]
            cell.text = ''
            for k, line in enumerate(str(val).split('\n')):
                if k == 0:
                    p = cell.paragraphs[0]
                else:
                    p = cell.add_paragraph()
                run = p.add_run(line)
                run.font.size = Pt(8)
        # Style ligne
        bg = 'FFFFFF' if idx % 2 == 0 else 'D3D3D3'
        for cell in cells:
            tcPr = cell._element.get_or_add_tcPr()
            shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), bg); tcPr.append(shd)
            borders = OxmlElement('w:tcBorders')
            for n in ['top','left','bottom','right','insideH','insideV']:
                bd = OxmlElement(f'w:{n}'); bd.set(qn('w:val'),'single'); bd.set(qn('w:sz'),'4'); bd.set(qn('w:space'),'0'); bd.set(qn('w:color'),'000000'); borders.append(bd)
            tcPr.append(borders)


def generate_word(tables, calc_tables, output_path):
    doc = Document()
    doc.add_heading('Documentation automatique du modèle Power BI', level=0)
    # Tables calculées
    if calc_tables:
        doc.add_heading('Tables calculées', level=1)
        t = doc.add_table(rows=1, cols=2)
        style_table(t, ['Nom','Expression'], [(c['name'], c['expression']) for c in calc_tables])
        doc.add_paragraph()
    # Tables usuelles
    for tbl in tables:
        doc.add_heading(f"Table: {tbl['name']}", level=1)
        if tbl['measures']:
            doc.add_paragraph('Mesures:', style='Intense Quote')
            t = doc.add_table(rows=1, cols=2)
            style_table(t, ['Nom','Expression'], [(m['name'], m['expression']) for m in tbl['measures']])
            doc.add_paragraph()
        if tbl['calculated_columns']:
            doc.add_paragraph('Colonnes calculées:', style='Intense Quote')
            t = doc.add_table(rows=1, cols=2)
            style_table(t, ['Nom','Expression'], [(c['name'], c['expression']) for c in tbl['calculated_columns']])
            doc.add_paragraph()
        if tbl['partitions']:
            doc.add_paragraph('Partitions:', style='Intense Quote')
            t = doc.add_table(rows=1, cols=4)
            headers = ['Nom','Mode','Type source','Expression source']
            rows = [(p['name'], p['mode'], p['source_type'], p['source_expression']) for p in tbl['partitions']]
            style_table(t, headers, rows)
            doc.add_paragraph()
        if tbl['hierarchies']:
            doc.add_paragraph('Hiérarchies:', style='Intense Quote')
            for hier in tbl['hierarchies']:
                doc.add_paragraph(f"{hier['name']}: {', '.join(hier['levels'])}")
            doc.add_paragraph()
    doc.save(output_path)
    print(f"Documentation Word générée: {output_path}")


if __name__ == '__main__':
    report = load_report(INPUT_FILE)
    tables, calc_tables = extract_metadata(report)
    generate_word(tables, calc_tables, OUTPUT_FILE)