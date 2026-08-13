import zipfile
import re
import sys
import os
import xml.etree.ElementTree as ET

DOCX_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XLSX_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
RELS_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
DOCX_STYLE = {
    'Heading1': '# ', 'Heading2': '## ', 'Heading3': '### ',
    'Heading4': '#### ', 'Heading5': '##### ', 'Heading6': '###### ',
    'ListBullet': '- ', 'ListNumber': '1. ',
}

def get_docx_runs(para):
    runs = para.findall(f'.//{DOCX_NS}r')
    result = ''
    for r in runs:
        text = ''
        for t in r.findall(f'.//{DOCX_NS}t'):
            if t.text:
                text += t.text
        if not text:
            continue
        rpr = r.find(f'{DOCX_NS}rPr')
        bold = rpr is not None and rpr.find(f'{DOCX_NS}b') is not None
        italic = rpr is not None and rpr.find(f'{DOCX_NS}i') is not None
        underline = rpr is not None and rpr.find(f'{DOCX_NS}u') is not None
        if bold:
            text = f'**{text}**'
        if italic:
            text = f'*{text}*'
        if underline:
            text = f'__{text}__'
        result += text
    return result

def convert_docx(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        tree = ET.parse(z.open('word/document.xml'))
    paragraphs = tree.getroot().findall(f'.//{DOCX_NS}body/{DOCX_NS}p')
    lines = []
    for para in paragraphs:
        ppr = para.find(f'{DOCX_NS}pPr')
        style = None
        if ppr is not None:
            ps = ppr.find(f'{DOCX_NS}pStyle')
            if ps is not None:
                style = ps.get(f'{DOCX_NS}val')
        text = get_docx_runs(para).strip()
        if not text:
            lines.append('')
            continue
        prefix = DOCX_STYLE.get(style, '')
        if style and style.startswith('Heading'):
            lines.append(prefix + text)
        elif style == 'ListBullet':
            lines.append('- ' + text)
        elif style == 'ListNumber':
            lines.append('1. ' + text)
        else:
            lines.append(text)
    return '\n'.join(lines)

def convert_xlsx(xlsx_path):
    with zipfile.ZipFile(xlsx_path) as z:
        file_list = [f.filename for f in z.filelist]
        shared_strings = []
        if 'xl/sharedStrings.xml' in file_list:
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            for si in tree.findall(f'.//{XLSX_NS}si'):
                parts = []
                for t in si.findall(f'.//{XLSX_NS}t'):
                    if t.text:
                        parts.append(t.text)
                shared_strings.append(''.join(parts))

        wb_tree = ET.parse(z.open('xl/workbook.xml'))
        sheets = []
        for s in wb_tree.findall(f'.//{XLSX_NS}sheet'):
            sheets.append({
                'name': s.get('name'),
                'rid': s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            })

        rels_tree = ET.parse(z.open('xl/_rels/workbook.xml.rels'))
        rid_map = {}
        for rel in rels_tree.findall(f'.//{RELS_NS}Relationship'):
            rid_map[rel.get('Id')] = rel.get('Target')

        all_md = []
        for sheet in sheets:
            target = rid_map.get(sheet['rid'])
            if not target:
                continue
            path = 'xl/' + target
            if path not in file_list:
                continue
            sheet_tree = ET.parse(z.open(path))
            rows_data = []
            for row in sheet_tree.findall(f'.//{XLSX_NS}row'):
                cells = []
                for c in row.findall(f'{XLSX_NS}c'):
                    t = c.get('t')
                    v_el = c.find(f'{XLSX_NS}v')
                    val = v_el.text if v_el is not None and v_el.text else ''
                    if t == 's':
                        idx = int(val) if val else 0
                        val = shared_strings[idx] if idx < len(shared_strings) else ''
                    elif t == 'b':
                        val = 'TRUE' if val == '1' else 'FALSE'
                    cells.append(val)
                rows_data.append(cells)

            if not rows_data:
                continue
            if len(sheets) > 1:
                all_md.append(f'## {sheet["name"]}\n')

            max_cols = max(len(r) for r in rows_data)
            for r in rows_data:
                while len(r) < max_cols:
                    r.append('')

            header = rows_data[0]
            all_md.append('| ' + ' | '.join(header) + ' |')
            all_md.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
            for r in rows_data[1:]:
                if all(c.strip() == '' for c in r):
                    continue
                all_md.append('| ' + ' | '.join(r) + ' |')
            all_md.append('')

    return '\n'.join(all_md)

def convert_pdf(pdf_path):
    try:
        from pdfminer.high_level import extract_text
        return extract_text(pdf_path)
    except ImportError:
        return f'[PDF: pip install pdfminer.six]\nFile: {pdf_path}'

def main():
    if len(sys.argv) < 2:
        print('Usage: python docx2md.py <input.docx|xlsx|pdf> [output.md]')
        sys.exit(1)
    src = sys.argv[1]
    if not os.path.isfile(src):
        print(f'File not found: {src}')
        sys.exit(1)
    ext = os.path.splitext(src)[1].lower()
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + '.md'

    if ext == '.docx':
        md = convert_docx(src)
    elif ext == '.xlsx':
        md = convert_xlsx(src)
    elif ext == '.pdf':
        md = convert_pdf(src)
    else:
        print(f'Unsupported: {ext}')
        sys.exit(1)

    with open(dst, 'w', encoding='utf-8') as f:
        f.write(md)
    print(f'Created: {dst}')

if __name__ == '__main__':
    main()
