"""Check PPTX relationships, slide order, glued connectors and chart/XLSX data.

This is a structural and numerical check. It does not certify text fit or visual
quality; render every slide and inspect it in the intended application as well.
"""
from __future__ import annotations
import argparse
from decimal import Decimal
import hashlib
import io
import json
from pathlib import Path
import posixpath
import re
import zipfile
from lxml import etree as E
import ooxml as O

NS = {**O.NS, 's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def source_part(relpath):
    if relpath == '_rels/.rels':
        return ''
    parent = posixpath.dirname(posixpath.dirname(relpath))
    return posixpath.join(parent, posixpath.basename(relpath)[:-5])


def resolve(origin, target):
    return target[1:] if target.startswith('/') else O.resolve(origin, target)


def validate(path, expected_slides=None):
    errors = []; warnings = []
    with zipfile.ZipFile(path) as archive:
        if archive.testzip():
            errors.append('ZIP CRC check failed')
        files = {n: archive.read(n) for n in archive.namelist()}
    xml = {}
    for name, raw in files.items():
        if name.endswith(('.xml', '.rels')):
            try:
                xml[name] = E.fromstring(raw)
            except E.XMLSyntaxError as exc:
                errors.append(f'{name}: {exc}')
    for name, root in xml.items():
        if not name.endswith('.rels'):
            continue
        origin = source_part(name)
        ids = [r.get('Id') for r in root]
        if len(ids) != len(set(ids)):
            errors.append(f'Duplicate relationship ID in {name}')
        for rel in root:
            if rel.get('TargetMode') == 'External':
                continue
            destination = resolve(origin, rel.get('Target'))
            if destination not in files:
                errors.append(f'Broken relationship: {name} -> {destination}')
    pres = xml['ppt/presentation.xml']
    rels = {r.get('Id'): r for r in xml['ppt/_rels/presentation.xml.rels']}
    slides = [resolve('ppt/presentation.xml', rels[s.get(O.tag('r', 'id'))].get('Target')) for s in pres.findall('p:sldIdLst/p:sldId', NS)]
    if expected_slides is not None and len(slides) != expected_slides:
        errors.append(f'Expected {expected_slides} slides, found {len(slides)}')
    tables = []; charts = []; connectors = 0; notes_count = 0
    for number, name in enumerate(slides, 1):
        slide = xml[name]
        ids = [s.get('id') for s in slide.findall('.//p:cNvPr', NS)]
        if len(ids) != len(set(ids)):
            errors.append(f'Slide {number}: duplicate shape IDs')
        for connector in slide.findall('.//p:cxnSp', NS):
            nv = connector.find('p:nvCxnSpPr/p:cNvPr', NS)
            if nv.get('name') != 'Attached process connector':
                continue
            endpoints = connector.findall('p:nvCxnSpPr/p:cNvCxnSpPr/*', NS)
            if len(endpoints) != 2 or any(p.get('id') not in ids for p in endpoints):
                errors.append(f'Slide {number}: connector endpoint is not attached')
            connectors += 1
        if slide.findall('.//a:tbl', NS):
            tables.append(number)
        slide_rels = {r.get('Id'): r for r in xml[O.relpath(name)]}
        for rel in slide_rels.values():
            if rel.get('Type', '').endswith('/notesSlide'):
                notes_count += 1
                note = xml[resolve(name, rel.get('Target'))]
                bodies = [s for s in note.findall('.//p:sp', NS) if s.find("p:nvSpPr/p:nvPr/p:ph[@type='body']", NS) is not None]
                if not bodies or not ''.join(bodies[0].xpath('.//a:t/text()', namespaces=NS)).strip():
                    warnings.append(f'Slide {number}: empty speaker notes')
        for ref in slide.findall('.//c:chart', NS):
            chart_path = resolve(name, slide_rels[ref.get(O.tag('r', 'id'))].get('Target'))
            chart = xml[chart_path]
            chart_rels = {r.get('Id'): r for r in xml[O.relpath(chart_path)]}
            ext = chart.find('c:externalData', NS)
            if ext is None:
                errors.append(f'Slide {number}: chart has no embedded workbook'); continue
            book = resolve(chart_path, chart_rels[ext.get(O.tag('r', 'id'))].get('Target'))
            with zipfile.ZipFile(io.BytesIO(files[book])) as xlsx:
                sheet = E.fromstring(xlsx.read('xl/worksheets/sheet1.xml'))
                shared = []
                if 'xl/sharedStrings.xml' in xlsx.namelist():
                    shared = [''.join(si.xpath('.//s:t/text()', namespaces=NS)) for si in E.fromstring(xlsx.read('xl/sharedStrings.xml'))]
                cells = {}
                for cell in sheet.findall('.//s:c', NS):
                    value = cell.findtext('s:v', namespaces=NS)
                    if cell.get('t') == 'inlineStr':
                        value = ''.join(cell.xpath('.//s:t/text()', namespaces=NS))
                    elif cell.get('t') == 's':
                        value = shared[int(value)]
                    cells[cell.get('r')] = value
            checked = 0
            for ref_kind in ('numRef', 'strRef'):
                for reference in chart.findall('.//c:'+ref_kind, NS):
                    formula = reference.findtext('c:f', namespaces=NS)
                    m = re.fullmatch(r"[^!]+!\$([A-Z]+)\$(\d+)(?::\$\1\$(\d+))?", formula or '')
                    if not m:
                        errors.append(f'Slide {number}: unsupported chart formula {formula}'); continue
                    col, start, end = m.group(1), int(m.group(2)), int(m.group(3) or m.group(2))
                    cache = reference.find('c:'+ref_kind[:3]+'Cache', NS)
                    if cache is None:
                        errors.append(f'Slide {number}: missing cache'); continue
                    points = {start+int(p.get('idx')): p.findtext('c:v', namespaces=NS) for p in cache.findall('c:pt', NS)}
                    if int(cache.find('c:ptCount', NS).get('val')) != end-start+1:
                        errors.append(f'Slide {number}: inconsistent cache point count')
                    for row in range(start, end+1):
                        expected = cells.get(f'{col}{row}'); actual = points.get(row)
                        if ref_kind == 'numRef' and expected is not None and actual is not None:
                            matches = Decimal(actual) == Decimal(expected)
                        else:
                            matches = actual == expected
                        if not matches:
                            errors.append(f'Slide {number}: chart cache does not match worksheet {col}{row}')
                        checked += 1
            charts.append({'slide': number, 'chart': chart_path, 'workbook': book, 'cached_cells_checked': checked})
    if notes_count != len(slides):
        errors.append('Not every slide has a notes part')
    generated = all(re.fullmatch(r'ppt/slides/reportSlide\d+\.xml', name) for name in slides)
    metadata_status = 'not checked for external reference decks'
    if generated:
        metadata_status = 'checked generated-deck counts and removed inherited statistics'
        app = xml.get('docProps/app.xml')
        ep = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
        if app is None:
            errors.append('Generated report has no application-properties part')
        else:
            hidden = sum(xml[name].get('show') == '0' for name in slides)
            for label, expected in (('Slides', len(slides)), ('Notes', notes_count), ('HiddenSlides', hidden)):
                if app.findtext('{'+ep+'}'+label) != str(expected):
                    errors.append(f'Application metadata {label} does not match generated parts')
            for label in ('Words', 'Paragraphs', 'HeadingPairs', 'TitlesOfParts'):
                if app.find('{'+ep+'}'+label) is not None:
                    errors.append(f'Generated report retains inherited {label} metadata')
    return {'status': 'passed' if not errors else 'failed', 'pptx': path.name,
            'pptx_sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'slide_count': len(slides),
            'native_table_slides': tables, 'charts': charts, 'attached_connectors': connectors,
            'notes_slides': notes_count, 'errors': errors, 'warnings': warnings,
            'application_metadata': metadata_status,
            'visual_review': 'not performed by this script', 'application_editability_test': 'not performed by this script'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pptx', type=Path)
    parser.add_argument('--expected-slides', type=int)
    parser.add_argument('--output', type=Path)
    args = parser.parse_args()
    report = validate(args.pptx.resolve(), args.expected_slides)
    value = json.dumps(report, ensure_ascii=False, indent=2)+'\n'
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(value, encoding='utf8')
    print(value)
    raise SystemExit(0 if report['status']=='passed' else 1)


if __name__ == '__main__':
    main()
