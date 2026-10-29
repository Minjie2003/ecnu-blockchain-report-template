"""Build an editable ECNU progress report from a JSON outline and frozen PPTX.

Example: python scripts/build_report.py --content examples/starter.json
         --output outputs/my-report.pptx
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
from pathlib import Path
import zipfile
from lxml import etree as E
import ooxml as O

ROOT = Path(__file__).resolve().parents[1]
NS, EMU = O.NS, O.EMU
INK, BLUE, RED = '263747', '245680', 'A23838'
PALETTE = {'blue': ('E9F2FA', BLUE), 'orange': ('FFF2DF', 'A46A25'),
           'green': ('ECF4EA', '3D7244'), 'red': ('FCECEA', RED)}


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def text(slide, value, x, y, w, h, size=22, color=INK, bold=False, center=False):
    shape = O.textbox(slide, value, x, y, w, h, size, color, bold, center,
                      '宋体', 'Editable report text')
    for spacing in shape.findall('.//a:lnSpc/a:spcPct', NS):
        spacing.set('val', '110000')
    return shape


def clean(slide, title, major=False):
    """Retain the reference header and original artwork; replace body content."""
    tree = slide.find('p:cSld/p:spTree', NS)
    top = None
    for shape in list(tree):
        if E.QName(shape).localname in {'nvGrpSpPr', 'grpSpPr'}:
            continue
        xf = shape.find('p:spPr/a:xfrm', NS)
        if xf is None:
            xf = shape.find('p:xfrm', NS)
        if xf is None:
            continue
        y = int(xf.find('a:off', NS).get('y')) / EMU
        if y > 1.31:
            tree.remove(shape)
        elif shape.find('p:txBody', NS) is not None:
            if top is None:
                top = shape
            else:
                tree.remove(shape)
    if top is None:
        raise ValueError('Selected reference slide has no editable title')
    O.text_style(top, title, 40 if major else 28, '000000', major, major,
                 '微软雅黑' if major else '宋体', 0)
    O.move(top, .65, .20, 12 if major else 9, .76 if major else .64)


def module(slide, value, x, y, w, h, color='blue'):
    fill, stroke = PALETTE[color]
    shape = text(slide, value, x, y, w, h, 21, center=True)
    properties = shape.find('p:spPr', NS)
    properties.remove(properties.find('a:noFill', NS))
    O.sub(O.sub(properties, 'a', 'solidFill'), 'a', 'srgbClr', val=fill)
    line = properties.find('a:ln', NS)
    line.clear(); line.set('w', '12700')
    O.sub(O.sub(line, 'a', 'solidFill'), 'a', 'srgbClr', val=stroke)
    body = shape.find('p:txBody/a:bodyPr', NS)
    body.set('anchor', 'ctr')
    for name in ('lIns', 'rIns', 'tIns', 'bIns'):
        body.set(name, str(round(.06 * EMU)))
    return shape


def connect(slide, source, target):
    """Right-to-left native connector with glued source and target endpoints."""
    tree = slide.find('p:cSld/p:spTree', NS)
    def point(shape, end):
        xf = shape.find('p:spPr/a:xfrm', NS)
        off, ext = xf.find('a:off', NS), xf.find('a:ext', NS)
        x, y, w, h = [int(e.get(k)) for e, k in [(off, 'x'), (off, 'y'), (ext, 'cx'), (ext, 'cy')]]
        return (x + w if end else x), y + h // 2
    x1, y1 = point(source, True); x2, y2 = point(target, False)
    ident = max(int(n.get('id')) for n in tree.findall('.//p:cNvPr', NS)) + 1
    connector = O.sub(tree, 'p', 'cxnSp')
    nv = O.sub(connector, 'p', 'nvCxnSpPr')
    O.sub(nv, 'p', 'cNvPr', id=ident, name='Attached process connector')
    props = O.sub(nv, 'p', 'cNvCxnSpPr')
    for label, shape, site in [('stCxn', source, 3), ('endCxn', target, 1)]:
        O.sub(props, 'a', label, id=shape.find('p:nvSpPr/p:cNvPr', NS).get('id'), idx=site)
    O.sub(nv, 'p', 'nvPr')
    properties = O.sub(connector, 'p', 'spPr')
    xf = O.sub(properties, 'a', 'xfrm')
    O.sub(xf, 'a', 'off', x=x1, y=y1)
    O.sub(xf, 'a', 'ext', cx=x2-x1, cy=y2-y1)
    O.sub(O.sub(properties, 'a', 'prstGeom', prst='line'), 'a', 'avLst')
    line = O.sub(properties, 'a', 'ln', w=15240)
    O.sub(O.sub(line, 'a', 'solidFill'), 'a', 'srgbClr', val=BLUE)
    O.sub(line, 'a', 'tailEnd', type='triangle')
    tree.remove(connector); tree.insert(2, connector)


def body(slide, spec, parts, relationships, index):
    kind = spec['layout']
    clean(slide, spec['title'], kind in {'goals', 'reflection', 'plan'})
    if kind in {'goals', 'plan'}:
        blocks = spec['blocks']
        if not 1 <= len(blocks) <= 3:
            raise ValueError('Use 1–3 goal/plan blocks or split the slide')
        for i, block in enumerate(blocks):
            y = 1.73 + i*1.62
            text(slide, block['label'], .9, y, 11.55, .43, 22, RED, True)
            text(slide, block['body'], .9, y+.51, 11.55, 1, 22)
    elif kind in {'text', 'reflection'}:
        paragraphs = spec['paragraphs']
        if not 1 <= len(paragraphs) <= 3:
            raise ValueError('Use 1–3 paragraphs or split the slide')
        for i, paragraph in enumerate(paragraphs):
            text(slide, paragraph, .9, 1.85 + i*1.70, 11.55, 1.42, 22)
    elif kind == 'table':
        text(slide, spec.get('lead', ''), .83, 1.45, 11.7, .8)
        rows = [spec['columns']] + spec['rows']
        widths = spec.get('widths', [11.7/len(spec['columns'])]*len(spec['columns']))
        if sum(widths) > 11.75 or len(rows) > 6:
            raise ValueError('Table exceeds reference layout; split or reduce rows')
        O.native_table(slide, rows, widths, y=2.55, height=3.15, size=20)
        text(slide, spec.get('takeaway', ''), .85, 6.16, 11.7, .7, 22, BLUE, True)
    elif kind == 'process':
        nodes = spec['nodes']
        if not 2 <= len(nodes) <= 4:
            raise ValueError('Process layout supports 2–4 modules in one row')
        text(slide, spec.get('lead', ''), .83, 1.45, 11.7, .85)
        gap = .62; width = (11.55-gap*(len(nodes)-1))/len(nodes)
        shapes = [module(slide, n['label'], .9+i*(width+gap), 2.9, width, 1.15, n.get('color', 'blue')) for i, n in enumerate(nodes)]
        for a, b in zip(shapes, shapes[1:]):
            connect(slide, a, b)
        text(slide, spec.get('explanation', ''), .9, 4.7, 11.55, 1.05)
        text(slide, spec.get('takeaway', ''), .9, 6.2, 11.55, .65, 22, BLUE, True)
    elif kind == 'algorithm':
        text(slide, spec.get('lead', ''), .85, 1.45, 11.65, .8, 21)
        text(slide, spec['algorithm_title'], .85, 2.4, 11.65, .6, 22, BLUE, True)
        steps = spec['steps']
        if not 1 <= len(steps) <= 7:
            raise ValueError('Algorithm layout supports 1–7 steps; split long algorithms')
        for i, step in enumerate(steps):
            text(slide, str(i+1), .92, 3.08+i*.40, .4, .35, 20, BLUE)
            text(slide, step, 1.5, 3.08+i*.40, 10.8, .35, 20)
        text(slide, spec.get('takeaway', ''), .85, 6.35, 11.65, .64, 20)
    elif kind == 'chart':
        chart = copy.deepcopy(spec['chart'])
        if not spec.get('source'):
            raise ValueError('Every chart requires a source or explicit demonstration label')
        if spec.get('example_data'):
            text(slide, '仅示例数据，不是实验结果', .85, 1.43, 11.65, .45, 22, RED, True)
        else:
            text(slide, spec.get('lead', ''), .85, 1.43, 11.65, .6, 20)
        chart_no = 1000+index
        O.native_chart(slide, parts, chart_no, chart, geometry=(.82, 2.13, 11.7, 3.9))
        O.sub(relationships, 'pr', 'Relationship', Id=f'rIdE16Chart{chart_no}', Type=O.REL+'chart', Target=f'../charts/chart{chart_no}.xml')
        text(slide, spec.get('takeaway', ''), .85, 6.23, 11.7, .60, 21, BLUE, True)
        text(slide, '来源：'+spec['source'], .85, 6.92, 11.7, .3, 15)
    else:
        raise ValueError(f'Unsupported layout: {kind}')


def prune(parts):
    keep = {'[Content_Types].xml', '_rels/.rels'}; pending = ['']
    while pending:
        part = pending.pop(); relationship = O.relpath(part) if part else '_rels/.rels'
        if relationship not in parts:
            continue
        keep.add(relationship)
        for rel in E.fromstring(parts[relationship]):
            if rel.get('TargetMode') == 'External':
                continue
            target = rel.get('Target'); target = target[1:] if target.startswith('/') else O.resolve(part, target)
            if target not in parts:
                raise ValueError(f'Broken relationship: {part} -> {target}')
            if target not in keep:
                keep.add(target); pending.append(target)
    parts = {name: value for name, value in parts.items() if name in keep}
    types = E.fromstring(parts['[Content_Types].xml'])
    for child in list(types):
        if child.get('PartName') and child.get('PartName')[1:] not in keep:
            types.remove(child)
    parts['[Content_Types].xml'] = O.xml(types)
    return parts


def build(content_path, output, template):
    content = json.loads(content_path.read_text(encoding='utf8'))
    slides = content['slides']
    if not slides:
        raise ValueError('The report must contain at least one slide')
    parts = O.read_package(template)
    presentation = E.fromstring(parts['ppt/presentation.xml'])
    rels = E.fromstring(parts['ppt/_rels/presentation.xml.rels'])
    mapping = {r.get('Id'): r for r in rels}
    ids = list(presentation.find('p:sldIdLst', NS))
    source_paths = [O.resolve('ppt/presentation.xml', mapping[s.get(O.tag('r', 'id'))].get('Target')) for s in ids]
    # The source is the approved 12-slide report. Its cover, goals, content and
    # closing pages provide four distinct template roles.
    roles = {'cover': 0, 'goals': 1, 'plan': 1, 'reflection': 1, 'closing': 11}
    for element in ids:
        element.getparent().remove(element)
    for relationship in list(rels):
        if relationship.get('Type', '').endswith('/slide'):
            rels.remove(relationship)
    types = E.fromstring(parts['[Content_Types].xml'])
    namespaces = {'dc': 'http://purl.org/dc/elements/1.1/'}
    if 'docProps/core.xml' in parts:
        core = E.fromstring(parts['docProps/core.xml'])
        title = core.find('dc:title', namespaces)
        if title is not None:
            title.text = content.get('title', '研究进展汇报')
        parts['docProps/core.xml'] = O.xml(core)
    for index, spec in enumerate(slides, 1):
        kind = spec['layout']; origin = source_paths[roles.get(kind, 2)]
        slide = E.fromstring(parts[origin]); slide_rels = E.fromstring(parts[O.relpath(origin)])
        path = f'ppt/slides/reportSlide{index}.xml'
        notes_path = f'ppt/notesSlides/reportNotes{index}.xml'
        notes_rel = next(r for r in slide_rels if r.get('Type', '').endswith('/notesSlide'))
        old_notes_path = O.resolve(origin, notes_rel.get('Target'))
        notes = E.fromstring(parts[old_notes_path]); notes_rels = E.fromstring(parts[O.relpath(old_notes_path)])
        notes_rel.set('Target', '../notesSlides/'+Path(notes_path).name)
        next(r for r in notes_rels if r.get('Type', '').endswith('/slide')).set('Target', '../slides/'+Path(path).name)
        if kind == 'cover':
            shape = next(s for s in slide.findall('p:cSld/p:spTree/p:sp', NS) if s.find('p:txBody', NS) is not None)
            lines = [content.get('report_label', '研究进展汇报'), content['title'], content.get('author', '汇报人')+'  '+content.get('date', '日期')]
            O.text_style(shape, lines, 32, '000000', True, True, '微软雅黑', 22)
            shape.find('p:txBody/a:bodyPr', NS).set('anchor', 'ctr')
        elif kind == 'closing':
            for shape in list(slide.findall('p:cSld/p:spTree/p:sp', NS)):
                value = ''.join(shape.xpath('.//a:t/text()', namespaces=NS))
                if value.isdigit():
                    shape.getparent().remove(shape)
        else:
            body(slide, spec, parts, slide_rels, index)
        if index > 1:
            text(slide, str(index), 12.41, 7.22, .32, .19, 9, '666666', center=True)
        note_shape = next(s for s in notes.findall('.//p:sp', NS) if s.find("p:nvSpPr/p:nvPr/p:ph[@type='body']", NS) is not None)
        note = spec.get('notes', '')
        if spec.get('source'):
            note += '\n\n来源：'+spec['source']
        O.text_style(note_shape, note, 14, INK, font='宋体')
        for shape in notes.findall('.//p:sp', NS):
            if shape.find("p:nvSpPr/p:nvPr/p:ph[@type='sldNum']", NS) is not None:
                for value in shape.findall('.//a:t', NS):
                    value.text = str(index)
        parts[path] = O.xml(slide); parts[O.relpath(path)] = O.xml(slide_rels)
        parts[notes_path] = O.xml(notes); parts[O.relpath(notes_path)] = O.xml(notes_rels)
        O.sub(presentation.find('p:sldIdLst', NS), 'p', 'sldId', id=255+index, **{O.tag('r', 'id'): f'rIdReport{index}'})
        O.sub(rels, 'pr', 'Relationship', Id=f'rIdReport{index}', Type=O.REL+'slide', Target='slides/'+Path(path).name)
        for name, mimetype in [(path, 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'), (notes_path, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')]:
            O.sub(types, 'ct', 'Override', PartName='/'+name, ContentType=mimetype)
        if kind == 'chart':
            O.sub(types, 'ct', 'Override', PartName=f'/ppt/charts/chart{1000+index}.xml', ContentType='application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
    if not any(c.get('Extension') == 'xlsx' for c in types):
        O.sub(types, 'ct', 'Default', Extension='xlsx', ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    parts['ppt/presentation.xml'] = O.xml(presentation)
    parts['ppt/_rels/presentation.xml.rels'] = O.xml(rels)
    parts['[Content_Types].xml'] = O.xml(types)
    parts = prune(parts)
    if 'docProps/app.xml' in parts:
        app = E.fromstring(parts['docProps/app.xml'])
        ep = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
        # The reference carries statistics and titles from an older deck.
        # Word/paragraph counts and composite title vectors are optional;
        # remove them rather than publish inherited values as current facts.
        for label in ('Words', 'Paragraphs', 'HeadingPairs', 'TitlesOfParts'):
            for child in app.findall('{'+ep+'}'+label):
                app.remove(child)
        hidden = sum(E.fromstring(parts[f'ppt/slides/reportSlide{i}.xml']).get('show') == '0'
                     for i in range(1, len(slides)+1))
        for label, count in (('Slides', len(slides)), ('Notes', len(slides)), ('HiddenSlides', hidden)):
            item = app.find('{'+ep+'}'+label)
            if item is None:
                item = E.SubElement(app, '{'+ep+'}'+label)
            item.text = str(count)
        parts['docProps/app.xml'] = O.xml(app)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, value in parts.items():
            archive.writestr(name, value)
    report = {'pptx': output.name, 'pptx_sha256': sha(output), 'slides': len(slides), 'template_sha256': sha(template), 'content_sha256': sha(content_path), 'chart_slides': [i for i, s in enumerate(slides, 1) if s['layout']=='chart'], 'rendered': False, 'visual_review': 'required after rendering'}
    output.with_suffix('.build.json').write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print(json.dumps(report, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--content', type=Path, required=True)
    parser.add_argument('--output', type=Path, default=ROOT/'outputs/my-report.pptx')
    parser.add_argument('--template', type=Path, default=ROOT/'templates/reference/approved-week9.pptx')
    args = parser.parse_args()
    build(args.content.resolve(), args.output.resolve(), args.template.resolve())


if __name__ == '__main__':
    main()
