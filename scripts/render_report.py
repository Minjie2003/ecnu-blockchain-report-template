"""Export a saved PPTX through the installed Windows presentation application."""
from pathlib import Path
from datetime import datetime, timezone
import argparse
import hashlib
import json
import sys


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', type=Path, required=True)
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--pdf', action='store_true')
    parser.add_argument('--pdf-path', type=Path, help='Default: beside the input PPTX')
    args = parser.parse_args()
    if sys.platform != 'win32':
        parser.error('Windows and an installed PowerPoint-compatible COM provider are required. Export manually in your presentation application on other systems.')
    import win32com.client
    source = args.input.resolve()
    out = args.output_dir.resolve()
    if not source.is_file():
        parser.error(f'Input does not exist: {source}')
    if args.pdf_path and not args.pdf:
        parser.error('--pdf-path requires --pdf')
    out.mkdir(parents=True, exist_ok=True)
    result = {'checked_utc': datetime.now(timezone.utc).isoformat(),
              'pptx': source.name, 'pptx_sha256': digest(source),
              'errors': [], 'text_bound_warnings': [],
              'visual_review': 'not_performed_by_this_script'}
    app = pres = None
    initial_count = None
    try:
        app = win32com.client.DispatchEx('PowerPoint.Application')
        initial_count = app.Presentations.Count
        result['com_progid'] = 'PowerPoint.Application'
        for attr in ('Name', 'Version', 'Path'):
            try:
                result['application_' + attr.lower()] = str(getattr(app, attr))
            except Exception:
                pass
        pres = app.Presentations.Open(str(source), True, False, False)
        result['slide_count'] = pres.Slides.Count
        result['rendered_images'] = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides.Item(i)
            name = f'slide_{i:02d}.png'
            slide.Export(str(out/name), 'PNG', 1600, 900)
            result['rendered_images'].append(name)
            for shape in slide.Shapes:
                try:
                    if shape.HasTextFrame and shape.TextFrame.HasText:
                        frame = shape.TextFrame
                        height = float(shape.Height) - float(frame.MarginTop) - float(frame.MarginBottom)
                        bound = float(frame.TextRange.BoundHeight)
                        if bound > height + 2:
                            result['text_bound_warnings'].append({'slide': i, 'shape_id': int(shape.Id),
                                'bound_height_points': round(bound, 2), 'available_height_points': round(height, 2)})
                except Exception:
                    pass
        if args.pdf:
            pdf = args.pdf_path.resolve() if args.pdf_path else source.with_suffix('.pdf')
            pdf.parent.mkdir(parents=True, exist_ok=True)
            pres.SaveAs(str(pdf), 32)
            result.update(pdf=pdf.name, pdf_sha256=digest(pdf))
    except Exception as exc:
        result['errors'].append({'type': type(exc).__name__, 'message': str(exc)})
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None and initial_count == 0:
            try:
                if app.Presentations.Count == 0:
                    app.Quit()
            except Exception:
                pass
    result['source_unchanged'] = digest(source) == result['pptx_sha256']
    images = result.get('rendered_images', [])
    result['render_complete'] = bool(images) and len(images) == result.get('slide_count') and all((out/n).is_file() for n in images)
    result['passed'] = not result['errors'] and result['source_unchanged'] and result['render_complete']
    (out/'render_validation.json').write_text(json.dumps(result, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print(json.dumps(result, ensure_ascii=False))
    raise SystemExit(0 if result['passed'] else 1)


if __name__ == '__main__':
    main()
