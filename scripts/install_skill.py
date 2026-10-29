"""Install the report skill with its standalone scripts and template inputs."""
from pathlib import Path
import argparse
import os
import shutil
import json


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--skills-dir', type=Path,
                        default=Path(os.environ.get('CODEX_HOME', str(Path.home()/'.codex')))/'skills')
    parser.add_argument('--name', default='ecnu-blockchain-report')
    args = parser.parse_args()
    if not args.name or any(c not in 'abcdefghijklmnopqrstuvwxyz0123456789-' for c in args.name):
        parser.error('--name must contain only lowercase letters, digits and hyphens')
    root = Path(__file__).resolve().parents[1]
    skill = root/'skills/ecnu-blockchain-report'
    if not skill.is_dir():
        parser.error('Run this installer from a complete clone of the template repository')
    target = args.skills_dir.expanduser().resolve()/args.name
    if target.exists():
        parser.error(f'{target} already exists. Preserve it or choose a new --name; no files were replaced.')
    inputs = [root/'templates/reference/approved-week9.pptx', root/'templates/ecnu-report-starter.pptx', root/'examples/starter.json']
    if not all(p.is_file() for p in inputs):
        parser.error('The required template inputs are missing')
    shutil.copytree(skill, target)
    for folder in ('scripts', 'templates'):
        shutil.copytree(root/folder, target/folder, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    (target/'examples').mkdir()
    shutil.copy2(root/'examples/starter.json', target/'examples/starter.json')
    shutil.copytree(root/'examples/hotrevoke-week9', target/'examples/hotrevoke-week9')
    readme=(root/'README.md').read_text(encoding='utf8').replace('skills/ecnu-blockchain-report/', '')
    (target/'README.md').write_text(readme,encoding='utf8')
    for filename in ('requirements.txt', 'LICENSE', 'NOTICE.md'):
        shutil.copy2(root/filename, target/filename)
    shutil.copytree(root/'docs', target/'docs')
    if args.name != 'ecnu-blockchain-report':
        for path in (target/'SKILL.md', target/'agents/openai.yaml', target/'README.md'):
            path.write_text(path.read_text(encoding='utf8').replace('ecnu-blockchain-report', args.name), encoding='utf8')
    print(json.dumps({'installed': str(target), 'skill': args.name,
                      'invoke': '$'+args.name, 'assets_included': True}, ensure_ascii=False))


if __name__ == '__main__':
    main()
