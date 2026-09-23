"""Build allowlisted source and self-contained Windows Python distributions.

Run with the project's venv Python. No mailbox or runtime settings are collected.
"""
import hashlib
import json
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from outlook_archiver import __version__


def copy_tree(source, target):
    shutil.copytree(source, target, ignore=shutil.ignore_patterns('__pycache__', '*.pyc', 'site-packages'))


def zip_tree(folder):
    output = folder.parent / (folder.name + '.zip')
    with zipfile.ZipFile(output, 'x', zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(folder.rglob('*')):
            if path.is_file():
                archive.write(path, path.relative_to(folder.parent))
    return output


def main():
    from datetime import datetime
    release = ROOT / 'releases' / f'{__version__}_{datetime.now():%Y%m%d_%H%M%S}'
    release.mkdir(parents=True, exist_ok=False)
    source = release / f'OutlookMarkdown-{__version__}-Source'
    portable = release / f'OutlookMarkdown-{__version__}-Windows-x64'
    source.mkdir(); portable.mkdir()
    for target in (source, portable):
        copy_tree(ROOT / 'outlook_archiver', target / 'outlook_archiver')
        for name in ('requirements.txt', 'requirements-lock.txt', '一页中文使用说明.md', 'README.md', 'LICENSE', 'CHANGELOG.md', 'TESTING.md'):
            shutil.copy2(ROOT / name, target / name)
        shutil.copy2(ROOT / 'packaging' / '00_开始这里.txt', target / '00_开始这里.txt')
    for name in ('tests', 'packaging'):
        copy_tree(ROOT / name, source / name)
    (source / '源码恢复说明.txt').write_text(
        '使用 Windows 上的 Python 3.13 x64：\n'
        'python -m venv .venv\n.venv\\Scripts\\python -m pip install -r requirements-lock.txt\n'
        '.venv\\Scripts\\python -m outlook_archiver gui\n'
        '重新构建便携包：.venv\\Scripts\\python packaging\\build_release.py\n', encoding='utf-8-sig')
    runtime = portable / 'runtime'; runtime.mkdir()
    base = Path(sys.base_prefix)
    for name in ('DLLs', 'Lib', 'tcl'):
        copy_tree(base / name, runtime / name)
    for name in ('python.exe', 'pythonw.exe', 'python3.dll', 'python313.dll', 'vcruntime140.dll', 'vcruntime140_1.dll', 'LICENSE.txt'):
        shutil.copy2(base / name, runtime / name)
    packages = Path(sys.prefix) / 'Lib' / 'site-packages'
    # Only dependency distributions listed in the lock file; never copy global/user packages.
    import importlib.metadata
    names = [line.split('==')[0] for line in (ROOT / 'requirements-lock.txt').read_text().splitlines() if '==' in line]
    for name in names:
        distribution = importlib.metadata.distribution(name)
        for relative in distribution.files or []:
            candidate = Path(distribution.locate_file(relative)).resolve()
            try:
                rel = candidate.relative_to(packages.resolve())
            except ValueError:
                continue
            if candidate.is_file() and '__pycache__' not in rel.parts and candidate.suffix != '.pyc':
                dest = runtime / 'Lib' / 'site-packages' / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(candidate, dest)
    for name in ('Start.bat', 'portable_launch.pyw'):
        shutil.copy2(ROOT / 'packaging' / name, portable / name)
    # Windows cmd is sensitive to batch line endings.
    batch = portable / 'Start.bat'
    batch.write_bytes(batch.read_text().replace('\r\n', '\n').replace('\n', '\r\n').encode('ascii'))
    info = {'version': __version__, 'python': sys.version, 'platform': sys.platform, 'dependencies': names,
            'contents': 'Allowlisted application source and Python runtime; no mail data or personal configuration.'}
    (portable / 'BUILD-INFO.json').write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding='utf-8')
    finalize(release, source, portable)
    print(release, flush=True)


def finalize(release, source, portable):
    archives = [zip_tree(source), zip_tree(portable)]
    sums = []
    for archive in archives:
        with archive.open('rb') as stream:
            digest = hashlib.file_digest(stream, 'sha256').hexdigest()
        sums.append(f'{digest}  {archive.name}')
    (release / 'SHA256SUMS.txt').write_text('\n'.join(sums) + '\n', encoding='ascii')


if __name__ == '__main__':
    main()
