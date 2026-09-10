"""Optional on-device macOS Vision OCR. No cloud calls or pip dependencies."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile


def binary():
    if sys.platform != 'darwin' or not shutil.which('swiftc'):
        raise ValueError('本地 OCR 需要 macOS 与 Swift 命令行工具；此图片保留为待提取。')
    source = Path(__file__).with_suffix('.swift')
    revision = hashlib.sha256(source.read_bytes()).hexdigest()[:12]
    root = Path.home() / '.cache' / 'dsh-doc-to-excel'
    root.mkdir(parents=True, exist_ok=True)
    target = root / ('vision-' + revision)
    if not target.exists():
        with tempfile.TemporaryDirectory(dir=root) as directory:
            temporary = Path(directory) / 'vision'
            subprocess.run(['swiftc', str(source), '-o', str(temporary)], check=True,
                           capture_output=True, timeout=90)
            temporary.replace(target)
    return target


def image_text(data):
    with tempfile.TemporaryDirectory(prefix='dsh-doc-ocr-') as directory:
        source = Path(directory) / 'image.png'
        source.write_bytes(data)
        return read(source)


def read(path, page=None):
    result = subprocess.run([str(binary()), str(path)] + ([str(page)] if page else []),
                            check=True, capture_output=True, timeout=30)
    return json.loads(result.stdout)['text'].strip()
