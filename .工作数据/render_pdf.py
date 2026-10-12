# -*- coding: utf-8 -*-
"""把 docx 批量转 PDF（经 LibreOffice），并可选逐页渲染 PNG 用于目视验收。

用法：
  python3 render_pdf.py <docx目录> [--png <PNG输出目录>] [--profile <LibreOffice用户目录>]
只用系统 /usr/bin/soffice 与 /usr/bin/pdftoppm（捆绑版有 GLIBC 不兼容问题）。
"""
import os, subprocess, sys, tempfile
from pathlib import Path

SOFFICE = '/usr/bin/soffice'
PDFTOPPM = '/usr/bin/pdftoppm'


def to_pdf(docx: Path, outdir: Path, profile: str):
    env = dict(os.environ)
    env['HOME'] = profile
    env['PATH'] = '/usr/bin:/bin:' + env.get('PATH', '')
    cmd = [SOFFICE, '--headless', '--norestore', '--convert-to', 'pdf',
           '--outdir', str(outdir), str(docx)]
    r = subprocess.run(cmd, env=env, capture_output=True, text=True, timeout=600)
    pdf = outdir / (docx.stem + '.pdf')
    if not pdf.exists():
        raise RuntimeError(f'转换失败 {docx.name}: {r.stdout[-500:]} {r.stderr[-500:]}')
    return pdf


def render_png(pdf: Path, outdir: Path, dpi=110):
    outdir.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([PDFTOPPM, '-r', str(dpi), '-png', str(pdf), str(outdir / 'page')],
                       capture_output=True, text=True, timeout=600)
    if r.returncode != 0:
        raise RuntimeError(f'渲染失败 {pdf.name}: {r.stderr[-300:]}')
    return sorted(outdir.glob('page-*.png'))


if __name__ == '__main__':
    src = Path(sys.argv[1]).resolve()
    png_dir = None
    profile = tempfile.mkdtemp(prefix='lo-profile-')
    if '--png' in sys.argv:
        png_dir = Path(sys.argv[sys.argv.index('--png') + 1]).resolve()
    if '--profile' in sys.argv:
        profile = sys.argv[sys.argv.index('--profile') + 1]
    files = sorted(src.glob('*.docx'))
    for f in files:
        pdf = to_pdf(f, src, profile)
        n = ''
        if png_dir:
            pages = render_png(pdf, png_dir / f.stem)
            n = f'  PNG {len(pages)} 页'
        print(f'{pdf.name}  {pdf.stat().st_size // 1024} KB{n}', flush=True)
    print(f'完成 {len(files)} 个文件；LibreOffice profile: {profile}')
