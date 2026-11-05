# 기획서 Word 만들기: python3 scripts/make-docx.py
# docs/01_프로젝트_기획서.md → docs/01_프로젝트_기획서.docx (python-docx)
# 제목(#·##·###), 표(| … |), 글머리(- ), 번호(1. ), 인용(> ), 코드 묶음(``` … ```), **굵게** · `코드` 를 옮깁니다.
import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'docs' / '01_프로젝트_기획서.md'
OUT = ROOT / 'docs' / '01_프로젝트_기획서.docx'
FONT = '맑은 고딕'


def set_font(run, size=None, bold=None, mono=False):
    name = 'D2Coding' if mono else FONT
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold


def add_inline(par, text, size=10.5, bold=False):
    # **굵게** · `코드` · [글](주소) → 글자만
    text = re.sub(r'\[([^\]]+)\]\([^)]+\)', r'\1', text)
    parts = re.split(r'(\*\*[^*]+\*\*|`[^`]+`)', text)
    for p in parts:
        if not p:
            continue
        if p.startswith('**') and p.endswith('**'):
            set_font(par.add_run(p[2:-2]), size, True)
        elif p.startswith('`') and p.endswith('`'):
            set_font(par.add_run(p[1:-1]), size, bold, mono=True)
        else:
            set_font(par.add_run(p), size, bold)


def shade(cell, hex_fill):
    tcpr = cell._element.get_or_add_tcPr()
    shd = tcpr.makeelement(qn('w:shd'), {qn('w:val'): 'clear', qn('w:color'): 'auto', qn('w:fill'): hex_fill})
    tcpr.append(shd)


def add_table(doc, rows):
    cells = [[c.strip() for c in r.strip().strip('|').split('|')] for r in rows]
    cells = [r for r in cells if not all(re.fullmatch(r':?-{2,}:?', c) for c in r)]
    width = max(len(r) for r in cells)
    t = doc.add_table(rows=len(cells), cols=width)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, r in enumerate(cells):
        for j in range(width):
            cell = t.cell(i, j)
            cell.text = ''
            add_inline(cell.paragraphs[0], r[j] if j < len(r) else '', 9.5, i == 0)
            if i == 0:
                shade(cell, 'EEF2F7')
    doc.add_paragraph()


def main():
    md = SRC.read_text(encoding='utf-8').splitlines()
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = FONT
    st.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)
    st.font.size = Pt(10.5)
    i = 0
    while i < len(md):
        line = md[i]
        if line.startswith('```'):
            block = []
            i += 1
            while i < len(md) and not md[i].startswith('```'):
                block.append(md[i])
                i += 1
            i += 1
            note = doc.add_paragraph().add_run('(흐름도 — Markdown 판에서 그림으로 보입니다)')
            set_font(note, 9.5)
            note.italic = True
            for b in block:
                p = doc.add_paragraph()
                set_font(p.add_run(b.replace('<br/>', ' ')), 8.5, mono=True)
                p.paragraph_format.space_after = Pt(0)
            doc.add_paragraph()
            continue
        if line.startswith('|'):
            rows = []
            while i < len(md) and md[i].startswith('|'):
                rows.append(md[i])
                i += 1
            add_table(doc, rows)
            continue
        m = re.match(r'^(#{1,3})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            h = doc.add_heading(level=0 if level == 1 else level - 1)
            set_font(h.add_run(m.group(2)), {1: 18, 2: 14, 3: 12}[level], True)
            for r in h.runs:
                r.font.color.rgb = RGBColor(0x0F, 0x25, 0x44)
        elif line.strip() == '---':
            pass
        elif re.match(r'^\s*- ', line):
            add_inline(doc.add_paragraph(style='List Bullet'), re.sub(r'^\s*- ', '', line))
        elif re.match(r'^\d+\. ', line):
            # 번호는 글자 그대로 둡니다(Word 번호 목록은 장을 넘어 이어져 10장이 7번부터 시작함)
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(14)
            add_inline(p, line)
        elif line.startswith('> '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(14)
            add_inline(p, line[2:], 10)
        elif line.strip():
            add_inline(doc.add_paragraph(), line)
        i += 1
    doc.save(OUT)
    print('wrote', OUT.relative_to(ROOT), '·', len(doc.paragraphs), '문단 ·', len(doc.tables), '표')


if __name__ == '__main__':
    sys.exit(main())
