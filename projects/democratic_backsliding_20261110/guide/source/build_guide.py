#!/usr/bin/env python3
"""Build the presenter's guide from guide_source.md.

Outputs (same content, two formats):
  presenter_guide.md   clean GitHub Markdown, readable in the browser
  presenter_guide.pdf  A4, for reading on screen and for printing

The "What you say" blocks are not typed by hand: they are read from the
speaker notes of the delivered PPTX, so guide and deck cannot drift apart.
The build stops if a script marker or a slide headline in the source no
longer matches the deck.

Usage:
  python3 build_guide.py [output_folder]

Needs: reportlab, and the Arimo and Gelasio fonts (free stand-ins for the
deck fonts Arial and Georgia). Get the fonts with
  npm pack @expo-google-fonts/arimo @expo-google-fonts/gelasio
unpack both archives into one folder as expo-google-fonts-<name>/package/
and point GUIDE_FONT_DIR at that folder.
"""
import os
import re
import sys
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer,
                                Table, TableStyle, CondPageBreak, KeepTogether,
                                Flowable)
from reportlab.platypus.flowables import PageBreakIfNotEmpty
from reportlab.platypus.tableofcontents import TableOfContents

HERE = Path(__file__).resolve().parent
SRC = HERE / 'guide_source.md'
PPTX_NAME = 'democratic_backsliding_20261111_110847.pptx'
# In the project this file lives in <project>/guide/source/, the deck in <project>/exports/.
_IN_PROJECT = HERE.parent.parent / 'exports' / PPTX_NAME
PPTX = _IN_PROJECT if _IN_PROJECT.exists() else Path(
    '/home/user/ppt-master/projects/democratic_backsliding_20261110/exports') / PPTX_NAME
FONT_DIR = Path(os.environ.get('GUIDE_FONT_DIR', '/home/user/fonts/_pack'))
if len(sys.argv) > 1:
    OUT_DIR = Path(sys.argv[1])
else:
    OUT_DIR = HERE.parent if HERE.name == 'source' else HERE / 'out'

DECK_TITLE = 'Democratic Backsliding and Institutional Capture'
WPM_CALM, WPM_FAST = 130, 150

# ----------------------------------------------------------------------------
# 1. Read the deck (speaker notes + slide text)
# ----------------------------------------------------------------------------
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'


def para_texts(root):
    out = []
    for p in root.iter(A + 'p'):
        parts = []
        for node in p.iter():
            if node.tag == A + 't' and node.text:
                parts.append(node.text)
            elif node.tag == A + 'br':
                parts.append(' ')
        s = ''.join(parts).strip()
        if s:
            out.append(s)
    return out


def norm(s):
    return re.sub(r'\s+', ' ', s).strip()


def read_deck(path):
    z = zipfile.ZipFile(path)
    notes, shapes = {}, {}
    for n in z.namelist():
        m = re.fullmatch(r'ppt/slides/slide(\d+)\.xml', n)
        if not m:
            continue
        num = int(m.group(1))
        root = ET.fromstring(z.read(n))
        shapes[num] = [norm(' '.join(para_texts(sp))) for sp in root.iter(P + 'sp')]
        rr = ET.fromstring(z.read(f'ppt/slides/_rels/slide{num}.xml.rels'))
        for r in rr:
            if r.get('Type', '').endswith('/notesSlide'):
                target = 'ppt/' + r.get('Target').replace('../', '')
                nroot = ET.fromstring(z.read(target))
                body = []
                for sp in nroot.iter(P + 'sp'):
                    ph = sp.find('.//' + P + 'ph')
                    if ph is not None and ph.get('type') == 'body':
                        body = para_texts(sp)
                notes[num] = ' '.join(body)
    return notes, shapes


NOTES, SHAPES = read_deck(PPTX)


def split_script(num, marks):
    note = NOTES[num]
    parts, rest = [], note
    for mk in marks:
        if note.count(mk) != 1:
            raise SystemExit(f'slide {num}: marker {mk!r} occurs {note.count(mk)} times in the notes')
        idx = rest.find(mk)
        if idx <= 0:
            raise SystemExit(f'slide {num}: marker {mk!r} out of order')
        parts.append(rest[:idx].strip())
        rest = rest[idx:]
    parts.append(rest.strip())
    if ' '.join(parts) != note:
        raise SystemExit(f'slide {num}: script does not reassemble to the notes')
    return parts


def check_headline(num, text):
    if norm(text) not in SHAPES[num]:
        raise SystemExit(f'slide {num}: headline not found on the slide: {text!r}')


# ----------------------------------------------------------------------------
# 2. Parse the source
# ----------------------------------------------------------------------------
def is_special(s):
    return (s.startswith(('@', '#', '>', '|', '- ', '{{')) or re.match(r'^\d+\. ', s) is not None)


def parse_inner(rows):
    out, i = [], 0
    while i < len(rows):
        if re.match(r'^\d+\. ', rows[i]):
            items = []
            while i < len(rows) and re.match(r'^\d+\. ', rows[i]):
                items.append(re.sub(r'^\d+\. ', '', rows[i]))
                i += 1
            out.append({'t': 'ol', 'items': items})
        elif rows[i]:
            out.append({'t': 'p', 'text': rows[i]})
            i += 1
        else:
            i += 1
    return out


def parse(text):
    lines = text.split('\n')
    blocks, i, cols = [], 0, None
    while i < len(lines):
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s.startswith('@'):
            name, _, arg = s[1:].partition(' ')
            if name == 'cols':
                cols = {'w': [float(x) for x in arg.split() if '=' not in x], 'bold': [0]}
                for x in arg.split():
                    if x.startswith('b='):
                        cols['bold'] = [int(v) for v in x[2:].split(',')]
            else:
                blocks.append({'t': name, 'text': arg.strip()})
            i += 1
            continue
        m = re.match(r'^(#{1,4}) (.*)$', s)
        if m:
            blocks.append({'t': 'h%d' % len(m.group(1)), 'text': m.group(2)})
            i += 1
            continue
        m = re.match(r'^\{\{script (\d+)(.*)\}\}$', s)
        if m:
            marks = [x.strip() for x in m.group(2).split('|')[1:]]
            blocks.append({'t': 'script', 'n': int(m.group(1)), 'parts': split_script(int(m.group(1)), marks)})
            i += 1
            continue
        if s.startswith('>'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('>'):
                rows.append(lines[i].strip()[1:].strip())
                i += 1
            title = re.fullmatch(r'\*\*(.+)\*\*', rows[0]).group(1)
            blocks.append({'t': 'callout', 'title': title, 'body': parse_inner(rows[1:])})
            continue
        if s.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append([c.strip() for c in lines[i].strip().strip('|').split('|')])
                i += 1
            assert set(''.join(rows[1])) <= set('-: '), rows[1]
            width = len(rows[0])
            assert all(len(r) == width for r in rows), [r for r in rows if len(r) != width]
            blocks.append({'t': 'table', 'head': rows[0], 'rows': rows[2:], 'cols': cols})
            cols = None
            continue
        if s.startswith('- '):
            items = []
            while i < len(lines) and lines[i].strip().startswith('- '):
                items.append(lines[i].strip()[2:])
                i += 1
            blocks.append({'t': 'ul', 'items': items})
            continue
        if re.match(r'^\d+\. ', s):
            items = []
            while i < len(lines) and re.match(r'^\d+\. ', lines[i].strip()):
                items.append(re.sub(r'^\d+\. ', '', lines[i].strip()))
                i += 1
            blocks.append({'t': 'ol', 'items': items})
            continue
        para = [s]
        i += 1
        while i < len(lines) and lines[i].strip() and not is_special(lines[i].strip()):
            para.append(lines[i].strip())
            i += 1
        bold = re.fullmatch(r'\*\*([^*]+)\*\*', para[0])
        if bold and len(para) >= 2:
            blocks.append({'t': 'qa', 'q': bold.group(1), 'a': ' '.join(para[1:])})
        elif bold and len(para) == 1:
            blocks.append({'t': 'lead', 'text': bold.group(1)})
        else:
            blocks.append({'t': 'p', 'text': ' '.join(para)})
    return blocks


BLOCKS = parse(SRC.read_text(encoding='utf-8'))

# attach @stress to the script block before it, check @headline against the deck
merged = []
current_slide = None
for b in BLOCKS:
    if b['t'] == 'h3':
        m = re.match(r'^Slide (\d+) · (.*)$', b['text'])
        current_slide = int(m.group(1)) if m else None
    if b['t'] == 'stress':
        assert merged and merged[-1]['t'] == 'script', 'stress must follow a script block'
        merged[-1]['stress'] = b['text']
        continue
    if b['t'] == 'headline':
        check_headline(current_slide, b['text'])
    merged.append(b)
BLOCKS = merged


def timing_rows():
    rows, total_calm, total_words = [], 0, 0
    for n in range(1, 11):
        words = len(NOTES[n].split())
        secs = int(round(words / WPM_CALM * 60 / 5.0)) * 5
        rows.append((n, words, secs))
        total_calm += secs
        total_words += words
    return rows, total_words, total_calm


def mmss(secs):
    return '%d:%02d' % (secs // 60, secs % 60)


TIMING, TOTAL_WORDS, TOTAL_CALM = timing_rows()
TOTAL_FAST = TOTAL_WORDS / WPM_FAST * 60
SLIDE_TITLES = {}
for b in BLOCKS:
    if b['t'] == 'h3':
        m = re.match(r'^Slide (\d+) · (.*)$', b['text'])
        if m:
            SLIDE_TITLES[int(m.group(1))] = m.group(2)


SHORT_TITLES = {}
for b in BLOCKS:
    if b['t'] == 'table' and b['head'][0] == 'Slide' and b['head'][1].startswith('The message'):
        for r in b['rows']:
            num, name = r[0].split(' · ', 1)
            SHORT_TITLES[int(num)] = name
assert sorted(SHORT_TITLES) == list(range(1, 11)), SHORT_TITLES


def timing_table_rows():
    rows = []
    for left, right in zip(TIMING[:5], TIMING[5:]):
        rows.append([f'{left[0]} · {SHORT_TITLES[left[0]]}', mmss(left[2]),
                     f'{right[0]} · {SHORT_TITLES[right[0]]}', mmss(right[2])])
    return rows


TIMING_HEAD = ['Slide', 'Time', 'Slide', 'Time']


def minutes_words(secs):
    """Round to the nearest half minute and say it in plain words."""
    halves = int(round(secs / 30.0))
    whole, half = divmod(halves, 2)
    return f'{whole} and a half minutes' if half else f'{whole} minutes'


TIMING_SENTENCE = (
    f'At a calm speed, the talk takes about {minutes_words(TOTAL_CALM)}. '
    f'If you speak faster, it takes about {minutes_words(TOTAL_FAST)}. '
    f'The times below are for a calm speed ({WPM_CALM} words per minute).'
)

PARTS = [b['text'] for b in BLOCKS if b['t'] == 'h2']

# ----------------------------------------------------------------------------
# 3. Markdown output
# ----------------------------------------------------------------------------
def gh_anchor(heading):
    a = heading.lower()
    a = re.sub(r'[^\w\- ]', '', a)
    return a.replace(' ', '-')


def build_markdown():
    out = []
    title = next(b['text'] for b in BLOCKS if b['t'] == 'h1')
    for b in BLOCKS:
        t = b['t']
        if t == 'h1':
            out.append(f'# {title}: {DECK_TITLE}\n')
        elif t == 'subtitle':
            pass
        elif t == 'lede':
            out.append(b['text'] + '\n')
        elif t == 'toc':
            out.append('## Contents\n')
            for p in PARTS:
                out.append(f'- [{p}](#{gh_anchor(p)})')
            out.append('')
        elif t in ('pagebreak', 'need'):
            pass
        elif t in ('h2', 'h3', 'h4'):
            out.append('#' * int(t[1]) + ' ' + b['text'] + '\n')
        elif t == 'headline':
            out.append(f'*On the slide: “{b["text"]}”*\n')
        elif t == 'script':
            out.append('\n>\n'.join('> ' + p for p in b['parts']) + '\n')
            if b.get('stress'):
                out.append(f'**Stress this:** {b["stress"]}\n')
        elif t == 'p':
            out.append(b['text'] + '\n')
        elif t == 'lead':
            out.append(f'**{b["text"]}**\n')
        elif t == 'qa':
            out.append(f'**{b["q"]}**\\\n{b["a"]}\n')
        elif t == 'ul':
            out.extend('- ' + it for it in b['items'])
            out.append('')
        elif t == 'ol':
            out.extend(f'{k}. {it}' for k, it in enumerate(b['items'], 1))
            out.append('')
        elif t == 'callout':
            rows = [f'> **{b["title"]}**', '>']
            for inner in b['body']:
                if inner['t'] == 'ol':
                    rows.extend(f'> {k}. {it}' for k, it in enumerate(inner['items'], 1))
                else:
                    rows.append('> ' + inner['text'])
            out.append('\n'.join(rows) + '\n')
        elif t == 'table':
            out.append('| ' + ' | '.join(b['head']) + ' |')
            out.append('|' + '|'.join(['---'] * len(b['head'])) + '|')
            out.extend('| ' + ' | '.join(r) + ' |' for r in b['rows'])
            out.append('')
        elif t == 'timing':
            out.append(TIMING_SENTENCE + '\n')
            out.append('| ' + ' | '.join(TIMING_HEAD) + ' |')
            out.append('|---|---|---|---|')
            out.extend('| ' + ' | '.join(r) + ' |' for r in timing_table_rows())
            out.append('')
        else:
            raise SystemExit(f'markdown: unknown block {t}')
    text = '\n'.join(out)
    text = re.sub(r'\n{3,}', '\n\n', text).strip() + '\n'
    return text


# ----------------------------------------------------------------------------
# 4. PDF output
# ----------------------------------------------------------------------------
NAVY, RED, SLATE = HexColor('#16233B'), HexColor('#B3261E'), HexColor('#4F6D8F')
TEXT, MUTED, LINE = HexColor('#1F2933'), HexColor('#5A6270'), HexColor('#D5CEBF')
CREAM, CREAM2 = HexColor('#FAF8F3'), HexColor('#F4F0E7')

FONTS = {
    'Arimo': 'expo-google-fonts-arimo-0.4.3/package/400Regular/Arimo_400Regular.ttf',
    'Arimo-Bold': 'expo-google-fonts-arimo-0.4.3/package/700Bold/Arimo_700Bold.ttf',
    'Arimo-Italic': 'expo-google-fonts-arimo-0.4.3/package/400Regular_Italic/Arimo_400Regular_Italic.ttf',
    'Arimo-BoldItalic': 'expo-google-fonts-arimo-0.4.3/package/700Bold_Italic/Arimo_700Bold_Italic.ttf',
    'Gelasio': 'expo-google-fonts-gelasio-0.4.1/package/400Regular/Gelasio_400Regular.ttf',
    'Gelasio-Bold': 'expo-google-fonts-gelasio-0.4.1/package/700Bold/Gelasio_700Bold.ttf',
    'Gelasio-Italic': 'expo-google-fonts-gelasio-0.4.1/package/400Regular_Italic/Gelasio_400Regular_Italic.ttf',
    'Gelasio-BoldItalic': 'expo-google-fonts-gelasio-0.4.1/package/700Bold_Italic/Gelasio_700Bold_Italic.ttf',
}
for name, rel in FONTS.items():
    pdfmetrics.registerFont(TTFont(name, str(FONT_DIR / rel)))
for fam in ('Arimo', 'Gelasio'):
    pdfmetrics.registerFontFamily(fam, normal=fam, bold=fam + '-Bold', italic=fam + '-Italic',
                                  boldItalic=fam + '-BoldItalic')

W, H = A4
LM = RM = 25 * mm
TM, BM = 20 * mm, 23 * mm
FW = W - LM - RM
BAND_H = 68 * mm
NB = '\u00a0'


def nb(t):
    """Keep short number groups on one line: 'slide 5', '41 percent', '12 of 16' (not years)."""
    t = re.sub(r'\b([Ss]lides?|[Pp]arts?|[Pp]ages?|[Qq]uestion|[Ss]tep) (\d)', r'\1' + NB + r'\2', t)
    t = re.sub(r'(\d) (percent|billion|million)\b', r'\1' + NB + r'\2', t)
    t = re.sub(r'\b(\d{1,3}) (of|to|and) (\d{1,3})\b', r'\1' + NB + r'\2' + NB + r'\3', t)
    return t


def esc(t):
    return nb(t).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def fmt(t):
    t = esc(t)
    t = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', t)
    t = re.sub(r'\*(.+?)\*', r'<i>\1</i>', t)
    return t


def inline(text, bold_color=None):
    out, pos = [], 0
    for m in re.finditer(r'\[([^\]]+)\]\(([^)\s]+)\)', text):
        out.append(fmt(text[pos:m.start()]))
        href = m.group(2).replace('&', '&amp;').replace('"', '%22')
        out.append(f'<a href="{href}" color="#B3261E"><u>{fmt(m.group(1))}</u></a>')
        pos = m.end()
    out.append(fmt(text[pos:]))
    s = ''.join(out)
    if bold_color:
        s = s.replace('<b>', f'<b><font color="{bold_color}">').replace('</b>', '</font></b>')
    return s


def PS(name, **kw):
    base = dict(fontName='Arimo', fontSize=10.6, leading=15.8, textColor=TEXT,
                allowWidows=0, allowOrphans=0, splitLongWords=0, bulletFontName='Arimo')
    base.update(kw)
    return ParagraphStyle(name, **base)


ST = {
    'body': PS('body', spaceAfter=6.5),
    'lede': PS('lede', fontSize=12.6, leading=18.5, textColor=SLATE, spaceAfter=14),
    'h2': PS('h2', fontName='Gelasio-Bold', fontSize=19, leading=23, textColor=NAVY),
    'h3': PS('h3', fontName='Gelasio-Bold', fontSize=13.4, leading=17.5, textColor=NAVY,
             spaceBefore=13, spaceAfter=6, keepWithNext=1),
    'headline': PS('headline', fontName='Gelasio-Italic', fontSize=11.4, leading=16, textColor=SLATE,
                   spaceAfter=2, keepWithNext=1),
    'script': PS('script', fontName='Gelasio', fontSize=11.2, leading=16.8),
    'stress': PS('stress', fontSize=10.4, leading=14.5, textColor=NAVY),
    'lead': PS('lead', fontName='Arimo-Bold', textColor=NAVY, spaceBefore=5, spaceAfter=4, keepWithNext=1),
    'q': PS('q', fontName='Arimo-Bold', textColor=NAVY, spaceAfter=2, keepWithNext=1),
    'a': PS('a', spaceAfter=8),
    'ul': PS('ul', leftIndent=15, bulletIndent=3, spaceAfter=4.2),
    'ol': PS('ol', leftIndent=19, bulletIndent=1, spaceAfter=4.2),
    'co_title': PS('co_title', fontName='Arimo-Bold', fontSize=10.8, textColor=NAVY, spaceAfter=3),
    'co_body': PS('co_body', spaceAfter=2.5),
    'co_ol': PS('co_ol', leftIndent=17, bulletIndent=1, spaceAfter=2.5),
    'th': PS('th', fontName='Arimo-Bold', fontSize=9.3, leading=12.5, textColor=white),
    'td': PS('td', fontSize=9.9, leading=14.2),
    'td1': PS('td1', fontName='Arimo-Bold', fontSize=9.9, leading=14.2, textColor=NAVY),
    'toc0': PS('toc0', fontName='Arimo-Bold', fontSize=10.4, leading=14.5, textColor=NAVY, spaceBefore=5.5),
    'toc1': PS('toc1', fontSize=9.5, leading=14.3, textColor=TEXT, leftIndent=17),
}


class Label(Flowable):
    """Small tracked capitals with a hairline to the right margin."""

    def __init__(self, text, color=RED, size=8.3, track=1.15, rule=True, before=11, after=5):
        super().__init__()
        self.text, self.color, self.size, self.track, self.rule = text.upper(), color, size, track, rule
        self.spaceBefore, self.spaceAfter, self.keepWithNext = before, after, 1

    def wrap(self, aw, ah):
        self.aw = aw
        return aw, self.size + 2

    def draw(self):
        c = self.canv
        t = c.beginText(0, 1.6)
        t.setFont('Arimo-Bold', self.size)
        t.setCharSpace(self.track)
        t.setFillColor(self.color)
        t.textOut(self.text)
        t.setCharSpace(0)
        c.drawText(t)
        if self.rule:
            w = pdfmetrics.stringWidth(self.text, 'Arimo-Bold', self.size) + self.track * len(self.text)
            c.setStrokeColor(LINE)
            c.setLineWidth(0.6)
            y = 1.6 + self.size * 0.34
            c.line(w + 7, y, self.aw, y)


class PartHeading(Flowable):
    def __init__(self, num, title, key):
        super().__init__()
        self.para = Paragraph(f'<font color="#B3261E">{num}</font>&nbsp;&nbsp;&nbsp;{esc(title)}', ST['h2'])
        self.spaceBefore, self.spaceAfter, self.keepWithNext = 20, 9, 1
        self.toc_level, self.bm_key = 0, key
        self.outline_title = self.toc_text = f'{num}. {title}'

    def wrap(self, aw, ah):
        _, h = self.para.wrap(aw, ah)
        self.h = h + 10
        return aw, self.h

    def draw(self):
        c = self.canv
        c.setFillColor(RED)
        c.rect(0, self.h - 2.8, 30, 2.8, stroke=0, fill=1)
        self.para.drawOn(c, 0, 0)


class SlideBar(Flowable):
    HT, TAB = 30, 66

    def __init__(self, num, title):
        super().__init__()
        self.num, self.title = num, title
        self.spaceBefore, self.spaceAfter, self.keepWithNext = 6, 8, 1
        self.toc_level, self.bm_key = 1, f'slide-{num}'
        self.outline_title = self.toc_text = f'Slide {num} · {title}'

    def wrap(self, aw, ah):
        self.aw = aw
        return aw, self.HT

    def draw(self):
        c = self.canv
        c.setFillColor(NAVY)
        c.roundRect(0, 0, self.aw, self.HT, 3.2, stroke=0, fill=1)
        c.setFillColor(RED)
        c.roundRect(0, 0, self.TAB, self.HT, 3.2, stroke=0, fill=1)
        c.rect(self.TAB - 6, 0, 6, self.HT, stroke=0, fill=1)
        label = f'SLIDE {self.num}'
        track = 1.3
        lw = pdfmetrics.stringWidth(label, 'Arimo-Bold', 9) + track * (len(label) - 1)
        t = c.beginText((self.TAB - lw) / 2.0, 10.6)
        t.setFont('Arimo-Bold', 9)
        t.setCharSpace(track)
        t.setFillColor(white)
        t.textOut(label)
        t.setCharSpace(0)
        c.drawText(t)
        c.setFillColor(white)
        c.setFont('Gelasio-Bold', 14.2)
        c.drawString(self.TAB + 12, 9.8, self.title)


def bullet(items, style, numbered=False):
    out = []
    for k, it in enumerate(items, 1):
        mark = (f'<font name="Arimo-Bold" color="#16233B">{k}.</font>' if numbered
                else '<font color="#B3261E">•</font>')
        out.append(Paragraph(inline(it, '#16233B'), style, bulletText=None))
        out[-1] = Paragraph(f'<bullet>{mark}</bullet>' + inline(it, '#16233B'), style)
    return out


def data_table(head, rows, cols):
    if not cols:
        cols = {'w': [100.0 / len(head)] * len(head), 'bold': [0]}
    widths = [FW * c / sum(cols['w']) for c in cols['w']]
    data = [[Paragraph(esc(h), ST['th']) for h in head]]
    for r in rows:
        data.append([Paragraph(inline(c), ST['td1'] if k in cols['bold'] else ST['td'])
                     for k, c in enumerate(r)])
    t = Table(data, colWidths=widths, repeatRows=1)
    style = [
        ('FONTNAME', (0, 0), (-1, -1), 'Arimo'),
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 7), ('RIGHTPADDING', (0, 0), (-1, -1), 7),
        ('TOPPADDING', (0, 0), (-1, -1), 4.6), ('BOTTOMPADDING', (0, 0), (-1, -1), 5.4),
        ('LINEBELOW', (0, 1), (-1, -1), 0.5, LINE),
    ]
    for k in range(1, len(data)):
        if k % 2 == 0:
            style.append(('BACKGROUND', (0, k), (-1, k), CREAM))
    t.setStyle(TableStyle(style))
    t.spaceBefore, t.spaceAfter = 4, 10
    return t


def script_box(parts, stress):
    rows = [[Paragraph(esc(p), ST['script'])] for p in parts]
    n = len(rows)
    style = [
        ('FONTNAME', (0, 0), (-1, -1), 'Arimo'),
        ('BACKGROUND', (0, 0), (-1, -1), CREAM2),
        ('LINEBEFORE', (0, 0), (0, -1), 2.4, NAVY),
        ('LEFTPADDING', (0, 0), (-1, -1), 14), ('RIGHTPADDING', (0, 0), (-1, -1), 13),
        ('TOPPADDING', (0, 0), (-1, -1), 3.6), ('BOTTOMPADDING', (0, 0), (-1, -1), 3.6),
        ('TOPPADDING', (0, 0), (-1, 0), 10), ('BOTTOMPADDING', (0, n - 1), (-1, n - 1), 10),
    ]
    if stress:
        rows.append([Paragraph(
            '<font name="Arimo-Bold" size="8.3" color="#B3261E">STRESS THIS</font>'
            f'&nbsp;&nbsp;&nbsp;&nbsp;<b>{esc(stress)}</b>', ST['stress'])])
        style += [('LINEABOVE', (0, n), (-1, n), 0.6, LINE),
                  ('TOPPADDING', (0, n), (-1, n), 7), ('BOTTOMPADDING', (0, n), (-1, n), 8)]
    t = Table(rows, colWidths=[FW])
    t.setStyle(TableStyle(style))
    t.spaceBefore, t.spaceAfter = 0, 4
    return t


def callout(title, body):
    cell = [Paragraph(esc(title), ST['co_title'])]
    for inner in body:
        if inner['t'] == 'ol':
            cell += bullet(inner['items'], ST['co_ol'], numbered=True)
        else:
            cell.append(Paragraph(inline(inner['text'], '#16233B'), ST['co_body']))
    t = Table([[cell]], colWidths=[FW])
    t.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Arimo'),
        ('BACKGROUND', (0, 0), (-1, -1), CREAM),
        ('BOX', (0, 0), (-1, -1), 0.6, LINE),
        ('LINEBEFORE', (0, 0), (0, -1), 2.6, RED),
        ('LEFTPADDING', (0, 0), (-1, -1), 14), ('RIGHTPADDING', (0, 0), (-1, -1), 13),
        ('TOPPADDING', (0, 0), (-1, -1), 10), ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    t.spaceBefore, t.spaceAfter = 8, 12
    return t


def draw_page(c, doc):
    c.saveState()
    if doc.page == 1:
        c.setFillColor(NAVY)
        c.rect(0, H - BAND_H, W, BAND_H, stroke=0, fill=1)
        c.setFillColor(RED)
        c.rect(LM, H - 22 * mm, 30, 2.8, stroke=0, fill=1)
        t = c.beginText(LM, H - 30 * mm)
        t.setFont('Arimo-Bold', 9.6)
        t.setCharSpace(2.0)
        t.setFillColor(HexColor('#D5CEBF'))
        t.textOut('PRESENTER’S GUIDE')
        t.setCharSpace(0)
        c.drawText(t)
        c.setFillColor(white)
        c.setFont('Gelasio-Bold', 28)
        c.drawString(LM, H - 43.5 * mm, 'Democratic Backsliding')
        c.drawString(LM, H - 55.5 * mm, 'and Institutional Capture')
    c.setStrokeColor(LINE)
    c.setLineWidth(0.6)
    c.line(LM, 16 * mm, W - RM, 16 * mm)
    c.setFont('Arimo', 8)
    c.setFillColor(MUTED)
    c.drawString(LM, 11.2 * mm, f'Presenter’s guide · {DECK_TITLE}')
    c.drawRightString(W - RM, 11.2 * mm, str(doc.page))
    c.restoreState()


class GuideDoc(BaseDocTemplate):
    def __init__(self, filename):
        super().__init__(filename, pagesize=A4, leftMargin=LM, rightMargin=RM, topMargin=TM,
                         bottomMargin=BM, title=f'Presenter’s guide: {DECK_TITLE}',
                         subject='Study and speaking guide for the seminar presentation',
                         author='', creator='', lang='en-GB', initialFontName='Arimo')
        frame = Frame(LM, BM, FW, H - TM - BM, leftPadding=0, rightPadding=0, topPadding=0,
                      bottomPadding=0, id='main')
        self.addPageTemplates([PageTemplate(id='page', frames=[frame], onPage=draw_page)])

    def afterFlowable(self, fl):
        level = getattr(fl, 'toc_level', None)
        if level is None:
            return
        self.canv.bookmarkPage(fl.bm_key)
        self.canv.addOutlineEntry(fl.outline_title, fl.bm_key, level=level, closed=True)
        if getattr(fl, 'toc_text', None):
            self.notify('TOCEntry', (level, fl.toc_text, self.page, fl.bm_key))


GROUPED_LABELS = ('Good to know', 'If someone asks', 'Help for the discussion',
                  'If you have too little time')
STOP = ('h1', 'h2', 'h3', 'h4', 'pagebreak', 'need')


def block_flowables(b):
    """Flowables for one simple content block."""
    t = b['t']
    if t == 'p':
        return [Paragraph(inline(b['text'], '#16233B'), ST['body'])]
    if t == 'lead':
        return [Paragraph(esc(b['text']), ST['lead'])]
    if t == 'qa':
        # flat on purpose: a KeepTogether inside a KeepTogether reports an endless height
        return [Paragraph(esc(b['q']), ST['q']), Paragraph(inline(b['a'], '#16233B'), ST['a'])]
    if t == 'ul':
        return bullet(b['items'], ST['ul']) + [Spacer(1, 3)]
    if t == 'ol':
        return bullet(b['items'], ST['ol'], numbered=True) + [Spacer(1, 3)]
    if t == 'callout':
        return [callout(b['title'], b['body'])]
    if t == 'table':
        return [data_table(b['head'], b['rows'], b['cols'])]
    if t == 'script':
        return [script_box(b['parts'], b.get('stress'))]
    if t == 'headline':
        return [Paragraph(
            '<font name="Arimo-Bold" size="7.8" color="#5A6270">ON THE SLIDE</font>&nbsp;&nbsp;&nbsp;'
            f'“{esc(b["text"])}”', ST['headline'])]
    if t == 'timing':
        return [Paragraph(esc(TIMING_SENTENCE), ST['body']),
                data_table(TIMING_HEAD, timing_table_rows(), {'w': [36, 14, 36, 14], 'bold': [0, 2]})]
    raise SystemExit(f'pdf: unknown block {t}')


def build_story():
    story = []
    read_no = 0
    k = 0
    while k < len(BLOCKS):
        b = BLOCKS[k]
        t = b['t']
        if t in ('h1', 'subtitle'):
            pass
        elif t == 'lede':
            story.append(Spacer(1, BAND_H - TM + 8 * mm))
            story.append(Paragraph(esc(b['text']), ST['lede']))
        elif t == 'toc':
            story.append(Label('Contents', before=6, after=1))
            toc = TableOfContents()
            toc.levelStyles = [ST['toc0'], ST['toc1']]
            toc.dotsMinLevel = 0
            toc.tableStyle = TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Arimo'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ('VALIGN', (0, 0), (-1, -1), 'TOP')])
            story.append(toc)
        elif t == 'pagebreak':
            story.append(PageBreakIfNotEmpty())
        elif t == 'need':
            story.append(CondPageBreak(float(b['text']) * mm))
        elif t == 'h2':
            num, title = re.match(r'^(\d+)\. (.*)$', b['text']).groups()
            story.append(CondPageBreak(110 * mm))
            story.append(PartHeading(num, title, f'part-{num}'))
        elif t == 'h3':
            m = re.match(r'^Slide (\d+) · (.*)$', b['text'])
            if m:
                n = int(m.group(1))
                if n == 11:
                    story.append(CondPageBreak(45 * mm))
                elif n > 1:
                    story.append(PageBreakIfNotEmpty())
                story.append(SlideBar(n, m.group(2)))
            else:
                # a reading: keep the heading and its list on one page
                read_no += 1
                head = Paragraph(inline(b['text']), ST['h3'])
                head.toc_level, head.bm_key = 1, f'reading-{read_no}'
                head.outline_title = re.sub(r'\*', '', b['text'])
                group = [head]
                while k + 1 < len(BLOCKS) and BLOCKS[k + 1]['t'] not in STOP + ('callout',):
                    k += 1
                    group += block_flowables(BLOCKS[k])
                story.append(KeepTogether(group))
        elif t == 'h4':
            label = Label(b['text'])
            if b['text'] in GROUPED_LABELS:
                group = [label]
                while k + 1 < len(BLOCKS) and BLOCKS[k + 1]['t'] not in STOP:
                    k += 1
                    group += block_flowables(BLOCKS[k])
                story.append(KeepTogether(group))
            else:
                story.append(label)
        elif t == 'lead' and k + 1 < len(BLOCKS) and BLOCKS[k + 1]['t'] in ('ul', 'ol'):
            group = block_flowables(b) + block_flowables(BLOCKS[k + 1])
            k += 1
            story.append(KeepTogether(group))
        else:
            story.extend(block_flowables(b))
        k += 1
    return story


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    md = build_markdown()
    (OUT_DIR / 'presenter_guide.md').write_text(md, encoding='utf-8')
    pdf_path = OUT_DIR / 'presenter_guide.pdf'
    doc = GuideDoc(str(pdf_path))
    doc.multiBuild(build_story())
    words = len(re.findall(r"[A-Za-zÀ-ÿ’'\-]+", md))
    print(f'markdown: {len(md):,} chars, about {words:,} words')
    print(f'pdf: {pdf_path} ({pdf_path.stat().st_size:,} bytes), pages={doc.page}')
    print(f'script words (slides 1-10): {TOTAL_WORDS}; calm {mmss(TOTAL_CALM)}; fast {mmss(int(TOTAL_FAST))}')
    print('timing sentence:', TIMING_SENTENCE)


if __name__ == '__main__':
    main()
