"""mondai/ と answers/ の Markdown から第一問（種類・品目の判定）を取り出し、JSON にする。

使い方: python3 tools/extract_dai1.py <出力JSON>
"""
import json
import re
import sys

YEARS = [
    ('R1', 'mondai/R1.md', 'answers/R1_official.md', 'official', '令和元年度'),
    ('R3', 'mondai/R3.md', 'answers/R3_official.md', 'official', '令和3年度'),
    ('R4', 'mondai/R4.md', 'answers/R4_official.md', 'official', '令和4年度'),
    ('R5', 'mondai/R5.md', 'answers/R5.md', 'own', '令和5年度'),
    ('R6', 'mondai/R6.md', 'answers/R6.md', 'own', '令和6年度'),
    ('R7', 'mondai/R7.md', 'answers/R7.md', 'own', '令和7年度'),
]


def section(text, start, end):
    s = text.index(start)
    e = text.find(end, s + len(start))
    return text[s:e if e != -1 else len(text)]


def clean(s):
    return re.sub(r'<br>', '／', s).strip()


out = []
for key, qpath, apath, src, label in YEARS:
    q = section(open(qpath, encoding='utf-8').read(), '## 第一問', '\n---')
    lead = next(l for l in q.splitlines() if l.startswith('次のイ'))
    rows = {}
    for line in q.splitlines():
        m = re.match(r'^\| (イ|ロ|ハ|ニ|ホ) \| (.+) \|$', line)
        if m:
            rows[m.group(1)] = clean(m.group(2))

    a = section(open(apath, encoding='utf-8').read(), '## 第一問', '\n## 第二問')
    items = []
    for block in re.split(r'\n### ', a)[1:]:
        lines = block.splitlines()
        m = re.match(r'^(イ|ロ|ハ|ニ|ホ)\s+(.+?)／(.+?)(\s|$)', lines[0])
        kana, kind, item = m.group(1), m.group(2), m.group(3)
        reasons, notes = [], []
        for l in lines[1:]:
            l = l.strip()
            if not l or l.startswith('---'):
                continue
            if l.startswith('>'):
                body = l.lstrip('> ').strip()
                if body.startswith(('注', 'ポイント')):
                    notes.append(re.sub(r'^(注|ポイント)：', '', body))
                continue
            reasons.append(l)
        items.append({'kana': kana, 'kind': kind, 'item': item, 'problem': rows[kana],
                      'reasons': reasons, 'notes': notes})
    out.append({'key': key, 'label': label, 'src': src, 'lead': lead, 'items': items})

json.dump(out, open(sys.argv[1], 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print(sum(len(y['items']) for y in out), 'items')
