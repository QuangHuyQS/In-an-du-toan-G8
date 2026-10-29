# Sinh ma nguon VBA (ANSI, CRLF) tu src/ -> gen/
# Chuoi "..." co ky tu tieng Viet -> U("...\uXXXX...") vi VBE chi doc ANSI.
import os, re, sys

SRC = os.path.dirname(os.path.abspath(__file__))
GEN = os.path.join(os.path.dirname(SRC), 'gen')
os.makedirs(GEN, exist_ok=True)

CLS_HDR = ('VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  \'True\nEND\n'
           'Attribute VB_Name = "{n}"\nAttribute VB_GlobalNameSpace = False\n'
           'Attribute VB_Creatable = False\nAttribute VB_PredeclaredId = False\n'
           'Attribute VB_Exposed = False\nOption Explicit\n')

def esc(s):
    return ''.join(c if ord(c) < 128 else '\\u%04X' % ord(c) for c in s)

def convert_line(line):
    out, i, n = [], 0, len(line)
    while i < n:
        c = line[i]
        if c == "'":                     # chu thich: phan con lai
            rest = line[i:]
            if any(ord(ch) > 127 for ch in rest):
                raise ValueError('non-ASCII in comment: ' + line)
            out.append(rest); break
        if c == '"':
            j = i + 1
            while True:
                k = line.index('"', j)
                if k + 1 < n and line[k + 1] == '"':
                    j = k + 2; continue
                break
            lit = line[i:k + 1]
            if any(ord(ch) > 127 for ch in lit):
                out.append('U(' + esc(lit) + ')')
            else:
                out.append(lit)
            i = k + 1; continue
        out.append(c); i += 1
    return ''.join(out)

def gen(name, outname=None, cls=None):
    txt = open(os.path.join(SRC, name), encoding='utf-8').read()
    lines = [convert_line(l) for l in txt.replace('\r\n', '\n').split('\n')]
    body = '\n'.join(lines)
    if cls:
        body = CLS_HDR.format(n=cls) + body
    body.encode('ascii')
    with open(os.path.join(GEN, outname or name), 'w', encoding='ascii', newline='\r\n') as f:
        f.write(body)
    print('gen', outname or name)

gen('inanDTG8_V6.bas')
gen('modInAnMain.bas')
gen('BuildInAnForm.bas')
gen('frmInDuToan_code.txt')
gen('clsInAnBtn.cls', cls='clsInAnBtn')
