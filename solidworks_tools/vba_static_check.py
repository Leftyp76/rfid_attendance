"""
vba_static_check.py - catch the VBA mistakes that bit us before a macro is ever run in SolidWorks.

Usage:  python3 vba_static_check.py Macro.bas

Checks:
  - duplicate Sub/Function names
  - capitalised names that are called but never defined/declared (typos, removed helpers)
  - variables that shadow VBA built-ins (e.g. 'Dim sgn As Double' breaks Sgn())
  - integer division written as '\\\\' (escaping accident) instead of '\\'
  - a double-backslash string literal (path separator escaping accident)
  - an element of a Variant array (e.g. faces(i)) passed to a ByRef typed parameter
    (e.g. 'fc As Object' without ByVal): VBA refuses it with "ByRef argument type mismatch"
"""
import re, sys

KW = set('''If Then Else ElseIf End For Next To Step Each In Do Loop While Wend Exit Sub Function Dim As Set Select Case Is Not
And Or True False Nothing Empty Long Double String Boolean Integer Variant Object Optional ByVal ByRef On Error GoTo Resume
MsgBox InputBox IIf Left Mid Len Trim UCase InStr InStrRev Replace Format Date Now Abs Sqr Array IsArray IsEmpty CLng CDbl
CStr CInt Environ Dir ChrW Chr FreeFile Open Output Print Close Application CreateObject LBound UBound ReDim Err Const
Option Explicit Raise Clear Atn Sgn Mod Fail Preserve Call
LOF EOF Input FileLen Line Cos Sin Tan Split Join Round Space Hex TypeName DoEvents'''.split())
BUILTINS = ['Sgn', 'Len', 'Left', 'Right', 'Mid', 'Format', 'Date', 'Dir', 'Abs', 'Sqr', 'Atn', 'Val', 'Str', 'Int', 'Fix', 'Log', 'Exp']

def main():
    t = open(sys.argv[1], encoding='utf-8', errors='replace').read()
    ok = True
    defined = re.findall(r'^(?:Sub|Function) (\w+)\(', t, flags=re.M)
    d = sorted(n for n in set(defined) if defined.count(n) > 1)
    if d: ok = False; print('DUPLICATE procedures:', d)
    code = '\n'.join('' if l.strip().startswith("'") else re.sub(r'"[^"]*"', '""', l).split("'")[0] for l in t.split('\n'))
    called = set(re.findall(r'(?:^|[\s:(=,])([A-Z][A-Za-z0-9]+)(?=[\s(]|$)', code, flags=re.M))
    consts = set(re.findall(r'^Const (\w+)', t, flags=re.M))
    declared = set(m.group(1) for m in re.finditer(r'\b(\w+)\s*(?:\([^)]*\))?\s+As\b', code))
    unk = sorted(c for c in called if c not in KW and c not in defined and c not in consts and c not in declared and not c.startswith('vb'))
    if unk: ok = False; print('UNRESOLVED names (typo or missing helper?):', unk)
    for b in BUILTINS:
        if re.search(r'\b' + b + r'\s+As\b', code, flags=re.I) and not re.search(r'Function ' + b + r'\(', code):
            ok = False; print('SHADOWS built-in:', b)
    for i, l in enumerate(t.split('\n'), 1):
        c = re.sub(r'"[^"]*"', '""', l).split("'")[0]
        if '\\\\' in c: ok = False; print(f'line {i}: double backslash in code (integer division should be a single \\):', l.strip()[:90])
        if re.search(r'"[^"]*\\\\[^"]*"', l) and not l.strip().startswith("'"):
            ok = False; print(f'line {i}: double backslash inside a string (path separator should be a single \\):', l.strip()[:90])
    if not byref_check(t, code, defined): ok = False
    # a module-level variable with the same name as a procedure (VBA ignores case): "Ambiguous name detected"
    first_proc = re.search(r'^(?:Sub|Function) ', t, flags=re.M)
    head = code[:first_proc.start()] if first_proc else code
    modvars = set()
    for m in re.finditer(r'^(?:Dim|Private|Public)\s+(.*)$', head, flags=re.M):
        for part in m.group(1).split(','):
            mv = re.match(r'\s*(\w+)', part)
            if mv: modvars.add(mv.group(1).lower())
    clash = sorted(n for n in set(defined) if n.lower() in modvars)
    if clash: ok = False; print('AMBIGUOUS name (module variable and procedure share it, case-insensitive):', clash)
    print(f'{len(t.splitlines())} lines, {len(defined)} procedures -> ' + ('OK' if ok else 'FIX THE ITEMS ABOVE'))

def split_args(s):
    out, depth, cur = [], 0, ''
    for ch in s:
        if ch == '(': depth += 1
        elif ch == ')': depth -= 1
        if ch == ',' and depth == 0:
            out.append(cur.strip()); cur = ''
        else:
            cur += ch
    if cur.strip(): out.append(cur.strip())
    return out

def byref_check(t, code, defined):
    """Variant-array element -> ByRef typed parameter = compile error 'ByRef argument type mismatch'."""
    ok = True
    joined = re.sub(r' _\n\s*', ' ', t)
    sigs = {}
    for m in re.finditer(r'^(?:Sub|Function) (\w+)\((.*?)\)(?: As \w+)?\s*$', joined, flags=re.M):
        params = []
        for prm in split_args(m.group(2)):
            prm = re.sub(r'^Optional\s+', '', prm)
            byval = prm.startswith('ByVal ')
            prm = re.sub(r'^(ByVal|ByRef)\s+', '', prm)
            mt = re.match(r'(\w+)(\(\))?\s+As\s+(\w+)', prm)
            params.append(None if (byval or not mt or mt.group(2) or mt.group(3) == 'Variant') else mt.group(3))
        sigs[m.group(1)] = params
    arrtype = {m.group(1): m.group(2) for m in re.finditer(r'\b(\w+)\s*\([^)]*\)\s+As\s+(\w+)', code)}
    jcode = re.sub(r' _\n\s*', ' ', code)
    for i, line in enumerate(jcode.split('\n'), 1):
        for name, params in sigs.items():
            if not any(params): continue
            for m in re.finditer(r'(?:^|[^\w.])' + name + r'(\(|\s+(?=[^=\s]))', line):
                if m.group(1) == '(':
                    depth, j = 1, m.end()
                    while j < len(line) and depth:
                        depth += {'(': 1, ')': -1}.get(line[j], 0); j += 1
                    args = split_args(line[m.end():j - 1])
                else:
                    if not re.match(r'^\s*(?:\w+:\s*)?' + name + r'\s', line.split(':')[-1] if False else line.strip() + ' ') and not re.search(r':\s*' + name + r'\s', line):
                        continue
                    args = split_args(line[m.end():].split(':')[0])
                for k, a in enumerate(args):
                    if k < len(params) and params[k]:
                        ma = re.match(r'^(\w+)\([^()]*\)$', a.strip())
                        if ma and ma.group(1) not in defined and arrtype.get(ma.group(1)) != params[k]:
                            ok = False
                            print(f'line ~{i}: {a.strip()} (a Variant array element?) passed to ByRef "As {params[k]}" parameter of {name} - make that parameter ByVal')
    return ok

if __name__ == '__main__':
    main()
