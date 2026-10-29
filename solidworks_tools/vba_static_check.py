"""
vba_static_check.py - catch the VBA mistakes that bit us before a macro is ever run in SolidWorks.

Usage:  python3 vba_static_check.py Macro.bas

Checks:
  - duplicate Sub/Function names
  - capitalised names that are called but never defined/declared (typos, removed helpers)
  - variables that shadow VBA built-ins (e.g. 'Dim sgn As Double' breaks Sgn())
  - integer division written as '\\\\' (escaping accident) instead of '\\'
  - a double-backslash string literal (path separator escaping accident)
  - typed arrays passed to Variant ByRef helpers is not detectable here: see the playbook rule
"""
import re, sys

KW = set('''If Then Else ElseIf End For Next To Step Each In Do Loop While Wend Exit Sub Function Dim As Set Select Case Is Not
And Or True False Nothing Empty Long Double String Boolean Integer Variant Object Optional ByVal ByRef On Error GoTo Resume
MsgBox InputBox IIf Left Mid Len Trim UCase InStr InStrRev Replace Format Date Now Abs Sqr Array IsArray IsEmpty CLng CDbl
CStr CInt Environ Dir ChrW Chr FreeFile Open Output Print Close Application CreateObject LBound UBound ReDim Err Const
Option Explicit Raise Clear Atn Sgn Mod Fail Preserve Call'''.split())
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
    print(f'{len(t.splitlines())} lines, {len(defined)} procedures -> ' + ('OK' if ok else 'FIX THE ITEMS ABOVE'))

if __name__ == '__main__':
    main()
