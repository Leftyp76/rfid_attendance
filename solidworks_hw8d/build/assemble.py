# Builds HW8D_Drawing_Macro.bas: the part-specific pieces in this folder (p1..p4) plus the proven
# helpers copied word for word from the HW#8C macro (with the few listed text changes).
# Run from anywhere:  python3 solidworks_hw8d/build/assemble.py
import re, os
B = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(B))
def procs(path):
    t = open(path, encoding='utf-8').read().replace('\r\n', '\n')
    lines = t.split('\n'); out = {}; i = 0
    while i < len(lines):
        m = re.match(r'^(Sub|Function) (\w+)\(', lines[i])
        if m:
            j = i
            while j > 0 and lines[j-1].startswith("'") and not lines[j-1].startswith("' ====="):
                j -= 1
            k = i
            while not lines[k].startswith('End ' + m.group(1)): k += 1
            out[m.group(2)] = '\n'.join(lines[j:k+1]); i = k + 1
        else:
            i += 1
    return out
c8 = procs(os.path.join(ROOT, 'solidworks_hw8c', 'HW8C_Drawing_Macro.bas'))
def take(name, subs=()):
    body = c8[name]
    for a, b, *cnt in subs:
        n = body.count(a)
        assert n == (cnt[0] if cnt else 1), (name, a, n)
        body = body.replace(a, b)
    return body + '\n\n'
H = lambda t: "' =====================================================================================\n'  " + t + "\n' =====================================================================================\n"
out = open(f'{B}/p1_head.bas').read() + '\n'
out += H('Run control, part, template  (from HW#8C)')
for n in ['NoDimDialog', 'SpeedMode', 'ToDo', 'Info']: out += take(n)
out += take('Stage', [('"   (modelled threads make SolidWorks slow - please wait)"', '"   (please wait)"'),
                      ("' show the step in SolidWorks' status bar (the modelled threads make some steps slow)", "' show the step in SolidWorks' status bar")])
out += take('FindPart', [('Part 3', 'Part 4', 2)])
for n in ['PartPt', 'SheetPt', 'SX', 'SY', 'FindTemplate', 'FileExists']: out += take(n)
out += take('NewDrawing', [('"Sheet scale 1:2"', '"Sheet scale 1:1"')])
out += open(f'{B}/p2_views.bas').read() + '\n'
out += H('View helpers  (from HW#8C)')
for n in ['EnsureSheetScale', 'SetViewScale', 'ViewSc', 'ScaleTxt', 'StdIsoName', 'DeleteSeg',
          'ViewNames', 'NewViewSince', 'PlaceView', 'OrientedView', 'ModelViewAt', 'ViewMatches', 'FrameOK', 'StdFrontName',
          'MakeNamedView', 'ActivateDrawing', 'DeleteView', 'SetMode']:
    out += take(n)
out += open(f'{B}/p3_dims.bas').read() + '\n'
out += H('Picking, dimension, GD&T and note helpers  (from HW#8C)')
for n in ['DimP2', 'DimT', 'ReportW', 'DatumOnEdge', 'SkLineP', 'TryDim', 'TryDimSel', 'SelPt', 'EntDist', 'SelEdge', 'KeepOneIn',
          'PickOnly', 'NoteAt', 'PlaceNote', 'DatumLabel', 'GtolReadBack', 'SymInLibrary', 'ClearCommand', 'ToSk', 'TestH',
          'DeleteDim', 'TotalDims', 'SetTol', 'TolReads', 'TolText']:
    out += take(n)
out += open(f'{B}/p4_notes.bas').read() + '\n'
out += H('Edge table  (from HW#8C)')
for n in ['CollectEdges', 'VisibleEdges', 'Extend']: out += take(n)
out += take('AddEdge', [("    ' (For this part every arc's full circle stays inside the outline, so centre +- radius is exact.)\n", "")])
out += H('Title block, finish, save  (from HW#8C)')
out += take('TitleBlock', [('SHEET 3 OF 4', 'SHEET 4 OF 4', 5),
    ('If InStr(t, "1:4") = 0 Then nt.SetText "SCALE: 1:4"', 'If InStr(t, "1:1") = 0 Then nt.SetText "SCALE: 1:1"'),
    ('If InStr(UCase(nt.GetText), "1:4") = 0 Then nt.PropertyLinkedText = "SCALE: 1:4"', 'If InStr(UCase(nt.GetText), "1:1") = 0 Then nt.PropertyLinkedText = "SCALE: 1:1"'),
    ('scaleDone = (InStr(UCase(nt.GetText), "1:4") > 0)', 'scaleDone = (InStr(UCase(nt.GetText), "1:1") > 0)'),
    ('"CAST IRON"', '"6061-T6 ALUMINUM"'), ('"Material: CAST IRON"', '"Material: 6061-T6 ALUMINUM"'),
    ('Format(Date, "mm-dd-yy")', 'Format(Date, "mm-dd-yyyy")', 2),
    ('IN-LAB ASSIGNMENT #8C', 'IN-LAB ASSIGNMENT #8D'),
    ('Check scaleDone, "SCALE: 1:4 (as on the reference)"', 'Check scaleDone, "SCALE: 1:1"'),
    ('THREE PLACE " & ChrW(177) & ".008 (note 2, as on the reference)"', 'THREE PLACE " & ChrW(177) & ".005 (note 2, as on the reference)"')])
out += take('FixTol', [('"THREE PLACE DECIMAL")), ".008")', '"THREE PLACE DECIMAL")), ".005")'),
    ('"THREE PLACE DECIMAL  <MOD-PM> .008"', '"THREE PLACE DECIMAL  <MOD-PM> .005"'),
    ('threeOK = (InStr(LineOf(u, "THREE PLACE DECIMAL"), ".008") > 0)', 'threeOK = (InStr(LineOf(u, "THREE PLACE DECIMAL"), ".005") > 0)'),
    ("+-.01 / +-.008 (the reference)", "+-.01 / +-.005 (the reference)")])
for n in ['NewLine', 'LineOf', 'AddSheetNote']: out += take(n)
out += take('FinishSheet', [('Array(viewMain, viewRight, viewAA, viewDet, viewIso)', 'Array(viewMain, viewBB, viewAA, viewIso)'),
    ('Check (n >= 5 Or (n >= 4 And viewDet Is Nothing)), n & " views on the sheet (expected 5)"', 'Check (n >= 4), n & " views on the sheet (expected 4)"')])
for n in ['SaveAll', 'WriteReport', 'TrySave', 'TrySavePdf', 'SaveFolder', 'Check']: out += take(n)
out += H('Geometry helpers  (from HW#8C)')
for n in ['XfPt', 'XfVec', 'ManualXf', 'ViewDir', 'Arr3', 'Unit', 'Dot', 'Cross', 'Scl', 'Add3', 'Sub3', 'Near', 'VecTxt']:
    out += take(n)
out = out.rstrip('\n') + '\n'
open(os.path.join(ROOT, 'solidworks_hw8d', 'HW8D_Drawing_Macro.bas'), 'w', encoding='utf-8', newline='\r\n').write(out)
print(len(out.split('\n')), 'lines')
