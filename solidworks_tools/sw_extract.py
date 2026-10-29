"""
sw_extract.py - read what a SolidWorks file (.SLDPRT / .SLDDRW / .DRWDOT) says about itself.

SolidWorks files are compressed; this finds every zlib/deflate block and prints:
  1. the feature tree with each sketch's / feature's dimensions (the most useful part)
  2. units, material, title-block notes (templates)
  3. every readable string (with --all)

Usage:  python3 -I sw_extract.py "Part.SLDPRT"          (add --all for every string)
It can take a few minutes on a large template (5 MB).
"""
import sys, zlib, re

def streams(data, limit=4000):
    out = []; i = 0
    while i < len(data) - 2 and len(out) < limit:
        hit = False
        for wbits in (15, -15):
            try:
                d = zlib.decompressobj(wbits)
                r = d.decompress(data[i:i + 400000])
                if len(r) > 64:
                    used = len(data[i:i + 400000]) - len(d.unused_data)
                    out.append(r); i += max(used, 1); hit = True
                    break
            except Exception:
                pass
        if not hit:
            i += 1
    return out

def texts(blocks):
    t = set()
    for r in blocks:
        for m in re.findall(rb'(?:[\x20-\x7e]\x00){4,}', r):
            t.add(m.decode('utf-16le'))
        for m in re.findall(rb'[\x20-\x7e]{6,}', r):
            t.add(m.decode())
    return t

def main():
    path = sys.argv[1]
    joined = "\n".join(sorted(texts(streams(open(path, 'rb').read()))))
    print("=== FEATURES AND DIMENSIONS ===")
    skip = ('Annotations|Notes|Markups|Comments|Favorites|History|Sensors|Directional|Ambient|Lights and Cameras|'
            'Design Binder|Solid Bodies|Surface Bodies|Equations|Selection Sets|Exploded Views|SOLIDWORKS Materials|Origin')
    feats = set(re.findall(r'<(?:Extrusion|Sketch|Feature|HoleWizard|Revolution)[^>]*>(?:<Dimension Name="[^"]*">[^<]*</Dimension>)*', joined))
    for f in sorted(feats):
        if not re.search(r'Type="(' + skip + ')"', f):
            print(f.replace('&lt;', '<').replace('&gt;', '>'))
    print("\n=== UNITS / MATERIAL ===")
    for k in ['SW_UnitsLinear"', 'SW_UnitsLinearDecimalPlaces', 'Material=']:
        for m in re.findall(re.escape(k) + r'.{0,60}', joined)[:2]:
            print(m)
    print("\n=== NOTES (templates: title-block text) ===")
    for n in sorted(set(re.findall(r'<Note>([^<]*)</Note>', joined))):
        print(" ", n)
    if '--all' in sys.argv:
        print("\n=== ALL STRINGS ===")
        print(joined)

if __name__ == '__main__':
    main()
