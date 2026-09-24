"""
Reference model of HW#4C built in CadQuery, in SolidWorks coordinates:
  X = left/right (symmetric about the Right plane)
  Y = up  (Top plane = bottom face of the base plate, base top at Y=20)
  Z = toward the viewer (Front plane = BACK face of the arch wall; part front at Z=96)
"""
import math
import cadquery as cq

# ---------- key dimensions (mm) ----------
BASE_W, BASE_D, BASE_T = 222.0, 96.0, 20.0     # base plate
PAD_W, PAD_T = 120.0, 6.0                       # pad under the base
BASE_TOP = BASE_T                               # Y = 20
CY = BASE_TOP + 35.0                            # arch centre height, Y = 55
WALL_T = 14.0                                   # arch wall (Z 0..14)
BAND_T = 14.0                                   # raised horseshoe band + boss (Z 14..28)
R_OUT, R_IN = 59.0, 24.0                        # band outer / inner radius
LOBE_C, LOBE_R = 41.5, 17.5                     # band end lobes (centre x, radius)
SLOT_RO, SLOT_RI, SLOT_LR = 48.0, 35.0, 6.5     # through slot
BOSS_D, HOLE_D = 32.0, 16.0
GUS_T, GUS_X0 = 10.0, 88.0                      # gusset thickness, inner face x
GUS_TOP = CY                                    # gusset flat top at Y=55 (35 above base)
GUS_ZSLOPE_END = 96.0 - 56.0                    # slope reaches Y=55 at Z=40
EAR_T, EAR_X0 = 12.0, 111.0                     # ear thickness, inner face x
EAR_ZMIN = 96.0 - 32.0                          # ear spans Z 64..96
EAR_CZ, EAR_CY, EAR_R, EAR_HOLE_D = 80.0, BASE_TOP - 36.0, 16.0, 16.0
BIG_HOLE_D, BIG_HOLE_Z = 22.0, 96.0 - 44.0      # Ø22 at Z=52

# tangent point of the side slopes on the R59 arc (from the corner point (98, CY))
cos_t = R_OUT / 98.0
sin_t = math.sqrt(1 - cos_t**2)
TX, TY = R_OUT * cos_t, CY + R_OUT * sin_t       # (35.52, 102.11)

# ---------- 1. base plate ----------
base = cq.Workplane("XY").box(BASE_W, BASE_T, BASE_D, centered=(True, False, False))  # Y 0..20, Z 0..96

# ---------- 2. pad ----------
pad = cq.Workplane("XY").box(PAD_W, PAD_T, BASE_D, centered=(True, False, False)).translate((0, -PAD_T, 0))
part = base.union(pad)

# ---------- 3. Ø22 through hole ----------
big_hole = (cq.Workplane("XZ", origin=(0, 50, BIG_HOLE_Z)).circle(BIG_HOLE_D / 2).extrude(100, both=True))
part = part.cut(big_hole)

# ---------- 4. arch wall (Z 0..14) ----------
wall_profile = (
    cq.Workplane("XY")
    .moveTo(-98, BASE_TOP).lineTo(98, BASE_TOP).lineTo(98, CY)
    .lineTo(TX, TY)
    .threePointArc((0, CY + R_OUT), (-TX, TY))
    .lineTo(-98, CY).close()
)
wall = wall_profile.extrude(WALL_T)
part = part.union(wall)

# ---------- 5. raised horseshoe band + boss (Z 14..28) ----------
band = (
    cq.Workplane("XY", origin=(0, 0, WALL_T))
    .moveTo(-R_OUT, CY)
    .threePointArc((0, CY + R_OUT), (R_OUT, CY))
    .threePointArc((LOBE_C, CY - LOBE_R), (R_IN, CY))
    .threePointArc((0, CY + R_IN), (-R_IN, CY))
    .threePointArc((-LOBE_C, CY - LOBE_R), (-R_OUT, CY))
    .close()
    .extrude(BAND_T)
)
boss = cq.Workplane("XY", origin=(0, 0, WALL_T)).center(0, CY).circle(BOSS_D / 2).extrude(BAND_T)
part = part.union(band).union(boss)

# ---------- 6. through slot ----------
slot = (
    cq.Workplane("XY", origin=(0, 0, -10))
    .moveTo(-SLOT_RO, CY)
    .threePointArc((0, CY + SLOT_RO), (SLOT_RO, CY))
    .threePointArc((LOBE_C, CY - SLOT_LR), (SLOT_RI, CY))
    .threePointArc((0, CY + SLOT_RI), (-SLOT_RI, CY))
    .threePointArc((-LOBE_C, CY - SLOT_LR), (-SLOT_RO, CY))
    .close()
    .extrude(60)
)
part = part.cut(slot)

# ---------- 7. Ø16 through hole ----------
hole16 = cq.Workplane("XY", origin=(0, 0, -10)).center(0, CY).circle(HOLE_D / 2).extrude(60)
part = part.cut(hole16)

# ---------- 8. gussets (X 88..98, mirrored) ----------
def yz_profile_solid(pts, x0, thick):
    """closed polyline (z, y) pairs in the plane X = x0, extruded +X by thick."""
    plane = cq.Plane(origin=(x0, 0, 0), xDir=(0, 0, -1), normal=(1, 0, 0))  # sketch x = -Z, sketch y = +Y
    wp = cq.Workplane(plane).moveTo(-pts[0][0], pts[0][1])
    for p in pts[1:]:
        wp = wp.lineTo(-p[0], p[1])
    return wp.close().extrude(thick)

gus_pts = [(96, BASE_TOP), (GUS_ZSLOPE_END, GUS_TOP), (0, GUS_TOP), (0, BASE_TOP)]
gus_r = yz_profile_solid(gus_pts, GUS_X0, GUS_T)
gus_l = gus_r.mirror("YZ")
part = part.union(gus_r).union(gus_l)

# ---------- 9. ears (X 111..123, mirrored) ----------
plane = cq.Plane(origin=(EAR_X0, 0, 0), xDir=(0, 0, -1), normal=(1, 0, 0))  # sketch x = -Z
ear_r = (
    cq.Workplane(plane)
    .moveTo(-96, BASE_TOP).lineTo(-EAR_ZMIN, BASE_TOP).lineTo(-EAR_ZMIN, EAR_CY)
    .threePointArc((-EAR_CZ, EAR_CY - EAR_R), (-96, EAR_CY))
    .close()
    .extrude(EAR_T)
)
ear_l = ear_r.mirror("YZ")
part = part.union(ear_r).union(ear_l)

# ---------- 10. ear holes (Ø16 along X, through both ears) ----------
plane = cq.Plane(origin=(-150, 0, 0), xDir=(0, 0, -1), normal=(1, 0, 0))
ear_holes = cq.Workplane(plane).center(-EAR_CZ, EAR_CY).circle(EAR_HOLE_D / 2).extrude(300)
part = part.cut(ear_holes)

vol = part.val().Volume()
print(f"Volume = {vol:,.1f} mm^3   (drawing says 779,719)")
bb = part.val().BoundingBox()
print(f"BBox X[{bb.xmin:.1f},{bb.xmax:.1f}] Y[{bb.ymin:.1f},{bb.ymax:.1f}] Z[{bb.zmin:.1f},{bb.zmax:.1f}]")

import sys
out = sys.argv[1] if len(sys.argv) > 1 else "HW4C.step"
cq.exporters.export(part, out)
print("wrote", out)
