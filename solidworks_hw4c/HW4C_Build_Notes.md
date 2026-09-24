# HW#4C – Arch bracket – build notes

Everything below is what the macro `HW4C_SolidWorks_Macro.bas` does, in order.
Feature names in the SolidWorks tree carry the same `StepNN` prefix, so you can
drag the rollback bar to any step and see exactly what was added.

Model volume: **779,719 mm³** (drawing says 779,719 mm³ – matches).

## Orientation / reference planes

| SolidWorks plane | What it is on the part |
|---|---|
| Front Plane | Back face of the arch wall. The part grows toward the viewer (+Z). |
| Top Plane | Bottom face of the 20 mm base plate. The 6 mm pad hangs below it. |
| Right Plane | Plane of symmetry (X = 0). Gussets and ears are mirrored about it. |
| Origin | Centre of the back edge of the base plate's bottom face. |

With this setup the **Isometric** view (View Orientation > Isometric) shows the
arch face and the two ears toward the viewer, exactly like the figure.

Key heights measured from the base top face (Y = 20):
arch centre = 35 (Y = 55), gusset shoulders = 35, arch crown = 35 + 59 = 94
(Y = 114). Total height 114 + 6 = 120, which is the (120) reference dimension.

## Steps

| Step | Feature (name in tree) | Sketch plane | Geometry |
|---|---|---|---|
| 01 | Base plate 222x96x20 (boss) | Top Plane | Rectangle 222 (X) × 96 (Z), extruded up 20. |
| 02 | Pad 120x96x6 (boss) | Bottom face of base | Rectangle 120 × 96, extruded down 6 (base is 26 thick in the middle). |
| 03 | Hole Ø22 THRU (cut) | Top face of base | Ø22 at X = 0, 44 mm from the front edge. Through all. |
| 04 | Arch wall 14 thick (boss) | Front Plane | 196 wide, vertical sides 35 above the base, then straight lines from (±98, 35) tangent to the R59 crown arc. Tangent points are at (±35.52, 82.11) above the base. Extruded 14 toward the viewer. |
| 05 | Horseshoe band R59–R24 + boss Ø32 (boss) | Front face of wall | Outer arc R59, inner arc R24, rounded ends R17.5 centred at (±41.5, 35). Separate Ø32 circle at the arch centre. Both extruded 14 (total wall depth 28, the "28" in the side view). |
| 06 | Slot R48–R35 THRU (cut) | Front face of band | Curved slot: R48 outside, R35 inside, R6.5 rounded ends centred at (±41.5, 35). Through all, both directions (goes through band and wall). |
| 07 | Hole Ø16 THRU (cut) | Front face of boss | Ø16 at the arch centre, through all. |
| 08 | Gusset plane | Right Plane offset 93 | Mid-plane of the 10 mm gusset (faces at 88 and 98 = the 176 / 196 dimensions). |
| 09 | Gusset 10 thick (boss, mid-plane) | Gusset plane | Side profile: from the base front edge up a slope to 35 high at 56 mm back (the "40" from the back), flat top to the back face, down the back. Extruded mid-plane 10. |
| 10 | Mirror gusset | Right Plane | Mirror of step 09. |
| 11 | Ear plane | Right Plane offset 117 | Mid-plane of the 12 mm ear (faces at 111 and 123 = the 222 / 246 dimensions). |
| 12 | Ear 12 thick R16 (boss, mid-plane) | Ear plane | 32 deep from the front edge, flush with the base top, hangs 36 below the base top to the R16 centre, full R16 round at the bottom. Extruded mid-plane 12. |
| 13 | Mirror ear | Right Plane | Mirror of step 12. |
| 14 | 2× hole Ø16 THRU (cut) | Ear plane | Ø16 at the R16 centre, through all both directions – one cut makes both ear holes. |

## Where each drawing dimension ends up

| Drawing | Model |
|---|---|
| 246 / 222 | Ear outer faces / base plate width |
| 196 / 176 | Gusset outer / inner faces (and wall width) |
| 120 (front view) | Pad width |
| 96, 32, 44, 68, 82 (top view) | Base depth; ear depth; Ø22 hole position; band front face; wall front face (all from the front edge) |
| 40, 28 (side view) | Gusset slope top / wall + band total depth, from the back face |
| 35, 20, 26, 36 | Shoulder height; base thickness; base + pad; ear hole below base top |
| R59, R48, R35 | Wall crown & band outer; slot outer; slot inner |
| Ø32, Ø16 THRU | Boss and its hole |
| Ø22 THRU | Base hole |
| 2× Ø16, 2× R16 | Ear holes and ear rounds |
| (120) | Overall height, reference |

## Notes

* Sketches are drawn to exact coordinates by the macro; they are not dimension-
  driven (they will show as under-defined / blue). Add dimensions if your grader
  requires fully-defined sketches.
* The macro pops a message with the computed volume so you can compare with the
  drawing before submitting.
* `HW4C_reference.step` is the same part as a STEP file (no feature tree) – useful
  as a visual cross-check via File > Open in SolidWorks, but submit the .SLDPRT
  built by the macro.
