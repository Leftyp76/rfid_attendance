# HW#7A drawing macro: plan

Goal: recreate the IME 1140 In-Lab Assignment #7A drawing from the supplied part,
using the template `Hayden_p_ansi.drwdot`. Save it as `HW#7APerhach.SLDDRW` and
`HW#7APerhach.PDF`.

## What the files contain

**Part** (`IME 1140 ... In-Lab Assignment Part #1.SLDPRT`, inch units, 2 decimals)

| Feature | Sketch dimensions | Meaning |
|---|---|---|
| Boss-Extrude1 (2.00 deep) | Sketch4: .25, .25, 1.50, 3.00 | L-shaped profile, 2.00 wide |
| Cut-Extrude1 | Sketch5: Ø1.00, 1.50, 1.00 | Big hole in the upright |
| Cut-Extrude2 | Sketch6: 2× Ø.50, 1.00, .50, .50, .50 | Two holes in the flange |
| Fillet2 | R1.00 | Rounded top of the upright |
| Fillet3 | R.125 (stored as R0.13 at 2 decimals) | Outer corners |
| Fillet4 | R.25 | Inside bend |
| Chamfer1 | .025 × 45° (stored as 0.03) | Flange edges |

Material is not set in the part. The drawing only needs it as title block text.

**Template.** Sheet "Hayden's Drawing Template", A size, 1:1. The title block has
fixed text: UNLESS OTHERWISE SPECIFIED, tolerances, PER ASME Y14.5 2018, MATERIAL:,
DATE:, TITLE:, DRAWN BY: HAYDEN PERHACH, SCALE: 1:1, SHEET 1 OF 1, REV B, SIZE A.
The MATERIAL, DATE and TITLE values are blank and are not linked to properties, so
the macro has to add them as notes.

## Sheet layout (11 × 8.5 in, scale 1:1, third-angle projection)

Positions are measured from the reference image and given as view centres from the
sheet's lower-left corner.

| View | Orientation | Centre (in) | Display |
|---|---|---|---|
| Front | Face of the upright, flange at the bottom | about 2.5, 3.2 | Hidden lines visible |
| Top | Projected up from Front | about 2.5, 6.8 | Hidden lines visible |
| Right | Projected right from Front, flange pointing left | about 6.7, 3.2 | Hidden lines visible |
| Isometric | Standard isometric | about 9.5, 6.5 | Hidden lines removed, tangent edges shown |

**Orientation check.** I can't see how the part is oriented in its own coordinates.
Instead of assuming, the macro finds the Ø1.00 hole and the two Ø.50 holes in the part
geometry. Their axes show which way is "front" and which way is "up". It then builds a
custom view orientation, so the front view always shows the upright face with the
flange at the bottom.

## Annotations to recreate

| View | Item | Placement |
|---|---|---|
| Top | 2 X Ø.50 | Diameter dimension with leader to the left hole, "2 X" prefix |
| Top | 1.00 | Between hole centres, below the view |
| Top | .50 | Right hole centre to right edge, below the view |
| Top | 2.00 | Overall width, below the 1.00 |
| Top | .50 | Hole centres to the front edge, right side |
| Top | 1.50 | Overall depth, right side, outside the .50 |
| Front | R1.00 | Radius leader, upper right |
| Front | Ø1.00 | Diameter leader, upper right |
| Front | 1.50 | Hole centre to bottom edge, right side |
| Front | 1.00 | Hole centre to right edge, below the view |
| Right | .25 | Upright thickness, above the view |
| Right | 3.00 | Overall height, right side |
| Right | .25 | Flange thickness, left side, arrows outside |
| Right | R.25 | Inside bend, leader from the upper left |
| Right | 2 X R.125 | Outer bend corner, leader to the lower right, 3 decimals |
| Right | .025 X 45° TYP | Chamfer at the flange tip, 3 decimals |
| All | Centre marks and centrelines | Centre marks on holes, centrelines through hidden holes |

**How dimensions get made.** The macro does not import the model dimensions, because
SolidWorks would place them by its own rules (for example the 1.50 from Sketch4 would
land in the right view). Instead it reads each view's visible edges, picks the right
edge by geometry (for example "circle of radius .25 in the top view" or "rightmost
vertical line"), and places a reference dimension at the position from the table.

**Checks.** After each dimension the macro reads its value back and compares it with the
expected number (2.00, 1.50, R.125 and so on). The final message lists every dimension
as OK or wrong, like the volume checks in the HW#4C macro.

## Title block notes to add

| Field | Value |
|---|---|
| MATERIAL | 316 STAINLESS STEEL |
| TITLE | IME 1140 ENGINEERING DRAWINGS I / IN-LAB ASSIGNMENT #7A |
| DATE | See question 2 |
| DRAWN BY | Already in the template as HAYDEN PERHACH |

## Saving

1. You open the part in SolidWorks first. The macro uses the open part.
2. The macro looks for the template at
   `C:\Users\hperhach\OneDrive - Cal Poly\Documents\Hayden_p_ansi.DRWDOT`.
   If it isn't there, it asks you to pick it.
3. It saves `HW#7APerhach.SLDDRW` and `HW#7APerhach.PDF` in the same folder as the part.
4. The part file is not changed.

## Build order (each step is checked)

1. Find the part, the template and the part's orientation.
2. New drawing from the template, sheet scale 1:1.
3. Front view, then Top and Right projected from it, then the Isometric view.
4. Display modes, centre marks, centrelines.
5. Dimensions for the Top, Front and Right views, with value checks.
6. Title block notes.
7. Save the drawing and the PDF, then show the check report.

## Open questions

1. **Sheet number and ASME year.** The reference says SHEET 1 OF 4 and ASME Y14.5 2009.
   Your template says SHEET 1 OF 1 and 2018. Default: keep your template.
2. **Date.** The reference shows 07-18-2026. Default: today's date.
3. **Instructor note.** The line "NOTES: USE THE IN-LAB ASSIGNMENT PART#1 ON CANVAS..."
   at the top looks like the assignment's own instruction. Default: leave it off.
4. **Logo.** I can't tell whether your template has the Cal Poly logo. If it doesn't,
   the macro won't add one.
