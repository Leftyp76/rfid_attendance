# SolidWorks Part & Drawing Macro Playbook

How to build a SolidWorks part or engineering drawing from a reference sheet with a VBA macro, so it comes out right on the first run as often as possible. This is written for the AI assistant that writes the macro. Everything here was learned building IME 1140 HW#4C (a part) and HW#7A–7D (four drawings).

Repo: `Leftyp76/rfid_attendance`, branch `claude/nice-hawking-yolktd`

| Folder | What is in it | Use it as the base for |
|---|---|---|
| `solidworks_hw4c/` | Part macro (arch bracket), fully defined sketches, volume checks | Any **part** |
| `solidworks_hw7a/` | Drawing: front, top, right, isometric, 1:1 | Simple orthographic drawings |
| `solidworks_hw7b/` | Drawing with **section view** B-B | Section views |
| `solidworks_hw7c/` | Drawing with **auxiliary view** D-D at 1:2 | Auxiliary views, other scales |
| `solidworks_hw7d/` | Drawing with **detail view** E at 2:1, turned part | Detail views, revolved parts |
| `solidworks_hw8a/` | Drawing with **SECTION A-A**, hole callouts, toleranced dims (±, limit, bilateral), datum and parallelism frame, 1:3 shaded iso; plus a part macro that changes an engraving depth | Section views, tolerances, GD&T, hole callouts |
| `solidworks_hw8b/` | Turned part with **DETAIL A 3:1**, left end view, NOTES block, circularity / runout frames, datum on a dimension; hand steps listed as TO DO | Turned parts, detail views, notes |
| `solidworks_hw8c/` | Elbow casting: **VIEW A-A auxiliary** view square to an angled flange, BASIC (boxed) dims, position / angularity frames, datums on a face and on a dimension, DETAIL B 1:1, iso 1:4 | Auxiliary views, GD&T, basic dimensions |
| `solidworks_tools/` | `sw_extract.py` reads a SolidWorks file, `vba_static_check.py` checks a macro | Every job |

**Start every new drawing macro from the newest working drawing macro** (`HW8A_Drawing_Macro.bas` v4; `HW7D` for detail views). Read section 4.8 first: it lists the mistakes HW#8A v1–v4 made, and the new macro must not repeat them. Keep its proven helpers unchanged and replace only the part-specific procedures.

---

## 1. Workflow (do it in this order)

1. **Collect the inputs.** These are the reference image, the `.SLDPRT` for a drawing, the drawing template `.DRWDOT`, the file-name rule (e.g. `HW#7E` + last name) and the user's last name (Perhach).
2. **Read the part file.** Run `python3 -I solidworks_tools/sw_extract.py "Part.SLDPRT"`. It prints every feature with its sketch dimensions, e.g. `Sketch1: D1 4.75, D2 1.00 ...`, plus the units and material. Rounded values like `0.13` are often `.125`, so check them against the reference.
3. **Read the reference image.** Crop and enlarge each view with Pillow, at 2× to 5×. List every view, every dimension and its text (prefix, suffix, decimals, parentheses), and every note.
4. **Calibrate pixels to inches.** The template's blue border is 10.5 × 8.0 in inside an 11 × 8.5 sheet, so px/in = border width in px ÷ 10.5. Check this against a known part length. Measure every annotation position as an **offset from a part feature** in sheet inches, never as an absolute sheet position.
5. **Write the plan, then wait for answers.** The user asks for a plan before building. Ask only what the reference doesn't settle: sheet "x OF y", the date (today's or the reference's), whether to include the instructor's NOTES line, and anything ambiguous.
6. **Build the macro** from the base macro, following sections 3 and 4.
7. **Run `python3 solidworks_tools/vba_static_check.py Macro.bas`** before sending it. It must print OK.
8. **Send the `.bas` file** and tell the user exactly which file to use. They paste it into Tools > Macro > New and run **`main`**, never a helper.
9. **Iterate on the report.** Every run ends with a check summary listing OK, WRONG or MISSING items, and writes `HWxx_checks.txt`. The user sends a screenshot of the summary plus the whole sheet. Fix the root cause shown, never guess at several things at once.
10. **Commit and push** each version with a version note in the macro header (`v2: ...`).

**Standing user answers so far:** today's date, written in the same format as the reference (e.g. `10-22-2026` or `10-22-26`). Sheet "x OF y" copied from the reference. No instructor NOTES line (and no balloon that only points to a note). The logo comes from the template. A depth in a note (e.g. an engraving ".003 DEEP") is judged against the title block's general tolerance (HW#8A: .005 accepted as .003 ±.008), so don't make the user change the part for it. Units are inches (IPS) for drawings and MMGS when the part drawing says so.

---

## 2. VBA rules that prevent the worst bugs

| Rule | Why |
|---|---|
| Pass arrays to SolidWorks math **by value**: `swMath.CreatePoint((v))` (double parentheses). Keep a manual-transform fallback (`ManualXf`). | Without it every point came back (0,0,0) and every line collapsed to nothing. |
| A Variant array **element** passed to an `Object` parameter needs `ByVal` (`Sub AddEdge(ByVal ent As Object, ...)`). Pass the Variant copy, not a typed `Double()` array, to Variant parameters. | Otherwise VBA reports "ByRef argument type mismatch" and won't compile. |
| Never name a variable after a built-in (`sgn`, `len`, `left`, `date`...). | `Dim sgn` broke `Sgn()`. The static check catches this. |
| Integer division is a single `\`. A string path separator is `"\"`. Watch Python escaping when generating code. | `"\\"` in VBA is two characters, so InStrRev found nothing. |
| `MsgBox` shows about 1,000 characters. Show only the problems and write the full report to a `.txt` file. | Long reports were cut off. |
| **Reset module-level object variables at the start of `main`** (`Set viewDet = Nothing` ...). | VBA keeps them between runs in the same session, so a fresh drawing "had" a detail view from the last run. |
| Every API call that might not exist goes inside `On Error Resume Next`, with a fallback (e.g. `FeatureExtrusion3` → `FeatureExtrusion2`, `SaveAs3` → `SaveAs`). | SolidWorks versions differ. The user has **SOLIDWORKS 2026 Student**. |
| One `Check ok, label` per step, with `OK` / `PROBLEM` lines collected into `report` and `problems`. The final MsgBox shows counts and problems. | The user can say exactly what failed. |
| Report the **reason** on failure ("circle not drawn", "found 2 of 1 edge(s)", "measured 61.842"). | The reason found the real bug every time, e.g. a radian value or a double selection. |
| Find the open part by **name tag** (`"PART #4"`), not the active document. | The macro once grabbed Part #1 because the HW#7A drawing was active. |

---

## 3. Part macros (from HW#4C)

- **Coordinates:** draw every sketch entity from **model coordinates** converted with the active sketch's `ModelToSketchTransform`. Then plane orientation never matters.
- **Sketch mode:** draw with `AddToDB = False`, the normal mode, so coincident endpoints merge into closed profiles.
- **Extrude:** exit the sketch first with `InsertSketch True`, select it by name, then extrude. After each extrude, check its direction from the faces' bounding box (`EnsureExtent`) and reverse it if wrong.
- **Volume check after every feature.** Work out the expected volume change of each feature from the drawing, before writing code, with a CadQuery reference model (`build_cq.py` in hw4c). Compare after each step. If a boss went the wrong way, reverse it automatically. The final total must match the drawing, e.g. 779,719 mm³.
- **Fully defined sketches**, all observed in SOLIDWORKS 2026:
  - `GetConstrainedStatus` returns **2 = under defined, 3 = fully defined, 4 or more = over defined**. The documented enum is wrong for this install.
  - Relations that work: `sgHORIZONTAL2D` and `sgVERTICAL2D` on lines, `sgTANGENT`, `sgCONCENTRIC`, `sgATMIDDLE`.
  - **Avoid point-to-point alignment relations.** They failed silently. Instead, fix positions with **dimensions to a permanent model edge**, e.g. a horizontal 111 from the arc centre to the base plate's back-right edge.
  - After adding any relation or dimension, re-check the status. If it became over defined, `EditUndo2 1`.
- **Names:** features `Step 1..N` and sketches `Sketch N`, numbered by step.
- **Save:** use `WScript.Shell.SpecialFolders("Desktop")`, which handles a OneDrive Desktop (`OneDrive - Cal Poly`). If saving fails, fall back to Documents, then report the path and the error code.

---

## 4. Drawing macros (from HW#7A–7D)

### 4.1 Setup
- **Template:** search the usual places for `Hayden_p_ansi.drwdot`. That's OneDrive - Cal Poly\Documents, Documents, Downloads, Desktop and the OneDrive variants. Otherwise ask with InputBox ("Copy as path").
- **New drawing:** `NewDocument(tmpl, 12, 11 in, 8.5 in)`. Call `swDraw.EditSheet` right away.
- **Sheet properties:** read `GetProperties2`, then call `SetProperties2(paper, templ, scale1, scale2, firstAngle:=False, w, h, True)`. Check both the scale and third-angle projection.
- **Part orientation from geometry, never assumed.** Find characteristic cylinders by radius (`CylinderParams`) and planes (`PlaneParams`), then build the frame:
  - `vR` = right in the main view, `vU` = up, `vF` = toward the viewer, with `vR = vU × vF`.
  - Fix the signs using the face box centres of known features, e.g. "the boss is above the cross hole".
- **Part-frame origin and helpers:** compute the origin from `GetBodyBox` corners projected onto the frame. Then `PartPt(r, u, f)` gives a model point and `SheetPt(view, r, u, f)` gives the sheet position in inches via `view.ModelToViewTransform`.
- **Measure angled faces too.** For example, find the inclined plane, measure its angle and check it passes through the expected corner. Don't trust the image for angles.

### 4.2 Views
- **Model views:** try the standard named view first (`*Front`, `*Top`...). Check it with `ViewMatches` (screen-right = vR, screen-up = vU, from the view's transform). If it doesn't match, make a custom named view in the part (`ModelView.Orientation3` + `NameView`, trying both matrix conventions) and check again.
- **Projected views:** select the parent view, then `CreateUnfoldedViewAt3`. Check them with `FrameOK`.
- **Placing views:** SolidWorks places a view by its outline, which includes centre marks, so it lands up to 1 in off. **After creating each view, move it with `PlaceView`.** Pick a real part point, e.g. the left end on the centre line, and move it to the sheet position measured from the reference. Aligned child views follow their parent.
- **Display modes:** `SetDisplayMode3 False, mode, False, False`. Use 1 for hidden lines visible and 2 for hidden lines removed. Section, auxiliary and isometric views normally use 2.
- **Section view:** draw a line in the parent view (see 4.5). Select it with **`seg.Select4`**, never a screen pick, then call `CreateSectionViewAt5` (fallback `At4`). Check the direction. If it's wrong, use `GetSection.SetReversedCutDirection`.
- **Auxiliary view:** select the inclined edge in the parent view, then call `CreateAuxiliaryViewAt2` (fallback `At`). **It may build the view but return Nothing**, so record the view names first (`ViewNames`) and recover the new view with `NewViewSince`. Check that the face normal points at the viewer.
- **Detail view:** the API (`CreateDetailViewAt4` / `At3`) refused every combination, and the circle is hard to place (4.5). The approach that works:
  1. **Dimension the parent view BEFORE the detail exists.** Once a detail is taken from a view, selecting a parent edge also selects it in the detail, and dimensions are refused.
  2. Try to make the detail automatically, and keep it only if `DetailHasGroove` finds the expected edge inside it.
  3. If that fails, finish everything else, then tell the user to make it with the Detail View tool and run `main` again. **Continue mode** (`ResumeDrawing`) finds the existing views and adds only what's missing (`HasDims`, `SheetHasText`).
- **View labels and letters:** the letter of a section or auxiliary label, and the position of "SECTION B-B", **can't be changed by API**. Tell the user to drag the label or set the letter in the PropertyManager.

### 4.3 Picking edges (the core technique)
- **Edge table:** `CollectEdges(view)` reads `GetVisibleEntities2(Nothing, type)`. It tries types 2, 1, 3, 4 and keeps whichever returns curves. For each edge it stores:
  - straight lines, with end points from the **vertices**, or `GetCurveParams2` indices 0 to 5 (**indices 6 and 7 are not curve parameters**; using them blew the view outlines up);
  - circles and arcs seen face-on, with centre and radius;
  - **circles seen side-on, stored as straight lines**, e.g. the R1.00 top seen from above, or every edge of a turned part.
- **View outlines:** take them only from exact geometry (line end points, circle centre ± radius) or from the part box corners via `SheetPt`.
- **Select by part coordinates:** `SelPt(view, Array(r, u, f))` maps the point to the sheet and takes the **nearest table edge within 0.008 in**. Only if nothing is that close does it fall back to `SelectByID2 "EDGE"` and then `"SILHOUETTE"` at that sheet point.
- **Choose pick points that sit on one edge only:**
  - away from cutting lines, centrelines and sketch lines (an arc point on the B-B line picked the cutting line and made a 90° angle);
  - not where two edges overlap: for coincident side-on circles of Ø.75 and Ø.65, pick at r = .36, which only the Ø.75 edge reaches;
  - away from small chamfer offsets of .025.
- **Silhouette lines are not selectable by macro.** These are the outline lines of cylinders and cones in a side view. Measure diameters on the **side-on circular edge** instead (`DiaByEdge`, a vertical length of that edge), or leave the dimension for the user. 30° chamfer angles on turned parts need the cone outline, so they're usually done by hand.
- **Keep the selection inside one view:** after every pick, `KeepOneIn` drops selections from other views and keeps one new item. It is harmless and worth keeping, but it did **not** cure the parent/detail double selection. Dimensioning the parent first did (4.2).
- **When positions aren't known exactly, measure:** e.g. in VIEW D-D, find the edge .625 from the hole centre, then the edge 1.00 from that one. Every result is still value-checked.

### 4.4 Dimensions and notes
- **Creating dimensions:** `TryDim` selects, creates, formats, reads back and checks:
  - Calls: `AddHorizontalDimension2` / `AddVerticalDimension2` / `AddDiameterDimension2` / `AddRadialDimension2`, or `AddDimension2` for aligned dimensions, angles and a single side-on circle edge.
  - **Value check:** `dd.GetDimension2(0).SystemValue` is in metres (divide by 0.0254), or **radians for angles**. Allow ±0.003 in, or ±0.3°.
  - **Wrong value:** if a later attempt exists, delete it (`dd.GetAnnotation.Select3`, `EditDelete`) and try the next way.
- **Formatting:**
  - **Decimals:** `SetPrecision3 n, -1, -1, -1`. Match the reference per dimension (`1.500` = 3, `1.25` = 2).
  - **Prefix:** the prefix slot `SetText 1` **already holds the Ø or R**. Add in front: `"4 X " & existing`, or `<MOD-DIAM>` for a linear diameter, only if it's not already there. Replacing it lost the Ø and R.
  - **Suffix:** `SetText 2, " TYP"` or `" THRU"`.
  - **Parentheses:** `ShowParenthesis = True` for reference dimensions like `(.375)`.
- **Text position:** the X, Y passed is roughly where the dimension text goes, and SolidWorks draws the dimension line through it. Use reference offsets from part anchors, e.g. `xMax + 0.35`, `yTop + 0.48`.
- **Leader notes**, e.g. chamfer callouts: select the edge, then `InsertNote(".025 X 45<MOD-DEG> TYP")`, then `GetAnnotation.SetPosition2`. The note's anchor is the top-left of the text. If no edge was found, say so in the report.
- **Centre marks and centrelines:**
  - Centre mark: select a circle, then `InsertCenterMark3(0, False, True)`.
  - Centrelines: select the view, then `InsertCenterLine2`. **Do centrelines LAST.** The centreline command can stay active and swallow the clicks that follow.
- **Before each view's dimensions,** run `ClearCommand`: clear the selection and re-activate the sheet.

### 4.5 Sketching inside a drawing view (cutting lines, detail circles)
- The view sketch's coordinate system is **not reliably known**. The candidates are sheet coordinates, coordinates from the model origin, coordinates from the view centre (`v.Position`) and the view sketch's `ModelToSketchTransform`. Results differed by inches.
- **Try each one in turn,** and keep the one whose **result** proves right. For example, the section exists and faces the right way, or the detail contains the expected edge. Delete the attempts that failed.
- A long cutting line usually still works even if it's offset along its own length. A small detail circle does not.

### 4.6 Title block (Hayden's template)
- **Values to add:** the labels are fixed text with blank values. Find the `MATERIAL:`, `DATE:` and `TITLE:` notes (in the sheet first, then in sheet-format mode), take their positions, and add the values with offsets: material (+0.06, −0.17), date (+0.40, −0.18), title line 1 (+0.44, 0) and line 2 (+0.44, −0.17).
- **Fixed notes to change:** `SHEET 1 OF 1` and `SCALE: 1:1` are static notes. Change them with `SetText`, falling back to `PropertyLinkedText`, and enter `EditTemplate` if needed.
- **ALWAYS call `swDraw.EditSheet` afterwards.** Staying in sheet-format mode hides every view, and that produced a blank PDF.
- **Already in the template:** DRAWN BY HAYDEN PERHACH, REV B, SIZE A and the Cal Poly logo. The ASME line says 2018 (references say 2009); the user kept the template's.

### 4.7 Finish and save
- **`FinishSheet`:** activate the drawing, call `EditSheet`, show any hidden views, run `ForceRebuild3` and `GraphicsRedraw2`, then check the view count. Run it **before the report and again before the PDF export**.
- **Save:** ask for the last name (default Perhach). Save `HW#xx<Name>.SLDDRW` and `.PDF` next to the part, falling back to the Desktop. If the PDF fails, try `GetExportFileData(1)`. Write `HWxx_checks.txt` too. Never save or modify the part.

---

### 4.8 Lessons from HW#8A (errors to never repeat)

These are the problems found in HW#8A runs. Each lists what went wrong and what to do instead.

1. **A check reported a problem that wasn't there (the error the user flagged in v4).** `Gtol.SetFrameSymbols2` / `SetFrameSymbols3` returned failure in SOLIDWORKS 2026 although the // symbol was drawn. The summary said "symbol did not take" for a correct frame. **Never decide OK/PROBLEM from an API call's return value alone.** Read the result back (e.g. `GetFrameSymbols` / the frame's text, a dimension's value, a view's scale) and judge on that. If nothing can be read back, report it as `Info` ("look that // shows"), not as a PROBLEM. (In SW 2022+ these Gtol calls are documented as valid only for older Gtols.)
2. **Temporary dimensions were left on the sheet.** Test dimensions used to prove a sketched line or circle's position ("1.94" twice, above the front view) were not removed by `dd.GetAnnotation.Select3` + `EditDelete`. Count the view's display dimensions before and after and confirm the deletion, or delete by name (`SelectByID2 dd.GetNameForSelection, "DIMENSION"`). Better still, prove a sketch position without creating a dimension, e.g. from the segment's sketch points transformed to the sheet.
3. **SolidWorks auto-scaled new views to 1:2** although the sheet properties read 1:1. The template's sheet format is stored at 1:2, and the status bar still showed 1 : 2. After creating every view, set and read back its scale (`SetViewScale`: `ScaleRatio`, then `ScaleDecimal`). Also read the sheet scale back after a rebuild and set it again if it changed. Anything given in sheet inches must use the view scale: sketch radii, and edge-table radii `eR` (sheet size = model size × scale).
4. **The section label stayed** ("SECTION A-A" plus "SCALE 1 : 1", the scale line appearing because the view scale differed from the sheet's), overlapping the title block. Setting `Annotation.Visible` did not hide it. Keep the sheet scale equal to the view scales so no scale line appears, place the label clear of the title block, and list "right-click the label > Hide" under things for the user, not as a silent OK.
5. **Tolerance values silently stayed 0** (±.000, .080/.080). `tol.Type` stuck but `SetValues2(min, max, 2, Empty)` changed nothing. Try `SetValues2(min, max, 1, "")`, `SetValues2(..., 2, "")`, `SetValues(min, max)` in turn, and read back `GetMaxValue` / `GetMinValue` after each.
6. **Wrong symbol codes** render as literal text or a blank cell. Correct codes: counterbore `<HOLE-SPOT>`, countersink `<HOLE-SINK>`, depth `<HOLE-DEPTH>`, diameter `<MOD-DIAM>`, ± `<MOD-PM>`, degree `<MOD-DEG>`, parallelism `<IGTOL-PARA>` (the gtol codes are 4-letter: `<IGTOL-POSI>`, `<IGTOL-CONC>`). `<HOLE-CBORE>`, `<HOLE-CSINK>` and `<IGTOL-PARALLEL>` do not exist.
7a. **Compile error "Ambiguous name detected"** (caught before HW#8B shipped): a module variable `todo` and a procedure `ToDo` share a name, and VBA ignores case. `vba_static_check.py` now flags this too.
7. **Compile error "ByRef argument type mismatch"**: `CylSpan(faces(i))` passed a Variant array element to `fc As Object`. Declare such parameters `ByVal`. `vba_static_check.py` now flags this.
8. **Hole recognition**: blind-hole depths in Hole Wizard may run to the drill tip, so a depth window like .35 ± .03 failed. Measure a bore's length from the centres of its circular edges (not `GetBox`). Tell blind holes apart by relative depth (shallower vs deeper), and list every hole's measurements as `Info` lines in the check file.
9. **Circular-group centre marks** (`InsertCenterMark3(2, ...)`) gave plain + marks, not the reference's marks turned toward the centre. Still unsolved.
11. **Notes are included when the drawing refers to them** (HW#8A balloon "3", HW#8B notes 3 and 4). Place each number and its text as separate notes so the columns line up. Use `<MOD-DEG>` and `<HOLE-DEPTH>` in note text.
12. **Gtol codes:** circularity `<IGTOL-CIRC>`, circular runout `<IGTOL-SRUN>`, total runout `<IGTOL-TRUN>`, parallelism `<IGTOL-PARA>`, position `<IGTOL-POSI>`, concentricity `<IGTOL-CONC>`. The macro also looks the code up in the computer's gtol.sym (`ProgramData\SOLIDWORKS\SOLIDWORKS 20xx\lang\english`).
13. **A tolerance block that already matches the reference is left alone**: read it back, and change it only if it differs (HW#8B keeps ±.01 / ±.005).
14. **`GetBodyBox` is loose round modelled threads** (HW#8B: 4.020 for 4.000), so the origin and every pick point shifted. Take a turned part's ends from its flat end faces (`PlaneParams` with normal along the axis).
15. **"Already done?" tests must not match a substring**: note 1 contained "IN-LAB ASSIGNMENT #8B", so the title block was skipped. Compare whole note texts.
16. **Outline curves (silhouettes) can be dimensioned after all**: place sketch points, lines or arcs exactly on the outline with a proven sketch method (`FindCand`: a test point measured from a real edge), then dimension those and check the value. Used for Ø.775, (Ø.493), 30° and R.047 in HW#8B.
17. **Modelled threads make a view slow**: turn its hidden lines on only at the end, skip helix edges at once, read each view's edges once, and never pick on the sheet over the thread.
18. **Make every real-edge dimension in a view before adding any sketch item to it** (HW#8B v4 added test points first, and then the edge dimensions failed). Pick each edge two ways (edge table, then a click at its exact sheet position) and trim extra selections. Only then use exact sketch points as a value-checked fallback.
19. **Rebuild a newly created view before reading its edges** (`ForceRebuild3`). A new DETAIL A reported no edges, so a correct detail was judged wrong and deleted.
20. **On a heavy part (modelled threads) keep the number of SolidWorks operations small.** HW#8B v5 tried each dimension several ways and re-read the views' edges, and SolidWorks crashed after 20 minutes. Prove the sketch placement once per view, then dimension on exact sketch points, with single-click picks for leaders.
21. **Sketch points show as small x's that can't be clicked or hidden easily** (HW#8B). Prefer sketch lines, arcs and circles (they belong on the drawing anyway). Use points only where nothing else will do.
22. **A view's "toward you" sign**: prove it on a view whose facing direction is known (`ZSign` on the main view) before judging which side an auxiliary view looks from.
23. **BASIC dimensions** are tolerance type 1 (`swTolBASIC`), set and read back like any tolerance. A position frame's Ø comes from the `Dia1` argument of `SetFrameSymbols2`. Codes: position `<IGTOL-POSI>`, angularity `<IGTOL-ANGULAR>`.
24. **A local variable named like a procedure hides it** (HW#8C v1: `sp = SP(...)` indexed the empty variable, because VBA ignores case). `vba_static_check.py` now flags this.
25. **Capture `Err.Number` and `Err.Description` first in an error handler.** Clean-up calls that use `On Error Resume Next` / `Err.Clear` wipe them, and HW#8C v1 showed "Error 0". Give each big step its own handler so one failure doesn't stop the whole drawing.
10. **The MsgBox cut the problem list off.** Keep PROBLEM labels short (one line each) and ask the user for `HWxx_checks.txt` when there are several problems.

## 5. Things that need the user (say so up front)

- **Pick and run:** run `main` from the Macros dialog, never a helper. If several parts are open, close the others or click into the right part window.
- **View labels:** drag "SECTION B-B" below its dimensions. Change an auxiliary view's label letter (A → D) in the PropertyManager.
- **Detail views:** if the macro can't make one, draw the circle with the Detail View tool, set the label and the 2:1 scale, then run `main` again.
- **Silhouette-only dimensions** (cone angles, some diameters): use Smart Dimension on the outline lines, then add Ø.
- **Units:** check MMGS or IPS in the status bar before saving a part.

---

## 6. Pre-flight checklist before sending a macro

- [ ] The header states how to run it and what it builds, with a version note
- [ ] `PART_TAG`, `DRAW_NAME`, material, title line 2, sheet "x OF y", scale and date format are set
- [ ] The orientation comes from geometry, and a check confirms the part size
- [ ] Every view is checked for orientation and placed by a real edge with `PlaceView`
- [ ] Every pick point sits on one edge only, away from cutting lines, centrelines and coincident edges
- [ ] Every dimension has the expected value, decimals, prefix, suffix and parentheses from the reference
- [ ] The parent view is dimensioned before any detail view is created
- [ ] Centrelines come last, and `ClearCommand` runs before each view's dimensions
- [ ] The title block ends with `EditSheet`, and `FinishSheet` runs before the PDF
- [ ] View variables are reset in `main`
- [ ] Every OK/PROBLEM is judged on a value read back from SolidWorks, never on an API call's return value alone (4.8 #1)
- [ ] No temporary dimension or sketch item is left on the sheet: deletions are confirmed (4.8 #2)
- [ ] Every view's scale and the sheet scale are read back and match the reference; section labels show no extra "SCALE" line (4.8 #3, #4)
- [ ] Tolerance values are read back; symbol codes come from 4.8 #6
- [ ] `vba_static_check.py` prints OK

---

## 7. Starter prompt for a new chat

> I'm doing IME 1140 SolidWorks assignments. Read `solidworks_tools/SolidWorks_Macro_Playbook.md` in the repo `Leftyp76/rfid_attendance`, branch `claude/nice-hawking-yolktd`, first. If you can't reach the repo, use the copy I attached. Follow it exactly: its workflow, VBA rules, part- and drawing-macro techniques and pre-flight checklist. Base the new macro on the newest working macro in that branch (`solidworks_hw8a/HW8A_Drawing_Macro.bas`, or `solidworks_hw7d/HW7D_Drawing_Macro.bas` for detail views, or `solidworks_hw4c/HW4C_SolidWorks_Macro.bas` for a part), keeping its proven helpers unchanged, and avoid every error listed in section 4.8.
>
> This assignment: [paste the assignment text and attach the reference image, the .SLDPRT and Hayden_p_ansi.drwdot]. File name rule: [e.g. HW#8A + Last Name]. My last name is Perhach.
>
> Steps: extract the part data with `sw_extract.py`, read the reference image carefully, then give me a short plan and only the questions the reference doesn't answer. After I answer, build the macro, run `vba_static_check.py`, commit and push to the branch, and send me the `.bas` file. Tell me exactly which file to paste and to run `main`. I'll reply with the check summary and a screenshot of the sheet after each run.
