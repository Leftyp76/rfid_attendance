# SolidWorks Part, Drawing & Assembly Macro Playbook

How to build a SolidWorks part, engineering drawing or assembly from a reference sheet with a VBA macro, so it comes out right on the first run as often as possible. This is written for the AI assistant that writes the macro. Everything here was learned building IME 1140 HW#4C (a part), HW#7A–7D and HW#8A–8D (eight drawings). Section 8 prepares the next step, **assemblies**, which have not been built yet.

Repo: `Leftyp76/rfid_attendance`, branch **`claude/festive-clarke-9gu1sg`** (newest work: HW#8A–8D, the tools and this playbook). The older branch `claude/nice-hawking-yolktd` holds HW#4C–8B only.

| Folder | What is in it | Use it as the base for |
|---|---|---|
| `solidworks_hw4c/` | Part macro (arch bracket), fully defined sketches, volume checks, CadQuery reference model | Any **part**, and the parts of an assembly |
| `solidworks_hw7a/` | Drawing: front, top, right, isometric, 1:1 | Simple orthographic drawings |
| `solidworks_hw7b/` | Drawing with **section view** B-B | Section views |
| `solidworks_hw7c/` | Drawing with **auxiliary view** D-D at 1:2 | Auxiliary views, other scales |
| `solidworks_hw7d/` | Drawing with **detail view** E at 2:1, turned part | Detail views, revolved parts |
| `solidworks_hw8a/` | Drawing with **SECTION A-A**, hole callouts, toleranced dims (±, limit, bilateral), datum and parallelism frame, 1:3 shaded iso; plus a part macro that changes an engraving depth | Section views, tolerances, GD&T, hole callouts, changing a part |
| `solidworks_hw8b/` | Turned part with **DETAIL A 3:1**, left end view, NOTES block, circularity / runout frames, datum on a dimension | Turned parts, detail views, notes, heavy (threaded) parts |
| `solidworks_hw8c/` | Elbow casting: VIEW A-A square to an angled flange, BASIC (boxed) dims, position / angularity frames, datums on a face and on a dimension, DETAIL B 1:1, iso 1:4. **Ran (v3)**; the user finished VIEW A-A, datum A, (R6.000), DETAIL B and its note by hand (lesson 28) | GD&T, BASIC dims, sketch items in views, per-step error guards |
| `solidworks_hw8d/` | Plate with **SECTION A-A and B-B** (labels hidden), hole callouts (deep / counterbore / countersink) each with a position frame to A, B, C, BASIC dims, flatness and parallelism, datums on edges, balloon 3 on the logo, NOTES 1-3. **v1, not yet run.** `build/` holds its source pieces and `assemble.py` | Two section views, hole callouts with frames, balloons; **the build method in 4.9** |
| `solidworks_tools/` | `sw_extract.py` reads a SolidWorks file, `vba_static_check.py` checks a macro | Every job |

**Start every new drawing macro from the newest macro that has run** (now `HW8C_Drawing_Macro.bas` v3; its helpers are the most tested) and build it the HW#8D way (4.9): part-specific pieces plus helpers copied word for word by a script. Read section 4.8 first: the new macro must not repeat any error in it. Keep proven helpers unchanged and replace only the part-specific procedures.

---

## 1. Workflow (do it in this order)

1. **Collect the inputs.** These are the reference image, the `.SLDPRT` for a drawing (every `.SLDPRT` for an assembly), the drawing template `.DRWDOT`, the file-name rule (e.g. `HW#8D` + last name) and the user's last name (Perhach). Commit the reference image into the assignment folder as `HWxx_reference.png`.
2. **Read the part file.** Run `python3 -I solidworks_tools/sw_extract.py "Part.SLDPRT"`. It prints every feature with its sketch dimensions, e.g. `Sketch1: D1 4.75, D2 1.00 ...`, plus the units and material.
   - Values are shown to the part's display decimals: `0.13` is often `.125`, `0.26` may be `.257`, and a depth of `0` may really be `.003` (HW#8D's logo cuts). Check each against the reference, and have the macro measure anything that matters.
   - Hole Wizard features name the hole (`CBORE for 1/4 Socket Head Cap Screw1`, `#7 (0.201) Diameter Hole1`), which tells you which callout goes where.
3. **Read the reference image.** Crop and enlarge each view with Pillow, at 2× to 5×. List every view, every dimension and its text (prefix, suffix, decimals, parentheses, boxed = BASIC), every frame, datum, balloon and note. Look for typos the reference really shows (HW#8D: "2 X .125" and "2 X.125") and copy them.
4. **Calibrate pixels to inches.**
   - Find the border lines with a pixel scan (rows/columns that are mostly dark); the template's border is 10.5 × 8.0 in.
   - **Also measure px/in on the part itself** (a known length such as the 4.000 overall). On HW#8D the views measured about 2 % larger than the border suggests (82.5 vs 80.8 px/in).
   - Use the border only for the main view's absolute sheet position. Use the part's px/in for everything relative: the spacing between views and every annotation offset, measured **from a part feature**, never as an absolute sheet position.
5. **Plan, then build.** The user wants one-shot macros. Ask only what blocks you (a missing part file, a genuinely ambiguous requirement). Use the standing answers below for the rest, and state the defaults you used in the reply.
6. **Build the macro** from the base macro, following sections 3, 4 and 8.
7. **Run `python3 solidworks_tools/vba_static_check.py Macro.bas`** before sending it. It must print OK. Then re-read your own part-specific code once, adversarially: argument counts, pick points on exactly one edge, signs of every offset.
8. **Commit and push**, then **send the `.bas` file**. Tell the user exactly which file to use: paste it into Tools > Macro > New over everything in the module and run **`main`**, never a helper.
9. **Iterate on the report.** Every run ends with a check summary listing OK, PROBLEM and TO DO items, and writes `HWxx_checks.txt`. Ask for a screenshot of the sheet plus that file. Fix the root cause shown, never guess at several things at once. Record each new lesson in section 4.8 (or 8.6 for assemblies).

**Standing user answers:**
- **Date:** today's date, written in the same format as the reference (`10-28-26` or `10-29-2026`).
- **Sheet "x OF y" and SCALE note:** copied from the reference.
- **NOTES block: include it** with every numbered note, and a balloon that points to a note (HW#8A–8D). The user changed this answer after HW#8A.
- **Logo:** the Cal Poly logo comes from the template. An engraving depth in a note (".003 DEEP") is judged against the title block's general tolerance, so don't make the user change the part for it; the drawing macro measures it and reports.
- **Hand steps (TO DO):** at most 1–2 per macro. Anything else must be done by the macro or reported as a PROBLEM with its reason.
- **When the user asks how to do something by hand:** simple numbered steps in plain words, no code.
- **ASME line:** the template's (2018), although references say 2009.
- **Units:** inches (IPS) for drawings; MMGS when the part drawing says so.

---

## 2. VBA rules that prevent the worst bugs

| Rule | Why |
|---|---|
| Pass arrays to SolidWorks math **by value**: `swMath.CreatePoint((v))` (double parentheses). Keep a manual-transform fallback (`ManualXf`). | Without it every point came back (0,0,0) and every line collapsed to nothing. |
| A Variant array **element** passed to an `Object` parameter needs `ByVal` (`Sub AddEdge(ByVal ent As Object, ...)`). Pass the Variant copy, not a typed `Double()` array, to Variant parameters. | Otherwise VBA reports "ByRef argument type mismatch" and won't compile. |
| **VBA ignores case.** Never give a variable or parameter the name of a procedure (`sp` hides `SP()`), of a module variable (`todo` / `ToDo`), of a built-in (`sgn`, `len`, `left`, `date`) or of a keyword (`eNum` reads as `Enum`). | HW#8C v1 indexed an empty variable instead of calling `SP` ("Error 0"); v2 would not compile ("Syntax error"). The static check catches all four kinds. |
| Integer division is a single `\`. A string path separator is `"\"`. Watch Python escaping when generating code. | `"\\"` in VBA is two characters, so InStrRev found nothing. |
| `MsgBox` shows about 1,000 characters. Show only the problems and the TO DO list, and write the full report to a `.txt` file. | Long reports were cut off. |
| **Reset module-level object variables at the start of `main`** (`Set viewDet = Nothing` ...). | VBA keeps them between runs in the same session, so a fresh drawing "had" a detail view from the last run. |
| Every API call that might not exist goes inside `On Error Resume Next`, with a fallback (e.g. `FeatureExtrusion3` → `FeatureExtrusion2`, `SaveAs3` → `SaveAs`). | SolidWorks versions differ. The user has **SOLIDWORKS 2026 Student**. |
| **Every big step has its own error handler** (`On Error GoTo Oops` ... `Oops: Check False, curStep & " - stopped by error " ...`), and every handler **captures `Err.Number` / `Err.Description` first**. | One failure is reported and the rest of the drawing is still built; clean-up code had wiped the error (HW#8C). |
| One `Check ok, label` per item, with `OK` / `PROBLEM` lines collected into `report` and `problems`; `ToDo label` for hand steps; `Info label` for facts that are not judged. | The user can say exactly what failed. |
| Report the **reason** on failure ("circle not drawn", "found 2 of 1 edge(s)", "measured 61.842"). | The reason found the real bug every time. |
| Find the open part by **name tag** (`PartMatches`: "DRAWINGS II" and "PART 4"), not the active document. | The macro once grabbed Part #1 because the HW#7A drawing was active. |
| Speed: `swApp.CommandInProgress = True` and `FeatureManager.EnableFeatureTree = False` while the macro runs (always switched back, also in the error path); `EditRebuild3` rather than `ForceRebuild3` except where a view's contents must be fresh; no per-step redraws. | HW#8B v6 was very slow; v7 with these was much faster. |
| Turn off "Input dimension value" (`SetUserPreferenceToggle 10`) while dimensioning sketch items, and put it back (`NoDimDialog`). | Otherwise a Modify box pops up for every dimension. |

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
- **Changing an existing part** (HW#8A engrave macro): find the feature by what it does (e.g. every cut shallower than .02), change its dimension, rebuild, and check the result by geometry (the engraved floor faces now sit .003 below the pocket floor) and by volume. Ask before saving.
- **Save:** use `WScript.Shell.SpecialFolders("Desktop")`, which handles a OneDrive Desktop (`OneDrive - Cal Poly`). If saving fails, fall back to Documents, then report the path and the error code.

---

## 4. Drawing macros (from HW#7A–8D)

### 4.1 Setup
- **Template:** search the usual places for `Hayden_p_ansi.drwdot`. That's OneDrive - Cal Poly\Documents, Documents, Downloads, Desktop and the OneDrive variants. Otherwise ask with InputBox ("Copy as path").
- **New drawing:** `NewDocument(tmpl, 12, 11 in, 8.5 in)`. Call `swDraw.EditSheet` right away.
- **Sheet properties:** read `GetProperties2`, then call `SetProperties2(paper, templ, scale1, scale2, firstAngle:=False, w, h, True)`. Check both the scale and third-angle projection. Set it again after the first view goes in (`EnsureSheetScale`): auto-scale can change it.
- **Part orientation from geometry, never assumed.** Find characteristic features by shape, then build the frame:
  - `vR` = right in the main view, `vU` = up, `vF` = toward the viewer, with `vR = vU × vF` (so `vU = vF × vR`).
  - Recognise features by radius (`CylinderParams`), by plane normal (`PlaneParams`), by torus (`TorusParams`: bend radius and pipe radius, HW#8C) and by cones (countersinks, drill points).
  - Fix the signs with features that only fit one way: "the counterbore is on the platform side", "the countersunk end of the hole row is nearer the R.50 notch" (HW#8D), "the round flange's normal points away from the pipe" (HW#8C).
- **Part frame origin:** take it from **flat faces** (the smallest `Dot(planePoint, axis)` of planes square to each axis), never from `GetBodyBox` (loose round threads, lesson 14). Then `PartPt(x, u, z)` gives a model point and `SheetPt(view, x, u, z)` the sheet position in inches via `view.ModelToViewTransform`.
- **Check every feature the drawing dimensions** against the reference in `FindOrientation` (overall size, each hole position, step heights, pocket depth, groove position) and report differences by name. If the part differs, the user learns it before looking at the sheet.
- **Measure angled faces too** (HW#8C measured the 45° between flange normals and R6.000 from the torus). Don't trust the image for angles.

### 4.2 Views
- **Model views:** try the standard named view first (`*Front`, `*Top`...). Check it with `ViewMatches` (screen-right = vR, screen-up = vU, from the view's transform). If it doesn't match, make a custom named view in the part (`ModelView.Orientation3` + `NameView`, trying both matrix conventions) and check again (`OrientedView`).
- **Isometric:** `dv = Unit(vR + vU + vF)`, up = vU made square to dv, right = up × dv; standard `*Isometric` only when vR, vU, vF are X, Y, Z.
- **Projected views:** select the parent view, then `CreateUnfoldedViewAt3`. Check them with `FrameOK`.
- **Placing views:** SolidWorks places a view by its outline, which includes centre marks, so it lands up to 1 in off. **After creating each view, move it with `PlaceView`** so a real part point lands on the measured sheet position (to .001 in). Aligned child views move only along their free axis (`doX` / `doY`). `CheckScalesAndAlignment` reads every position back at the end.
- **Display modes:** `SetDisplayMode3 False, mode, False, False`. Use 1 for hidden lines visible and 2 for hidden lines removed. Look at the reference: dashed lines just inside an outline mean hidden lines visible (HW#8D main view).
- **Section views** (see also lessons 26, 27):
  1. Dimension the parent view first; the cutting line is a sketch item (lesson 18).
  2. Draw the cutting line in the parent view (`SkLineP`, 4.5), longer than the part at both ends so the cut goes fully through, and prove its position with a test dimension that is deleted again (`CutLine`).
  3. Select the line itself with **`seg.Select4`**, never a screen pick, then `CreateSectionViewAt5(x, y, 0, "A", opt, Nothing, 0)` (fallback `At4`, then `NewViewSince`).
  4. Check the direction on the view's own axes (`FrameOK`); if wrong, `GetSection.SetReversedCutDirection`. Read the letter back (`GetSection.GetLabel`, `SetLabel2`).
  5. Set the scale, the display mode, then `PlaceView` along the free axis.
- **"VIEW A-A" with an arrow and a letter at both ends of a cutting line is a section view**, even when it looks like a plain projection (lesson 26).
- **Auxiliary view** (single arrow, "VIEW A"): select the inclined edge in the parent view, then `CreateAuxiliaryViewAt2` (fallback `At`). It may build the view but return Nothing, so record the view names first (`ViewNames`) and recover the new view with `NewViewSince`. Check that the face normal points at the viewer (`ZSign`).
- **Detail view:**
  1. **Dimension the parent view BEFORE the detail exists.** Once a detail is taken from a view, selecting a parent edge also selects it in the detail.
  2. Draw the circle with a proven sketch method (`DrawViewCircle`), select it, `CreateDetailViewAt4` (several style combinations), then **check the detail shows the expected feature** (its hole at the expected place) and report clearly if not. HW#8C's detail came out round the wrong spot (lesson 28).
  3. If SolidWorks refuses, that is the one TO DO: make it with the Detail View tool.
- **View labels:** hide "SECTION A-A" when the reference shows none (`HideSectionLabels`: try `Annotation.Visible` values and read back; otherwise a TO DO "right-click > Hide"). To make "SECTION A-A" read "VIEW A-A", the user double-clicks the label and edits the word.

### 4.3 Picking edges (the core technique)
- **Edge table:** `CollectEdges(view)` reads `GetVisibleEntities2(Nothing, type)`. It tries types 2, 1, 3, 4 and keeps whichever returns curves. For each edge it stores:
  - straight lines, with end points from the **vertices**, or `GetCurveParams2` indices 0 to 5 (**indices 6 and 7 are not curve parameters**);
  - circles and arcs seen face-on, with centre and radius;
  - **circles seen side-on, stored as straight lines**;
  - nothing for splines and helixes (skipped at once, for speed).
  - The table is cached per view (`eKey`: name, position, scale, display mode). Set `eKey = ""` after anything that changes the view.
- **View outlines:** take them only from exact geometry (line end points, circle centre ± radius) or from part points via `SheetPt`.
- **Select by part coordinates:** `SelPt(view, Array(x, u, z))` maps the point to the sheet and takes the **nearest table edge within 0.008 in**. Only if nothing is that close does it fall back to `SelectByID2 "EDGE"` and then `"SILHOUETTE"`. `PickOnly` is a single click at the exact point with no table read (fast, for leaders).
- **Choose pick points that sit on one edge only:**
  - away from cutting lines, centrelines and sketch lines (an arc point on the B-B line picked the cutting line);
  - not where two edges overlap: pick where only the wanted edge reaches;
  - **exactly on the wanted edge when a parallel edge is close** (a .01 chamfer, a .025 hidden back chamfer): the nearest edge wins, so put the point exactly on it;
  - in a section view, on the cut outline, away from the hatch.
- **A circle picks its centre:** a linear dimension from a line to a hole's circle measures to the hole centre, so pick any point on the circle (its top point, away from cutting lines).
- **Edges with no table entry** (logo letters, splines): walk `VisibleEdges(view)` directly, read each edge's start vertex, convert it to part inches with `ToPart`, keep the one nearest the target, and select it with `view.SelectEntity` (HW#8D balloon on the logo).
- **Silhouette lines** (outline of a cylinder or cone in a side view) are not real edges. Dimension them via exact sketch items placed on them (lesson 16), or on the side-on circular edge.
- **Keep the selection inside one view:** after every pick, `KeepOneIn` drops selections from other views and keeps one new item.

### 4.4 Dimensions, GD&T, notes and balloons
- **Creating dimensions:** `TryDim` / `TryDimSel` select, create, format, read back and check:
  - Calls: `AddHorizontalDimension2` / `AddVerticalDimension2` / `AddDiameterDimension2` / `AddRadialDimension2`, or `AddDimension2` for aligned dimensions, angles and a single side-on circle edge.
  - **Value check:** `dd.GetDimension2(0).SystemValue` is in metres (divide by 0.0254), or **radians for angles**. Allow ±0.003 in, or ±0.3°.
  - **Wrong value:** delete it and confirm the deletion (`DeleteDim`, lesson 2), then try the next way: edge table first, single clicks second (`DimP2`). `DimT` adds the tolerance and one check line.
- **Formatting:**
  - **Decimals:** `SetPrecision3 n, -1, tolDecimals, -1`. Match the reference per dimension (`1.1875` = 4, `1.500` = 3, `R.50` = 2).
  - **Prefix:** the prefix slot `SetText 1` **already holds the Ø or R**. Add in front: `"4X " & existing`. Replacing it lost the Ø and R. For a hole callout written "Ø .257" set slot 1 to `"<MOD-DIAM> "`.
  - **Suffix:** `SetText 2`, e.g. `" THRU"`, or a whole hole callout: `" <HOLE-DEPTH> .38"`, `" <HOLE-SPOT> <MOD-DIAM> .44 <HOLE-DEPTH> .25"`, `" <HOLE-SINK> <MOD-DIAM> .50 X 82<MOD-DEG>"`.
  - **Parentheses:** `ShowParenthesis = True` for reference dimensions like `(R6.000)`.
  - **Tolerances** (`SetTol`, lesson 5): BASIC 1, BILAT 2, LIMIT 3, SYM 4; values tried several ways until they read back.
- **Text position:** the X, Y passed is roughly where the text goes; SolidWorks draws the dimension line through it. Use reference offsets from a part anchor (`o = SP(view, Array(0,0,0))`, then `o(0) + 1.83`).
- **Feature control frames** (`MakeFrame3`): `InsertGtol`, then `SetFrameSymbols2 1, sym, dia, ...` and `SetFrameValues2 1, tol, "", datumA, datumB, datumC`.
  - **Under a dimension:** select the dimension's annotation first, so the frame attaches under its text (HW#8C 5 X Ø.257 and HW#8D hole callouts).
  - **On a face:** pick the edge first, so the frame gets a leader to it.
  - Judge it on `GtolReadBack`; if SolidWorks reports nothing, write an `Info`, not a PROBLEM (lesson 1).
- **Datum feature symbols:** on an edge (`DatumOnEdge`: pick, `InsertDatumTag2`, `SetLabel`, position; the extension line is drawn automatically when the tag sits beyond the edge) or on a dimension (`DatumOnDim`). Read the label back.
- **Leader notes**, e.g. chamfer callouts: pick the edge, then `InsertNote(".025 X 45<MOD-DEG>")`, then `GetAnnotation.SetPosition2`. The note's anchor is the top-left of the text.
- **Balloon to a note** (HW#8D "3"): pick the target edge, `InsertNote("3")`, `SetBalloon 1, 2` (circular, two-character size), position it, read back `GetBalloonStyle`.
- **NOTES block:** each number and each text line is its own note (`PlaceNote`) so the columns line up: `NOTE_X .45`, `NOTE_Y 8.14`, line step .167, text tab .41 on Hayden's template. Use `""` for an inch mark in a VBA string (`.003"" DEEP`).
- **Centre marks:** select a circle, then `InsertCenterMark3(0, False, True)`, one per hole (dedupe by centre).
- **Centrelines:** select the view, then `InsertCenterLine2`. **Do centrelines LAST**; the command can stay active and swallow the clicks that follow.
- **Before each view's dimensions,** run `ClearCommand`: clear the selection and re-activate the sheet.

### 4.5 Sketching inside a drawing view (cutting lines, circles, centre-line arcs)
- The view sketch's coordinate system is **not reliably known**. The candidates are: 0 the view sketch's `ModelToSketchTransform`, 1 sheet coordinates, 2 from the model origin, 3 from the view centre (`v.Position`). Results differed by inches.
- **Prove the method once per view** with a test item measured from a real edge (`FindCandP`, `CutLine`, `FindCandAA`), delete the test dimension and confirm, then reuse the method (`skCand`) for every sketch item in that view.
- Sketch items that belong on the drawing: cutting lines, bolt circles, angle lines, centre-line arcs (HW#8C (R6.000)), detail circles. Make them with `AddToDB = True`, and give centre lines the Center style (`CentreStyle`).
- **Sketch items show blue** (under-defined colour) on screen. They print black in most setups; if the user wants them black on screen: Line Format > Line Color, or Fix them only if no driving dimension is on them.

### 4.6 Title block (Hayden's template)
- **Values to add:** the labels are fixed text with blank values. Find the `MATERIAL:`, `DATE:` and `TITLE:` notes (in the sheet first, then in sheet-format mode), take their positions, and add the values with offsets: material (+0.06, −0.17), date (+0.40, −0.18), title line 1 (+0.44, 0) and line 2 (+0.44, −0.17).
- **Fixed notes to change:** `SHEET 1 OF 1` and `SCALE: 1:1` are static notes. Change them with `SetText`, falling back to `PropertyLinkedText`, and enter `EditTemplate` if needed.
- **Tolerance block:** make the TWO / THREE PLACE lines match the reference (`FixTol`: ±.01 with ±.005 or ±.008), and read them back. Leave a block that already matches.
- **ALWAYS call `swDraw.EditSheet` afterwards.** Staying in sheet-format mode hides every view, and that produced a blank PDF.
- **Already in the template:** DRAWN BY HAYDEN PERHACH, REV B, SIZE A and the Cal Poly logo.

### 4.7 Finish and save
- **`FinishSheet`:** activate the drawing, call `EditSheet`, show any hidden views, run `ForceRebuild3` and `GraphicsRedraw2`, then check the view count. Run it **before the report and again before the PDF export**.
- **Save:** ask for the last name (default Perhach). Save `HW#xx<Name>.SLDDRW` and `.PDF` next to the part, falling back to the Desktop. If the PDF fails, try `GetExportFileData(1)`. Write `HWxx_checks.txt` too. Never save or modify the part.

### 4.8 Lessons (errors never to repeat)

Each lists what went wrong and what to do instead. The job number says where it was found.

1. **A check reported a problem that wasn't there** (HW#8A, flagged by the user). `Gtol.SetFrameSymbols2` / `SetFrameSymbols3` returned failure in SOLIDWORKS 2026 although the // symbol was drawn. **Never decide OK/PROBLEM from an API call's return value alone.** Read the result back (a frame's values, a dimension's value, a view's scale) and judge on that. If nothing can be read back, report it as `Info`, not as a PROBLEM.
2. **Temporary dimensions were left on the sheet** (HW#8A). Test dimensions were not removed by `Select3` + `EditDelete`. Count the sheet's dimensions before and after and confirm the deletion, else delete by name (`SelectByID2 dd.GetNameForSelection, "DIMENSION"`) (`DeleteDim`).
3. **SolidWorks auto-scaled new views to 1:2** although the sheet read 1:1 (HW#8A). After creating every view, set and read back its scale (`SetViewScale`). Read the sheet scale back too. Anything in sheet inches must use the view scale: sketch radii and edge-table radii (sheet size = model size × scale).
4. **A section label showed an extra "SCALE 1 : 1" line and overlapped the title block** (HW#8A). Keep view scales equal to the sheet scale where the reference does, place labels clear of the title block, and if hiding fails, list "right-click the label > Hide" as the TO DO.
5. **Tolerance values silently stayed 0** (HW#8A). Try `SetValues2(min, max, 1, "")`, `SetValues2(..., 2, "")`, `SetValues(min, max)` in turn, reading back `GetMaxValue` / `GetMinValue` after each.
6. **Wrong symbol codes** render as literal text or a blank cell. Hole and text codes: counterbore `<HOLE-SPOT>`, countersink `<HOLE-SINK>`, depth `<HOLE-DEPTH>`, diameter `<MOD-DIAM>`, ± `<MOD-PM>`, degree `<MOD-DEG>`. `<HOLE-CBORE>`, `<HOLE-CSINK>` and `<IGTOL-PARALLEL>` do not exist.
7. **Gtol codes:** flatness `<IGTOL-FLAT>`, circularity `<IGTOL-CIRC>`, parallelism `<IGTOL-PARA>`, angularity `<IGTOL-ANGULAR>`, position `<IGTOL-POSI>`, concentricity `<IGTOL-CONC>`, circular runout `<IGTOL-SRUN>`, total runout `<IGTOL-TRUN>`. The macro also looks the code up in the computer's gtol.sym (`ProgramData\SOLIDWORKS\SOLIDWORKS 20xx\lang\english`).
8. **Compile errors that VBA only shows when it runs** — all now caught by `vba_static_check.py`:
   - "ByRef argument type mismatch": a Variant array element passed to `fc As Object` (HW#8A). Declare such parameters `ByVal`.
   - "Ambiguous name detected": module variable `todo` and procedure `ToDo` (HW#8B).
   - "Error 0" at run time: local `sp` hid the function `SP` (HW#8C v1).
   - "Syntax error": a variable named `eNum`, which reads as the keyword `Enum` (HW#8C v2).
9. **Hole recognition** (HW#8A): blind-hole depths in Hole Wizard may run to the drill tip. Measure a bore's length from the centres of its circular edges, tell blind holes apart by relative depth, and list every hole's measurements as `Info`. Accept a radius window that covers the display rounding (Ø.253–.269 for ".26").
10. **Circular-group centre marks** (`InsertCenterMark3(2, ...)`) gave plain + marks, not marks turned toward the centre. Still unsolved.
11. **The MsgBox cut the problem list off.** Keep PROBLEM labels short and ask for `HWxx_checks.txt`.
12. **Notes are included when the drawing refers to them.** Place each number and its text as separate notes.
13. **A tolerance block that already matches the reference is left alone**: read it back and change it only if it differs.
14. **`GetBodyBox` is loose round modelled threads** (HW#8B: 4.020 for 4.000). Take extents from flat faces.
15. **"Already done?" tests must not match a substring** (HW#8B: the title block was skipped because note 1 contained the title). Compare whole note texts.
16. **Outline curves (silhouettes) can be dimensioned after all**: place sketch items exactly on the outline with a proven sketch method, dimension those and check the value (HW#8B Ø.775, (Ø.493), 30°, R.047).
17. **Modelled threads make a view slow**: turn its hidden lines on only at the end, skip helix edges at once, read each view's edges once, and never pick on the sheet over the thread.
18. **Make every real-edge dimension in a view before adding any sketch item to it** (HW#8B v4 added test points first, and then the edge dimensions failed). Cutting lines are sketch items too.
19. **Rebuild a newly created view before reading its edges.** A new DETAIL A reported no edges, so a correct detail was judged wrong and deleted.
20. **On a heavy part keep the number of SolidWorks operations small.** HW#8B v5 tried each dimension several ways and re-read views, and SolidWorks crashed after 20 minutes.
21. **Sketch points show as small x's that can't be clicked or hidden easily** (HW#8B). Prefer sketch lines, arcs and circles.
22. **A view's "toward you" sign**: prove it on a view whose facing direction is known (`ZSign` on the main view) before judging which side another view looks from.
23. **BASIC dimensions** are tolerance type 1, set and read back like any tolerance. A position frame's Ø comes from the `Dia1` argument of `SetFrameSymbols2`.
24. **Capture `Err.Number` and `Err.Description` first in an error handler**, and give each big step its own handler (HW#8C).
25. **The user's hand-fix steps are part of the job.** When something is left for the user, the TO DO line says exactly where to click; when the user asks, answer in simple numbered steps.
26. **"VIEW A-A" with an arrow and a letter at BOTH ends of a cutting line is a section view** (HW#8C). The macro tried `CreateAuxiliaryViewAt2`, SolidWorks refused, and the fallback plain view had no arrows or label. The user rebuilt it as a section view by hand: line parallel to the flange face, Section View, hatch set to None, label edited to VIEW A-A.
27. **A cutting line that only clips the part makes a "partial section" with a stray hatched piece** (HW#8C, by hand: the orange box). Draw the line right across the part, longer at both ends. If a stray piece appears, crop the section view round the wanted picture (sketch a rectangle in the view, Drawing > Crop View).
28. **HW#8C items that failed on the run** (cause not reported, because the check file wasn't sent): datum A on a face did not appear; datum B (on a dimension) landed at the far left edge of the sheet; (R6.000) was missing; the DETAIL B circle sat on the bore edge instead of the hole, so the detail showed the wrong spot; "4 X Ø.375 THRU" was missing. What to do from now on:
    - read back every annotation's position after `SetPosition2` and report it if it is off the sheet or far from its target;
    - check a detail view's contents by geometry (the expected hole inside it) and say plainly when it is wrong;
    - ask for `HWxx_checks.txt` with the first screenshot, every time.
29. **Logo engraving depth**: the extract shows depths to display decimals (`0` for .003). Measure the engraved faces below the pocket floor in `FindOrientation`, and judge them against the note and the title block's tolerance (HW#8A, HW#8D).

### 4.9 How to build a new drawing macro (the HW#8D method)

1. Make a `build/` folder in the assignment folder with the part-specific pieces:
   - `p1_head.bas`: header (how to run, what it builds, version notes), globals, constants (layout measured from the reference), `main`, `PartMatches`, `FindOrientation` with every check;
   - `p2_views.bas`: views, sections, labels, centre marks, alignment check;
   - `p3_dims.bas`: each view's dimensions, frames, datums, balloons;
   - `p4_notes.bas`: the NOTES block.
2. Copy `solidworks_hw8d/build/assemble.py`. It pulls the proven helpers **by name, word for word**, from the newest macro that has run, applies only listed text changes (each asserted to match a given number of times), and writes the `.bas` with Windows line ends.
3. Run it, then `vba_static_check.py`, then read your part-specific code once more.

This keeps every proven helper unchanged, makes the part-specific code small enough to check, and makes a fix easy to rebuild.

---

## 5. Things that need the user (say so up front)

- **Pick and run:** run `main` from the Macros dialog, never a helper. If several parts are open, close the others or click into the right part window. Close any half-built drawing without saving before running again.
- **View labels:** hide or edit a label the API couldn't change (right-click > Hide; double-click to edit the text).
- **Detail views:** if the macro can't make one, draw the circle with the Detail View tool, set the label and scale.
- **A PropertyManager left open** (a section or detail view selected): click the green ✓ or press Esc before clicking elsewhere.
- **Units:** check MMGS or IPS in the status bar before saving a part.

---

## 6. Pre-flight checklist before sending a macro

- [ ] The header states how to run it and what it builds, with a version note
- [ ] Part name tag, `DRAW_NAME`, material, title line 2, sheet "x OF y", SCALE note, date format and tolerance block match the reference
- [ ] The orientation comes from geometry, and checks confirm the part size and every dimensioned feature's position
- [ ] Every view is checked for orientation and scale, and placed by a part point with `PlaceView`
- [ ] Every pick point sits on exactly one edge, away from cutting lines, centrelines and close parallel edges
- [ ] Every dimension has the expected value, decimals, prefix, suffix, parentheses and tolerance (BASIC boxed) from the reference
- [ ] Every frame, datum, balloon and note on the reference is made, attached and read back
- [ ] Real-edge dimensions come before any sketch item (cutting lines included) in each view; the parent view is dimensioned before any detail view
- [ ] Section vs auxiliary decided from the arrows (both ends = section); section directions checked with `FrameOK`
- [ ] Centrelines come last, and `ClearCommand` runs before each view's dimensions
- [ ] The title block ends with `EditSheet`, and `FinishSheet` runs before the PDF
- [ ] View variables are reset in `main`; every big step has its own error handler that captures `Err` first
- [ ] Every OK/PROBLEM is judged on a value read back from SolidWorks (lesson 1); no test item left on the sheet (lesson 2)
- [ ] At most 1–2 TO DO items
- [ ] `vba_static_check.py` prints OK

---

## 7. Starter prompt for a new chat

> I'm doing IME 1140 SolidWorks assignments. Read `solidworks_tools/SolidWorks_Macro_Playbook.md` in the repo `Leftyp76/rfid_attendance`, branch `claude/festive-clarke-9gu1sg`, first. If you can't reach the repo, use the copy I attached. Follow it exactly: its workflow, VBA rules, techniques and pre-flight checklist. Build drawing macros the HW#8D way (section 4.9), with helpers copied from the newest macro that has run (`solidworks_hw8c/HW8C_Drawing_Macro.bas`); use `solidworks_hw4c/HW4C_SolidWorks_Macro.bas` for a part and section 8 for an assembly. Avoid every error listed in section 4.8.
>
> This assignment: [paste the assignment text and attach the reference image(s), every .SLDPRT (and .SLDASM if given) and Hayden_p_ansi.drwdot]. File name rule: [e.g. HW#9A + Last Name]. My last name is Perhach.
>
> Steps: extract the data with `sw_extract.py`, read the reference image carefully, ask only what blocks you, build the macro, run `vba_static_check.py`, commit and push to the branch, and send me the `.bas` file. Tell me exactly which file to paste and to run `main`. I'll reply with a screenshot and the checks file after each run.

---

## 8. Assemblies (next: not built yet — verify every call on the first run)

Nothing in this section has run yet. It carries over what worked for parts and drawings, and lists the API calls to try. Judge every step on what reads back (lesson 1), give each step its own error handler, and record what really happens in 8.6.

### 8.1 Inputs and questions
- Every part file (`.SLDPRT`), the assembly reference (picture, exploded view, assembly drawing with BOM), the file-name rule, and the template if a drawing is wanted.
- Run `sw_extract.py` on every part (and on a given `.SLDASM`; it has not been tried on one). Note each part's key faces: holes, pins, slots, mating flats.
- Questions only if the reference doesn't settle them: which part is fixed (normally the base / first part), whether mates must be specific types (the grader may open the mate folder), whether an exploded view and an assembly drawing with BOM and balloons are wanted.

### 8.2 Building the assembly
- **Template:** `swApp.GetUserPreferenceStringValue(swDefaultTemplateAssembly)` (enum value 9; check), else search for an `.asmdot`, else ask with InputBox. `NewDocument(tmpl, 0, 0, 0)`; check `GetType = 2` (assembly).
- **Open each part first** (silently, read-only use): `swApp.OpenDoc6(path, 1, 1, "", errs, warns)` (1 = part, 1 = silent). Never save or change a part.
- **Insert:** `swAssy.AddComponent5(path, 0, "", False, "", x, y, z)` (fallback `AddComponent4`). It returns a `Component2`; check it is not Nothing and that its name reads back.
- **First component:** fixed at the origin with its planes aligned (insert at 0,0,0, then select it and `FixComponent`), or mated to the assembly planes if the reference shows that.
- **Finding mate faces:** do it like `FindOrientation` — recognise each face by shape in the **part** (cylinder radius, plane normal, position in that part's own frame), then get the same face in the assembly. A face taken from `comp.GetBody` is already in the assembly context; its surface parameters are in the part's coordinates, so map points with `comp.Transform2` before comparing positions. Alternatively `comp.GetCorrespondingEntity(partFace)` (check which works).
- **Mates:** select the two entities with `Entity.Select4 True, selData`, then `swAssy.AddMate5(type, align, flip, dist, ...)` (fallback `AddMate4` / `AddMate3`), or the newer `CreateMateData(type)` + `CreateMate`. Types: coincident 0, concentric 1, perpendicular 2, parallel 3, tangent 4, distance 5, angle 6. Align: aligned 0, anti-aligned 1, closest 2. Read the mate's error status back (`Feature.GetErrorCode2`) and its type.
- **After every mate, check the result by geometry:** the moved part's key point (a hole centre, a face) is where the reference puts it, measured through `comp.Transform2`. If it flipped, change the alignment and try again (like reversing an extrude).
- **Name things:** mates `Mate N - part A hole to part B pin`, so the tree reads like a recipe.

### 8.3 Checks for an assembly
- Every component present, none suppressed, the count matches the BOM.
- Every mate without error, and each component's degrees of freedom as the reference implies (fully defined unless it should move; read the component's constrained status).
- Key positions by geometry (lesson 1), the total mass or volume (`Extension.CreateMassProperty`), and interference detection (`InterferenceDetectionManager`) reporting none where parts should not overlap.
- The check file `HWxx_checks.txt` lists each component and mate with OK / PROBLEM, exactly as for drawings.

### 8.4 Assembly drawings, exploded views, BOMs
- Model views come from the assembly path exactly like part drawings (4.2); `PartMatches` becomes an assembly-name match.
- **Exploded view:** make the explode steps in the assembly (direction and distance per component), then show the exploded state in the drawing view (`view.ShowExploded`). Check positions by geometry.
- **BOM:** `view.InsertBomTable4(...)` with the template's BOM template, then read its rows back (item numbers, part names, quantities) against the reference.
- **Balloons:** `AutoBalloon5` on the view, or one balloon per component with a leader to one of its edges (4.4), reading each item number back.

### 8.5 Saving an assembly
- Save the assembly as `HW#xx<Name>.SLDASM` **in the same folder as its parts** (references stay relative). Never save the parts.
- **Canvas needs the parts too:** an assembly file only references its parts. Tell the user to upload every part with it, or use Pack and Go (`Extension.GetPackAndGo`) to a zip of the assembly with its parts.

### 8.6 Assembly lessons (fill in as they are learned)
- (none yet)
