' =====================================================================================
'  HW#8D  -  IME 1140 Engineering Drawings II, In-Lab Assignment #8D  -  DRAWING macro
'            (inches, third angle, every view 1:1)
'
'  HOW TO RUN
'   1. Open "IME 1140 Engineering Drawings II In-Lab Assignment Part 4.SLDPRT" in SolidWorks
'      (close the other parts).
'   2. Tools > Macro > New...  save it as  HW8D_Drawing.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main.  The current step shows in SolidWorks' status bar.
'   5. At the end enter your last name; it saves  HW#8D<LastName>.SLDDRW  and  .PDF  next to
'      the part, plus HW8D_checks.txt with every check.
'
'  WHAT IT BUILDS (matches the reference sheet)
'   - Main view: logo pocket on the left, the three dia .257 holes in a row, R.50 notch lower
'     right; hidden lines visible.  Cutting lines A-A (through the pocket) and B-B (through
'     the holes).
'   - SECTION B-B above it, SECTION A-A to its right (labels hidden, as on the reference),
'     isometric top right
'   - Main:  4.000, 3.375 / 2.750 / 2.125 / 1.250 BASIC, 1.875, .500, .250, 2.500, 3.750,
'            4X R.325, 4X R.250, R.50, dia .201, 1/4-20 UNC, datums B and C, balloon 3 on the logo,
'            hole callouts  dia .257 deep .38 / cbore dia .44 deep .25 / csk dia .50 X 82 deg,
'            each with position dia .01 | A | B | C
'   - B-B:   .125, .125, 1.1875, 1.563, 3.750, .200, .480, datum A, flatness .002,
'            parallelism .002 | A, .025 X 45 deg, hole centre lines
'   - A-A:   1.625, 1.375, 2 X .125, 2 X.125, .125
'   - NOTES 1-3; title block 6061-T6 ALUMINUM, today's date (mm-dd-yyyy), #8D, SHEET 4 OF 4,
'     SCALE: 1:1; tolerance block THREE PLACE +-.005 (as on the reference)
'
'  NOTE 3 (logo .003 deep): the part is never changed. The macro measures how deep the logo is
'   engraved and judges it against .003 and the title block's THREE PLACE tolerance (+-.005).
'
'  BUILT FROM EVERY LESSON SO FAR (playbook 4.8)
'   - orientation measured from the part (the hole row, the counterbore, the notch); every
'     feature position checked against the reference before drawing
'   - real-edge dimensions in the main view BEFORE the cutting lines (sketch items) go in;
'     the cutting lines' placement is proved with a test dimension that is then deleted
'   - per-step error guards: one failure is reported and the rest of the drawing is built
'   - OK / PROBLEM judged on read-back values; GD&T symbols looked up in your gtol.sym
'   - no variable named like a procedure or a VBA keyword (vba_static_check.py)
'
'  v1: first version
'
'  SAFETY
'   - The part file is never saved or changed by this macro.
' =====================================================================================
Option Explicit

Dim swApp As Object, swMath As Object, swPart As Object, swModel As Object, swDraw As Object
Dim partPath As String, curStep As String, report As String, problems As String, todoList As String
Dim nOK As Long, nBad As Long, nTodo As Long
Dim lastDim As Object
Dim oldInputDim As Boolean, savedInputDim As Boolean
Dim skCand As Integer
Dim eKey As String

Const IN_M = 0.0254
Const DRAW_NAME = "HW#8D"
Const FILE_TAG = "HW8D"
Const SHEET_DEN = 1#             ' sheet scale 1 : 1 (every view, as the title block says)
Const EMAX = 12000               ' edge table size (the engraved logo adds many edges)
Const PLACE_TOL = 0.001
Const DEG = 0.0174532925199433
Const LOGO_TOL = 0.005           ' THREE PLACE DECIMAL tolerance in the title block

' ---- layout (sheet inches; measured on the reference: border 848 x 653 px, part 82.5 px/in) ----
Const MAIN_X = 1.408             ' main view: lower-left corner of the plate (x = 0, u = 0)
Const MAIN_Y = 1.978
Const BB_Y = 6.92                ' SECTION B-B: the back face (datum A)
Const AA_X = 10.06               ' SECTION A-A: the back face
Const ISO_X = 8.88               ' isometric: the plate's centre
Const ISO_Y = 6.15
Const NOTE_X = 0.45
Const NOTE_Y = 8.14
Const NOTE_DY = 0.167
Const NOTE_TAB = 0.41

Const DM_HIDDEN_VISIBLE = 1
Const DM_HIDDEN_REMOVED = 2
Const TOL_BASIC = 1              ' swTolType_e
Const TOL_LIMIT = 3
Const TOL_SYM = 4

' ---- the reference part (part inches: x from the left end, u up from the bottom edge (datum B),
'      z toward you from the back face (datum A)) ----
Const XA = 0.84375               ' cutting line A-A: the logo pocket's centre (.250 + 1.1875 / 2)
Const UB = 1.25                  ' cutting line B-B: the row of dia .257 holes (1.250 BASIC)
Const Z_TOP = 0.48               ' platform face
Const Z_RIM = 0.28               ' rim face (.480 - .200)
Const Z_FLOOR = 0.355            ' logo pocket floor (.125 deep)

' ---- measured in FindOrientation ----
Dim vR As Variant                ' along the plate, pocket end -> notch end   (right in the main view)
Dim vU As Variant                ' from the bottom edge (datum B) up          (up in the main view)
Dim vF As Variant                ' from the back face toward the platform     (toward you)
Dim pOrg As Variant              ' x = 0, u = 0, z = 0 (model, metres)
Dim lenX As Double, lenU As Double, thk As Double, zRimM As Double, zFloorM As Double
Dim xGroove As Double, zGroove As Double, gotGroove As Boolean
Dim hX(4) As Double, hU(4) As Double  ' 0 blind dia .257, 1 counterbore, 2 countersink, 3 1/4-20, 4 dia .201
Dim hOK(4) As Boolean

Dim viewMain As Object, viewBB As Object, viewAA As Object, viewIso As Object

Dim eN As Long
Dim eObj() As Object, eKind() As String
Dim eX1() As Double, eY1() As Double, eX2() As Double, eY2() As Double
Dim eCX() As Double, eCY() As Double, eR() As Double, eFull() As Boolean
Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double

Sub main()
    On Error GoTo Fail
    Dim tmpl As String, errNo As Long, errText As String
    Set swApp = Application.SldWorks
    Set swMath = swApp.GetMathUtility
    report = "": problems = "": todoList = "": nOK = 0: nBad = 0: nTodo = 0
    Set viewMain = Nothing: Set viewBB = Nothing: Set viewAA = Nothing: Set viewIso = Nothing
    Set lastDim = Nothing: eKey = "": skCand = -1: Set swModel = Nothing
    NoDimDialog True
    SpeedMode True

    Stage "Step 1 - find the open part":                  FindPart
    Stage "Step 2 - measure the part and the logo":       FindOrientation
    Stage "Step 3 - find your template":                  tmpl = FindTemplate()
    Stage "Step 4 - new drawing":                         NewDrawing tmpl
    SpeedMode True
    Stage "Step 5 - main view":                           MakeMain
    Stage "Step 6 - isometric view":                      MakeIso
    Stage "Step 7 - main view dimensions and GD&T":       ClearCommand: DimMain
    Stage "Step 8 - balloon 3 on the logo":              ClearCommand: Balloon3
    Stage "Step 9 - centre marks":                        ClearCommand: CentreMarks
    Stage "Step 10 - SECTION A-A and SECTION B-B":       ClearCommand: MakeSections
    Stage "Step 11 - SECTION B-B dimensions and GD&T":    ClearCommand: DimBB
    Stage "Step 12 - SECTION A-A dimensions":             ClearCommand: DimAA
    Stage "Step 13 - notes":                              ClearCommand: AddNotes
    Stage "Step 14 - title block":                        ClearCommand: TitleBlock
    Stage "Step 15 - section labels":                     ClearCommand: HideSectionLabels
    Stage "Step 16 - scales and alignment":               CheckScalesAndAlignment
    Stage "Step 17 - centre lines":                       ClearCommand: CentreLines
    Stage "Step 18 - finish":                             FinishSheet
    NoDimDialog False
    SpeedMode False

    MsgBox DRAW_NAME & " drawing built." & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & "      to do by hand: " & nTodo & vbCrLf & _
           IIf(nBad = 0, "", vbCrLf & "PROBLEMS:" & vbCrLf & Left(problems, 450)) & _
           IIf(nTodo = 0, "", vbCrLf & "TO DO BY HAND:" & vbCrLf & Left(todoList, 500)) & vbCrLf & _
           "Full list: " & FILE_TAG & "_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), DRAW_NAME

    Stage "Step 19 - save drawing and PDF":              SaveAll
    Exit Sub
Fail:
    errNo = Err.Number: errText = Err.Description       ' before the clean-up below clears it
    NoDimDialog False
    SpeedMode False
    MsgBox "The macro stopped at: " & curStep & vbCrLf & vbCrLf & _
           "Error " & errNo & ": " & errText & vbCrLf & vbCrLf & _
           "Checks passed so far: " & nOK & vbCrLf & _
           IIf(nBad = 0, "", "Problems so far:" & vbCrLf & Left(problems, 500)), vbCritical, DRAW_NAME & " macro"
End Sub

Function PartMatches(ByVal t As String) As Boolean
    t = UCase(Replace(Replace(t, "_", " "), "#", ""))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    PartMatches = (InStr(t, "DRAWINGS II") > 0 And InStr(t, "PART 4") > 0)
End Function

' part inches (x along vR, u along vU, z along vF) of a model point (metres)
Function ToPart(p As Variant) As Variant
    Dim d As Variant
    d = Sub3(p, pOrg)
    ToPart = Array(Dot(d, vR) / IN_M, Dot(d, vU) / IN_M, Dot(d, vF) / IN_M)
End Function

' part point (inches) -> sheet point of view v
Function SP(v As Object, p As Variant) As Variant
    SP = SheetPt(v, CDbl(p(0)), CDbl(p(1)), CDbl(p(2)))
End Function

Function Len3(a As Variant) As Double
    Len3 = Sqr(a(0) * a(0) + a(1) * a(1) + a(2) * a(2))
End Function

' q moved along n into the plane through p0 square to n
Function Flat(q As Variant, p0 As Variant, n As Variant) As Variant
    Flat = Sub3(q, Scl(n, Dot(Sub3(q, p0), n)))
End Function

' add centre q to list c (k entries) unless one is already within .02 in
Sub AddCentre(c() As Variant, k As Long, q As Variant)
    Dim j As Long
    For j = 0 To k - 1
        If Len3(Sub3(c(j), q)) / IN_M < 0.02 Then Exit Sub
    Next
    If k <= UBound(c) Then c(k) = q: k = k + 1
End Sub

' The plate: three dia .257 holes in a row (blind, counterbore, countersink), the counterbore on the
' platform side, the R.50 notch at the countersink's end. Then every feature is checked.
Sub FindOrientation()
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, j As Long, k As Long
    Dim s As Object, cp As Variant, r As Double, ax As Variant, q As Variant, pp As Variant, nn As Variant, bx As Variant
    Dim n As Variant, p0 As Variant, gotN As Boolean, t As Double, z As Double, ar As Double, d As Double, dMax As Double
    Dim c257(7) As Variant, k257 As Long, c201(7) As Variant, k201 As Long
    Dim cbMid As Variant, gotCb As Boolean, cNotch As Variant, gotNotch As Boolean, gq As Variant
    Dim e1 As Long, e2 As Long, em As Long
    Dim xLo As Double, xHi As Double, uLo As Double, uHi As Double, zLo As Double, zHi As Double, tLo As Double, tHi As Double
    Dim aRim As Double, aFloor As Double, nEng As Long, nRaised As Long, dLo As Double, dHi As Double
    Dim want As Variant, lbl As Variant, msg As String, ok As Boolean

    bodies = swPart.GetBodies2(0, True)
    If Not IsArray(bodies) Then Err.Raise vbObjectError + 3, , "The part has no solid body."

    ' pass 1: the dia .257 holes fix the plate's normal
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder Then
                cp = s.CylinderParams
                If Abs(cp(6) / IN_M - 0.1305) < 0.004 Then              ' dia .253 - .269 (not the R.125 corners or groove)
                    ax = Unit(Arr3(cp(3), cp(4), cp(5)))
                    q = Arr3(cp(0), cp(1), cp(2))
                    If Not gotN Then n = ax: p0 = q: gotN = True
                    If Abs(Dot(ax, n)) > 0.999 Then AddCentre c257, k257, Flat(q, p0, n)
                End If
            End If
        Next
    Next
    If k257 <> 3 Then Err.Raise vbObjectError + 4, , "Found " & k257 & " dia .257 holes (the In-Lab #8D part has 3). Click into the Part 4 window and run again."

    tLo = 1E+30: tHi = -1E+30
    ' pass 2: dia .201 holes, the counterbore (dia .44), the notch (R.50), the groove (R.125 across the plate)
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder Then
                cp = s.CylinderParams
                r = cp(6) / IN_M
                ax = Unit(Arr3(cp(3), cp(4), cp(5)))
                q = Arr3(cp(0), cp(1), cp(2))
                If Abs(Dot(ax, n)) > 0.999 Then
                    If Abs(r - 0.1005) < 0.003 Then AddCentre c201, k201, Flat(q, p0, n)
                    If Abs(r - 0.22) < 0.003 And Not gotCb Then
                        bx = faces(i).GetBox
                        cbMid = Arr3((bx(0) + bx(3)) / 2, (bx(1) + bx(4)) / 2, (bx(2) + bx(5)) / 2): gotCb = True
                    End If
                    If Abs(r - 0.5) < 0.003 And Not gotNotch Then cNotch = Flat(q, p0, n): gotNotch = True
                ElseIf Abs(Dot(ax, n)) < 0.01 And Abs(r - 0.125) < 0.003 And Not gotGroove Then
                    gq = q: gotGroove = True
                End If
            ElseIf s.IsPlane Then
                pp = s.PlaneParams
                nn = Unit(Arr3(pp(0), pp(1), pp(2)))
                If Abs(Dot(nn, n)) > 0.9999 Then
                    t = Dot(Arr3(pp(3), pp(4), pp(5)), n)
                    If t < tLo Then tLo = t
                    If t > tHi Then tHi = t
                End If
            End If
        Next
    Next
    If Not gotCb Then Err.Raise vbObjectError + 5, , "No counterbore (dia .44) found - this is not the In-Lab #8D part."

    ' the hole row: the two holes furthest apart are its ends; the countersunk end is nearer the notch
    dMax = -1
    For j = 0 To 2
        For k = j + 1 To 2
            d = Len3(Sub3(c257(j), c257(k)))
            If d > dMax Then dMax = d: e1 = j: e2 = k
        Next
    Next
    em = 3 - e1 - e2
    If gotNotch Then
        If Len3(Sub3(c257(e1), cNotch)) < Len3(Sub3(c257(e2), cNotch)) Then k = e1: e1 = e2: e2 = k
    End If
    vR = Unit(Sub3(c257(e2), c257(e1)))
    If Dot(cbMid, n) > (tLo + tHi) / 2 Then vF = n Else vF = Scl(n, -1)      ' the counterbore is on the platform side
    vU = Cross(vF, vR)

    ' pass 3: the plate's extents (flat end faces, never the body box)
    xLo = 1E+30: xHi = -1E+30: uLo = 1E+30: uHi = -1E+30: zLo = 1E+30: zHi = -1E+30
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsPlane Then
                pp = s.PlaneParams
                nn = Unit(Arr3(pp(0), pp(1), pp(2)))
                q = Arr3(pp(3), pp(4), pp(5))
                If Abs(Dot(nn, vR)) > 0.9999 Then
                    t = Dot(q, vR): If t < xLo Then xLo = t
                    If t > xHi Then xHi = t
                ElseIf Abs(Dot(nn, vU)) > 0.9999 Then
                    t = Dot(q, vU): If t < uLo Then uLo = t
                    If t > uHi Then uHi = t
                ElseIf Abs(Dot(nn, vF)) > 0.9999 Then
                    t = Dot(q, vF): If t < zLo Then zLo = t
                    If t > zHi Then zHi = t
                End If
            End If
        Next
    Next
    pOrg = Add3(Add3(Scl(vR, xLo), Scl(vU, uLo)), Scl(vF, zLo))
    lenX = (xHi - xLo) / IN_M: lenU = (uHi - uLo) / IN_M: thk = (zHi - zLo) / IN_M

    ' pass 4: the rim, the pocket floor and the logo (flat faces square to vF)
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsPlane Then
                pp = s.PlaneParams
                nn = Unit(Arr3(pp(0), pp(1), pp(2)))
                If Abs(Dot(nn, vF)) > 0.9999 Then
                    z = (Dot(Arr3(pp(3), pp(4), pp(5)), vF) - zLo) / IN_M
                    ar = faces(i).GetArea
                    If Abs(z - Z_RIM) < 0.02 And ar > aRim Then aRim = ar: zRimM = z
                    If Abs(z - Z_FLOOR) < 0.02 And ar > aFloor Then aFloor = ar: zFloorM = z
                End If
            End If
        Next
    Next
    dLo = 1#: dHi = 0#
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsPlane Then
                pp = s.PlaneParams
                nn = Unit(Arr3(pp(0), pp(1), pp(2)))
                If Abs(Dot(nn, vF)) > 0.9999 Then
                    z = (Dot(Arr3(pp(3), pp(4), pp(5)), vF) - zLo) / IN_M
                    If z < zFloorM - 0.0003 And z > zFloorM - 0.02 Then
                        nEng = nEng + 1
                        If zFloorM - z < dLo Then dLo = zFloorM - z
                        If zFloorM - z > dHi Then dHi = zFloorM - z
                    ElseIf z > zFloorM + 0.0003 And z < zFloorM + 0.02 Then
                        nRaised = nRaised + 1
                    End If
                End If
            End If
        Next
    Next

    ' holes in part inches, named by the reference layout
    q = ToPart(c257(e1)): hX(0) = q(0): hU(0) = q(1)
    q = ToPart(c257(em)): hX(1) = q(0): hU(1) = q(1)
    q = ToPart(c257(e2)): hX(2) = q(0): hU(2) = q(1)
    hX(3) = 2.5: hU(3) = 0.5: hX(4) = 3.75: hU(4) = 0.25
    For k = 3 To 4
        dMax = 1E+30
        For j = 0 To k201 - 1
            q = ToPart(c201(j))
            d = Abs(q(0) - IIf(k = 3, 2.5, 3.75)) + Abs(q(1) - IIf(k = 3, 0.5, 0.25))
            If d < dMax Then dMax = d: hX(k) = q(0): hU(k) = q(1)
        Next
    Next
    If gotGroove Then
        q = ToPart(gq): xGroove = q(0): zGroove = q(2)
    End If

    ' ---- checks against the reference ----
    Check (Abs(lenX - 4#) < 0.005 And Abs(lenU - 1.875) < 0.005 And Abs(thk - 0.48) < 0.003), _
          "Plate 4.000 x 1.875 x .480 (measured " & Format(lenX, "0.000") & " x " & Format(lenU, "0.000") & " x " & Format(thk, "0.000") & ")"
    want = Array(Array(2.125, 1.25), Array(2.75, 1.25), Array(3.375, 1.25), Array(2.5, 0.5), Array(3.75, 0.25))
    lbl = Array(ChrW(216) & ".257 deep .38", ChrW(216) & ".257 counterbore", ChrW(216) & ".257 countersink", "1/4-20 tapped", ChrW(216) & ".201")
    msg = ""
    For k = 0 To 4
        hOK(k) = (Abs(hX(k) - want(k)(0)) < 0.005 And Abs(hU(k) - want(k)(1)) < 0.005)
        If k >= 3 And k201 = 0 Then hOK(k) = False
        If Not hOK(k) Then msg = msg & " " & lbl(k) & " at " & Format(hX(k), "0.000") & ", " & Format(hU(k), "0.000") & ";"
    Next
    Check (Len(msg) = 0), "Holes where the reference has them (2.125 / 2.750 / 3.375 at 1.250; 1/4-20 at 2.500, .500; " & ChrW(216) & ".201 at 3.750, .250)" & _
          IIf(Len(msg) = 0, "", " - DIFFERENT:" & msg)
    Check (Abs(thk - zRimM - 0.2) < 0.003), "Rim step .200 (measured " & Format(thk - zRimM, "0.000") & ")"
    Check (Abs(thk - zFloorM - 0.125) < 0.003), "Logo pocket .125 deep (measured " & Format(thk - zFloorM, "0.000") & ")"
    If gotGroove Then
        Check (Abs(xGroove - 1.6875) < 0.003), "Groove R.125 centre 1.563 from the platform edge (measured " & Format(xGroove - 0.125, "0.0000") & ")"
    Else
        Check False, "Groove (R.125 across the plate) not found - the 1.563 dimension will be missing"
    End If

    ' note 3: the logo engraved .003 deep in the pocket floor
    If nEng = 0 Then
        Check False, "Note 3: no engraved logo found below the pocket floor" & IIf(nRaised > 0, " (" & nRaised & " raised faces found instead)", "") & _
                     " - the part needs the logo cut .003 deep (Extruded Cut, Blind .003)"
    Else
        ok = (dLo > 0.003 - LOGO_TOL - 0.00005 And dHi < 0.003 + LOGO_TOL + 0.00005)
        Check ok, "Note 3: logo engraved " & Format(dHi, ".0000") & IIf(Abs(dHi - dLo) > 0.00005, " (shallowest " & Format(dLo, ".0000") & ")", "") & _
                  " deep in the pocket floor (" & nEng & " faces); note 3 asks .003 " & ChrW(177) & Mid(Format(LOGO_TOL, "0.000"), 2) & _
                  " (THREE PLACE tolerance)" & IIf(ok, "", " - edit the logo cut's depth to .003 in the part")
        If nRaised > 0 Then Info "  also " & nRaised & " logo faces sit above the pocket floor"
    End If
    Info "Main view frame: right " & VecTxt(vR) & ", up " & VecTxt(vU) & ", toward you " & VecTxt(vF)
End Sub

' =====================================================================================
'  Run control, part, template  (from HW#8C)
' =====================================================================================
Sub NoDimDialog(ByVal turnOff As Boolean)
    On Error Resume Next
    If turnOff Then
        Err.Clear
        oldInputDim = swApp.GetUserPreferenceToggle(10)       ' 10 = swInputDimValOnCreate
        savedInputDim = (Err.Number = 0)
        If savedInputDim Then swApp.SetUserPreferenceToggle 10, False
    ElseIf savedInputDim Then
        swApp.SetUserPreferenceToggle 10, oldInputDim
        savedInputDim = False
    End If
End Sub

' v7: tell SolidWorks a long macro is running (fewer screen and tree updates); always switched back
Sub SpeedMode(ByVal turnOn As Boolean)
    On Error Resume Next
    swApp.CommandInProgress = turnOn
    If Not swModel Is Nothing Then swModel.FeatureManager.EnableFeatureTree = Not turnOn
End Sub

' a step the macro cannot do (SolidWorks gives macros no access to it): listed for you, not a problem
Sub ToDo(ByVal label As String)
    nTodo = nTodo + 1
    report = report & "TO DO    " & label & vbCrLf
    todoList = todoList & "- " & label & vbCrLf
End Sub

' information only (not counted)
Sub Info(ByVal label As String)
    report = report & "INFO     " & label & vbCrLf
End Sub

' show the step in SolidWorks' status bar
Sub Stage(ByVal s As String)
    curStep = s
    On Error Resume Next
    swApp.Frame.SetStatusBarText DRAW_NAME & " macro: " & s & "   (please wait)"
    DoEvents
End Sub

Sub FindPart()
    ' Use the open part named "...Drawings II In-Lab Assignment Part 4". Otherwise the active part,
    ' otherwise the only open part.  Step 2 then checks its shape.
    Dim d As Object, docs As Variant, i As Long, nParts As Long, firstPart As Object
    Set swPart = Nothing
    On Error Resume Next
    docs = swApp.GetDocuments
    On Error GoTo 0
    If IsArray(docs) Then
        For i = LBound(docs) To UBound(docs)
            If docs(i).GetType = 1 Then
                nParts = nParts + 1
                If firstPart Is Nothing Then Set firstPart = docs(i)
                If PartMatches(docs(i).GetTitle & " " & docs(i).GetPathName) Then Set swPart = docs(i)
            End If
        Next
    End If
    If swPart Is Nothing Then
        Set d = swApp.ActiveDoc
        If Not d Is Nothing Then
            If d.GetType = 1 Then Set swPart = d
        End If
    End If
    If swPart Is Nothing And nParts = 1 Then Set swPart = firstPart
    If swPart Is Nothing Then Err.Raise vbObjectError + 1, , _
        "Could not tell which open part is In-Lab Assignment Part 4. Close the other parts, click into its window, and run again."
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then Err.Raise vbObjectError + 2, , "The part has never been saved. Save it first."
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

' part coordinates (inches): x along the axis from the left end, r up from the axis, z toward you
Function PartPt(ByVal x As Double, ByVal r As Double, ByVal z As Double) As Variant
    PartPt = Add3(pOrg, Add3(Add3(Scl(vR, x * IN_M), Scl(vU, r * IN_M)), Scl(vF, z * IN_M)))
End Function

Function SheetPt(v As Object, ByVal x As Double, ByVal r As Double, ByVal z As Double) As Variant
    Dim p As Variant
    p = PartPt(x, r, z)
    SheetPt = XfPt(v.ModelToViewTransform, p(0), p(1), p(2))
End Function

Function SX(v As Object, ByVal x As Double) As Double
    SX = SheetPt(v, x, 0, 0)(0)
End Function

Function SY(v As Object, ByVal r As Double) As Double
    SY = SheetPt(v, 0, r, 0)(1)
End Function

Function FindTemplate() As String
    Dim cands As Variant, c As Variant, p As String, up As String, od As String, odc As String
    up = Environ("USERPROFILE"): od = Environ("OneDrive"): odc = Environ("OneDriveCommercial")
    cands = Array(up & "\OneDrive - Cal Poly\Documents\Hayden_p_ansi.drwdot", _
                  odc & "\Documents\Hayden_p_ansi.drwdot", od & "\Documents\Hayden_p_ansi.drwdot", _
                  up & "\Documents\Hayden_p_ansi.drwdot", up & "\Downloads\Hayden_p_ansi.drwdot", _
                  up & "\Desktop\Hayden_p_ansi.drwdot", odc & "\Desktop\Hayden_p_ansi.drwdot", _
                  od & "\Desktop\Hayden_p_ansi.drwdot", up & "\OneDrive - Cal Poly\Desktop\Hayden_p_ansi.drwdot", _
                  up & "\OneDrive - Cal Poly\Downloads\Hayden_p_ansi.drwdot")
    For Each c In cands
        p = CStr(c)
        If FileExists(p) Then FindTemplate = p: Check True, "Template: " & p: Exit Function
    Next
    p = InputBox("Could not find Hayden_p_ansi.drwdot automatically." & vbCrLf & vbCrLf & _
                 "In File Explorer, Shift + right-click the template and choose 'Copy as path'," & vbCrLf & _
                 "then paste it here:", DRAW_NAME & " - drawing template")
    p = Replace(Trim(p), """", "")
    If Not FileExists(p) Then Err.Raise vbObjectError + 6, , "Template not found: " & p
    FindTemplate = p
    Check True, "Template: " & p
End Function

Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    If Len(p) < 5 Then Exit Function
    FileExists = (Len(Dir(p)) > 0)
End Function

Sub NewDrawing(ByVal tmpl As String)
    Dim sh As Object, p As Variant
    Set swModel = Nothing
    On Error Resume Next
    Set swModel = swApp.NewDocument(tmpl, 12, 11# * IN_M, 8.5 * IN_M)   ' 12 = user size (A, 11 x 8.5)
    If swModel Is Nothing Then Set swModel = swApp.NewDocument(tmpl, 0, 0#, 0#)
    On Error GoTo 0
    If swModel Is Nothing Then Err.Raise vbObjectError + 7, , "SolidWorks could not create a drawing from " & tmpl
    If swModel.GetType <> 3 Then Err.Raise vbObjectError + 8, , "The template did not create a drawing."
    Set swDraw = swModel
    On Error Resume Next
    swDraw.EditSheet
    On Error GoTo 0

    Set sh = swDraw.GetCurrentSheet
    On Error Resume Next
    p = sh.GetProperties2                 ' paper size, template, scale1, scale2, first angle, width, height, same props
    If IsArray(p) Then
        sh.SetProperties2 CLng(p(0)), CLng(p(1)), 1#, SHEET_DEN, False, CDbl(p(5)), CDbl(p(6)), True
    End If
    p = sh.GetProperties2
    On Error GoTo 0
    If IsArray(p) Then
        Check (Abs(p(2) - 1) < 0.001 And Abs(p(3) - SHEET_DEN) < 0.001), "Sheet scale 1:1"
        Check (p(4) = 0 Or p(4) = False), "Sheet projection: third angle"
    End If
End Sub

' =====================================================================================
'  Views
' =====================================================================================
Sub MakeMain()
    Dim ok As Boolean
    Set viewMain = OrientedView("HW8D_Main", vR, vU, vF, MAIN_X + 2#, MAIN_Y + 0.94, StdFrontName(), False)
    If viewMain Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented main view."
    Check True, "Main view: pocket and holes toward you, logo pocket on the left, datum B edge at the bottom"
    EnsureSheetScale                                   ' auto-scale can change the sheet when the first view goes in
    ok = SetViewScale(viewMain, 1#, SHEET_DEN)
    Check ok, "Main view scale 1:1 (reads " & ScaleTxt(viewMain) & ")"
    SetMode viewMain, DM_HIDDEN_VISIBLE, "Main"
    PlaceView viewMain, 0#, 0#, 0#, MAIN_X, MAIN_Y, True, True, "Main view placed (plate's lower-left corner)"
End Sub

Sub MakeIso()
    On Error GoTo Oops
    Dim dv As Variant, upI As Variant, rtI As Variant, ok As Boolean
    dv = Unit(Add3(Add3(vR, vU), vF))
    upI = Unit(Sub3(vU, Scl(dv, Dot(vU, dv))))
    rtI = Cross(upI, dv)
    Set viewIso = OrientedView("HW8D_Iso", rtI, upI, dv, ISO_X, ISO_Y, StdIsoName(), True)
    If viewIso Is Nothing Then
        Check False, "Isometric view - could not create it; add it by hand (View Palette > Isometric, scale 1:1)"
        Exit Sub
    End If
    Check True, "Isometric view (logo face toward you, pocket end top left)"
    ok = SetViewScale(viewIso, 1#, 1#)
    Check ok, "Isometric view scale 1:1 (reads " & ScaleTxt(viewIso) & ")"
    SetMode viewIso, DM_HIDDEN_REMOVED, "Isometric"
    PlaceView viewIso, lenX / 2, lenU / 2, thk / 2, ISO_X, ISO_Y, True, True, "Isometric view placed (plate centre)"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' Cutting lines A-A (vertical, through the logo pocket) and B-B (horizontal, through the holes), each
' proved in place by a test dimension; then the two sections, turned the reference's way and placed.
' (Made after every main-view dimension: lesson 18, sketch items last.)
Sub MakeSections()
    On Error GoTo Oops
    Dim segA As Object, segB As Object, ok As Boolean
    Set segA = CutLine(Array(XA, -0.15, Z_TOP), Array(XA, lenU + 0.155, Z_TOP), True, XA, "A-A")
    If Not segA Is Nothing Then
        Set viewAA = MakeSec(segA, "A", AA_X - 0.24, MAIN_Y + 0.94)
        If viewAA Is Nothing Then
            Check False, "SECTION A-A - SolidWorks would not make it from the cutting line; select the line, Drawing > Section View, place it right of the main view"
        Else
            ' looking left (arrows point left): platform face on the left, up stays up
            FixSecDir viewAA, vF, -1, 0, vU, 0, 1, "SECTION A-A looks left (arrows point left): platform face on its left, as on the reference"
            SecLabel viewAA, "A"
            Check SetViewScale(viewAA, 1#, 1#), "SECTION A-A scale 1:1 (reads " & ScaleTxt(viewAA) & ")"
            SetMode viewAA, DM_HIDDEN_REMOVED, "SECTION A-A"
            PlaceView viewAA, XA, 0#, 0#, AA_X, 0#, True, False, "SECTION A-A placed (back face at x " & Format(AA_X, "0.00") & ")"
        End If
    End If
    Set segB = CutLine(Array(-0.28, UB, Z_TOP), Array(lenX + 0.16, UB, Z_TOP), False, UB, "B-B")
    If Not segB Is Nothing Then
        Set viewBB = MakeSec(segB, "B", MAIN_X + 2#, BB_Y - 0.24)
        If viewBB Is Nothing Then
            Check False, "SECTION B-B - SolidWorks would not make it from the cutting line; select the line, Drawing > Section View, place it above the main view"
        Else
            ' looking down (arrows point down): back face (datum A) at the top
            FixSecDir viewBB, vR, 1, 0, vF, 0, -1, "SECTION B-B looks down (arrows point down): back face on top, as on the reference"
            SecLabel viewBB, "B"
            Check SetViewScale(viewBB, 1#, 1#), "SECTION B-B scale 1:1 (reads " & ScaleTxt(viewBB) & ")"
            SetMode viewBB, DM_HIDDEN_REMOVED, "SECTION B-B"
            PlaceView viewBB, 0#, UB, 0#, 0#, BB_Y, False, True, "SECTION B-B placed (back face at y " & Format(BB_Y, "0.00") & ")"
        End If
    End If
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' a cutting line in the main view from part point a to b, kept only when a test dimension measures
' it "want" from the left end (vertical line) or the bottom edge (horizontal line)
Function CutLine(a As Variant, b As Variant, ByVal vertLine As Boolean, ByVal want As Double, ByVal label As String) As Object
    Dim cand As Variant, seg As Object, d As Double, tried As String, order As Variant
    If skCand >= 0 Then order = Array(skCand) Else order = Array(2, 3, 1, 0)
    For Each cand In order
        Set seg = SkLineP(viewMain, a, b, CInt(cand))
        If seg Is Nothing Then
            tried = tried & "m" & cand & ": no line. "
        Else
            If vertLine Then
                d = TestH(viewMain, seg, Array(0#, 0.9375, Z_TOP))
            Else
                d = TestV(viewMain, seg, Array(3#, 0#, Z_TOP))
            End If
            If Abs(d - want) < 0.003 Then
                skCand = CInt(cand)
                Check True, "Cutting line " & label & " in place (sketch method " & cand & ", measured " & Format(d, "0.0000") & " from the " & _
                            IIf(vertLine, "left end", "bottom edge") & ")"
                Set CutLine = seg
                Exit Function
            End If
            tried = tried & "m" & cand & ": " & Format(d, "0.000") & ". "
            DeleteSeg seg
        End If
    Next
    If skCand >= 0 And UBound(order) = 0 Then                   ' the proven method failed this time: try them all
        skCand = -1
        Set CutLine = CutLine(a, b, vertLine, want, label)
        Exit Function
    End If
    Check False, "Cutting line " & label & " - no sketch method put it in place (" & Trim(tried) & "); draw it with Drawing > Section View"
End Function

Function MakeSec(seg As Object, ByVal lbl As String, ByVal px As Double, ByVal py As Double) As Object
    Dim opt As Integer, sv As Object, before As String
    For opt = 0 To 1
        before = ViewNames()
        swModel.ClearSelection2 True
        swDraw.ActivateView viewMain.GetName2
        On Error Resume Next
        seg.Select4 False, Nothing                                 ' the line itself (no screen picking)
        Set sv = Nothing
        Set sv = swDraw.CreateSectionViewAt5(px * IN_M, py * IN_M, 0#, lbl, opt, Nothing, 0#)
        If sv Is Nothing Then Err.Clear: Set sv = swDraw.CreateSectionViewAt4(px * IN_M, py * IN_M, 0#, lbl, opt, Nothing)
        Err.Clear
        On Error GoTo 0
        If sv Is Nothing Then Set sv = NewViewSince(before)        ' built but not handed back
        If Not sv Is Nothing Then Exit For
    Next
    swModel.ClearSelection2 True
    If Not sv Is Nothing Then swModel.EditRebuild3: eKey = ""
    Set MakeSec = sv
End Function

' turn a section round if it looks the wrong way; judged on the view's own axes
Sub FixSecDir(sv As Object, va As Variant, ByVal ax As Double, ByVal ay As Double, vb As Variant, ByVal bx As Double, ByVal by As Double, ByVal label As String)
    Dim sec As Object, ok As Boolean
    If Not FrameOK(sv, va, ax, ay, vb, bx, by) Then
        On Error Resume Next
        Set sec = sv.GetSection
        sec.SetReversedCutDirection Not sec.GetReversedCutDirection
        swModel.ForceRebuild3 False
        eKey = ""
        On Error GoTo 0
    End If
    ok = FrameOK(sv, va, ax, ay, vb, bx, by)
    Check ok, label & IIf(ok, "", " - click its cutting line and tick Flip Direction")
End Sub

Sub SecLabel(sv As Object, ByVal lbl As String)
    Dim sec As Object, t As String
    On Error Resume Next
    Set sec = sv.GetSection
    t = ""
    t = sec.GetLabel
    If UCase(Trim(t)) <> lbl Then sec.SetLabel2 lbl: t = sec.GetLabel
    On Error GoTo 0
    If Len(t) = 0 Then
        Info "SECTION " & lbl & "-" & lbl & ": SolidWorks does not report the letter - look that the arrows read " & lbl
    ElseIf UCase(Trim(t)) = lbl Then
        Check True, "SECTION " & lbl & "-" & lbl & " letter reads " & t
    Else
        ToDo "SECTION " & lbl & "-" & lbl & ": its arrows read " & t & " - click the cutting line and set the label to " & lbl
    End If
End Sub

' vertical test dimension from sketch item it to the edge through part point pEdge (deleted again)
Function TestV(v As Object, it As Object, pEdge As Variant) As Double
    Dim dd As Object, n As Long, tx As Double, ty As Double
    TestV = -1
    tx = SX(v, -0.6): ty = SY(v, 0.7)
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    On Error Resume Next
    it.Select4 False, Nothing
    n = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    On Error GoTo 0
    If n <> 1 Then swModel.ClearSelection2 True: Exit Function
    If Not PickOnly(v, pEdge, True) Then
        If Not SelPt(v, pEdge, True) Then swModel.ClearSelection2 True: Exit Function
    End If
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) <> 2 Then swModel.ClearSelection2 True: Exit Function
    Set dd = Nothing
    On Error Resume Next
    Set dd = swModel.AddVerticalDimension2(tx * IN_M, ty * IN_M, 0#)
    On Error GoTo 0
    swModel.ClearSelection2 True
    If dd Is Nothing Then Exit Function
    On Error Resume Next
    TestV = dd.GetDimension2(0).SystemValue / IN_M
    On Error GoTo 0
    If Not DeleteDim(dd) Then Check False, "A test dimension (" & Format(TestV, "0.000") & ") could not be deleted - click it and press Delete"
End Function

' the reference shows no "SECTION A-A" / "SECTION B-B" text under the sections: hide those labels
Sub HideSectionLabels()
    Dim v As Object, anns As Variant, i As Long, t As String, nFound As Long, nHidden As Long, orig As Long, vis As Long, w As Variant
    On Error Resume Next
    Set v = swDraw.GetFirstView
    Do While Not v Is Nothing
        anns = Empty
        anns = v.GetAnnotations
        If IsArray(anns) Then
            For i = LBound(anns) To UBound(anns)
                t = ""
                t = UCase(anns(i).GetSpecificAnnotation.GetText)
                If Len(t) = 0 Then t = UCase(anns(i).GetTextAtIndex(0))
                If InStr(t, "SECTION") > 0 Then
                    nFound = nFound + 1
                    orig = -1: orig = anns(i).Visible
                    For Each w In Array(2, 1, 0)                    ' swAnnotationHidden, then the other states
                        If CLng(w) <> orig Then
                            anns(i).Visible = CLng(w)
                            vis = -1: vis = anns(i).Visible
                            If vis <> orig And vis = CLng(w) Then Exit For
                        End If
                    Next
                    If vis <> orig And vis >= 0 Then nHidden = nHidden + 1
                End If
            Next
        End If
        Set v = v.GetNextView
    Loop
    swModel.GraphicsRedraw2
    On Error GoTo 0
    If nFound = 0 Then
        Check True, "Section labels: none shown (as on the reference)"
    ElseIf nHidden = nFound Then
        Info "Section labels: " & nHidden & " of " & nFound & " switched to hidden - if you still see SECTION A-A / B-B text, right-click it > Hide"
    Else
        ToDo "Right-click the SECTION A-A and SECTION B-B labels under the sections > Hide (the reference shows none)"
    End If
End Sub

' centre marks on the five holes of the main view (one per hole)
Sub CentreMarks()
    On Error GoTo Oops
    Dim i As Long, k As Long, c As Variant, cm As Object, n As Long, want As Long, done(4) As Boolean, rr As Double
    eKey = ""
    CollectEdges viewMain
    For k = 0 To 4
        If hOK(k) Then
            want = want + 1
            c = SP(viewMain, Array(hX(k), hU(k), Z_TOP))
            rr = IIf(k <= 2, 0.1285, 0.1005)
            For i = 0 To eN - 1
                If Not done(k) And eKind(i) = "C" Then
                    If eFull(i) And Sqr((eCX(i) - c(0)) ^ 2 + (eCY(i) - c(1)) ^ 2) < 0.005 And Abs(eR(i) / ViewSc(viewMain) - rr) < 0.004 Then
                        swModel.ClearSelection2 True
                        swDraw.ActivateView viewMain.GetName2
                        SelEdge viewMain, i, False
                        Set cm = Nothing
                        On Error Resume Next
                        Set cm = swDraw.InsertCenterMark3(0, False, True)
                        On Error GoTo Oops
                        If Not cm Is Nothing Then done(k) = True: n = n + 1
                    End If
                End If
            Next
        End If
    Next
    swModel.ClearSelection2 True
    Check (n = want And n > 0), "Main     centre marks on " & n & " of " & want & " holes"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' centre lines through the holes and the groove in SECTION B-B (last: the command can stay active)
Sub CentreLines()
    Dim n0 As Long, n1 As Long
    If viewBB Is Nothing Then Exit Sub
    On Error Resume Next
    n0 = CountCentreLines(viewBB)
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 viewBB.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    swDraw.InsertCenterLine2
    swModel.ClearSelection2 True
    swDraw.ActivateSheet swDraw.GetCurrentSheet.GetName
    n1 = CountCentreLines(viewBB)
    On Error GoTo 0
    If n1 > n0 Then
        Check True, "SECTION B-B centre lines (" & n1 & ") through the holes and the groove"
    Else
        Info "SECTION B-B centre lines: SolidWorks did not report them - if missing, Annotation > Centerline, click the section view"
    End If
End Sub

Function CountCentreLines(v As Object) As Long
    Dim c As Object, n As Long
    On Error Resume Next
    Set c = v.GetFirstCenterLine
    Do While Not c Is Nothing
        n = n + 1
        Set c = c.GetNext
    Loop
    CountCentreLines = n
End Function

' every view at its scale and where it was measured; the sections in line with the main view
Sub CheckScalesAndAlignment()
    On Error GoTo Oops
    Dim a As Variant, b As Variant, ok As Boolean
    EnsureSheetScale
    a = SheetPt(viewMain, 0, 0, 0)
    ok = (Abs(a(0) - MAIN_X) < 0.002 And Abs(a(1) - MAIN_Y) < 0.002)
    Check ok, "Alignment: main view corner at " & Format(a(0), "0.000") & ", " & Format(a(1), "0.000")
    If Not viewBB Is Nothing Then
        b = SheetPt(viewBB, 0, UB, 0)
        Check (Abs(b(0) - a(0)) < 0.001), "Alignment: SECTION B-B straight above the main view (left ends at x " & Format(a(0), "0.000") & " / " & Format(b(0), "0.000") & ")"
    End If
    If Not viewAA Is Nothing Then
        b = SheetPt(viewAA, XA, 0, 0)
        Check (Abs(b(1) - a(1)) < 0.001), "Alignment: SECTION A-A level with the main view (bottom edges at y " & Format(a(1), "0.000") & " / " & Format(b(1), "0.000") & ")"
    End If
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.ClearSelection2 True
End Sub

' =====================================================================================
'  View helpers  (from HW#8C)
' =====================================================================================
' sheet scale 1:1, read back (set again if a view changed it)
Sub EnsureSheetScale()
    Dim sh As Object, p As Variant, k As Integer, ok As Boolean
    On Error Resume Next
    Set sh = swDraw.GetCurrentSheet
    For k = 1 To 2
        p = Empty
        p = sh.GetProperties2
        If IsArray(p) Then
            ok = (Abs(p(2) - 1#) < 0.001 And Abs(p(3) - SHEET_DEN) < 0.001)
            If ok Then Exit For
            sh.SetProperties2 CLng(p(0)), CLng(p(1)), 1#, SHEET_DEN, False, CDbl(p(5)), CDbl(p(6)), True
            swModel.EditRebuild3
            eKey = ""
        End If
    Next
    p = Empty
    p = sh.GetProperties2
    On Error GoTo 0
    If IsArray(p) Then
        ok = (Abs(p(2) - 1#) < 0.001 And Abs(p(3) - SHEET_DEN) < 0.001)
        Check ok, "Sheet scale reads 1:1 after the first view" & IIf(ok, "", " (" & p(2) & ":" & p(3) & ") - right-click the sheet > Properties > Scale 1:1")
    End If
End Sub

' set view v to scale num:den and read it back
Function SetViewScale(v As Object, ByVal num As Double, ByVal den As Double) As Boolean
    Dim sr(1) As Double, va As Variant, want As Double
    want = num / den
    If Abs(ViewSc(v) - want) > 0.0005 Then
        sr(0) = num: sr(1) = den: va = sr
        On Error Resume Next
        v.UseSheetScale = False
        v.ScaleRatio = va
        If Abs(ViewSc(v) - want) > 0.0005 Then Err.Clear: v.ScaleDecimal = want
        swModel.EditRebuild3
        eKey = ""
        On Error GoTo 0
    End If
    SetViewScale = (Abs(ViewSc(v) - want) < 0.0005)
End Function

Function ViewSc(v As Object) As Double
    Dim sc As Double
    On Error Resume Next
    sc = 0
    sc = v.ScaleDecimal
    If sc <= 0 Then sc = 1#
    ViewSc = sc
End Function

Function ScaleTxt(v As Object) As String
    Dim sc As Double
    sc = ViewSc(v)
    If sc >= 1 Then ScaleTxt = Format(sc, "0.##") & ":1" Else ScaleTxt = "1:" & Format(1# / sc, "0.##")
End Function

Function StdIsoName() As String
    If Near(vR, 1, 0, 0) And Near(vU, 0, 1, 0) And Near(vF, 0, 0, 1) Then StdIsoName = "*Isometric"
End Function

Sub DeleteSeg(seg As Object)
    On Error Resume Next
    If seg Is Nothing Then Exit Sub
    swModel.ClearSelection2 True
    seg.Select4 False, Nothing
    swModel.EditDelete
    swModel.ClearSelection2 True
End Sub

Function ViewNames() As String
    Dim v As Object, t As String
    On Error Resume Next
    Set v = swDraw.GetFirstView
    Do While Not v Is Nothing
        t = t & "|" & v.GetName2 & "|"
        Set v = v.GetNextView
    Loop
    ViewNames = t
End Function

Function NewViewSince(ByVal before As String) As Object
    Dim v As Object
    On Error Resume Next
    Set v = swDraw.GetFirstView
    Do While Not v Is Nothing
        If InStr(before, "|" & v.GetName2 & "|") = 0 Then Set NewViewSince = v: Exit Function
        Set v = v.GetNextView
    Loop
End Function

' move view v so that part point (r,u,f) lands at sheet (tx,ty); doX / doY choose which axes move
Sub PlaceView(v As Object, ByVal r As Double, ByVal u As Double, ByVal f As Double, ByVal tx As Double, ByVal ty As Double, _
              ByVal doX As Boolean, ByVal doY As Boolean, ByVal label As String)
    Dim cur As Variant, pos As Variant, k As Integer, ok As Boolean
    For k = 1 To 5
        cur = SheetPt(v, r, u, f)
        ok = (Not doX Or Abs(cur(0) - tx) < PLACE_TOL) And (Not doY Or Abs(cur(1) - ty) < PLACE_TOL)
        If ok Then Exit For
        On Error Resume Next
        pos = v.Position
        If doX Then pos(0) = pos(0) + (tx - cur(0)) * IN_M
        If doY Then pos(1) = pos(1) + (ty - cur(1)) * IN_M
        v.Position = pos
        swModel.EditRebuild3                          ' v7: partial rebuild; a full one only if the position reads stale
        If Abs(SheetPt(v, r, u, f)(0) - cur(0)) + Abs(SheetPt(v, r, u, f)(1) - cur(1)) < 0.000001 Then Call swModel.ForceRebuild3(False)
        On Error GoTo 0
    Next
    cur = SheetPt(v, r, u, f)
    ok = (Not doX Or Abs(cur(0) - tx) < PLACE_TOL) And (Not doY Or Abs(cur(1) - ty) < PLACE_TOL)
    Check ok, label & IIf(ok, "", " - could not move it there; drag it by hand")
End Sub

' create a model view, verify it, fall back to a custom named view in the part if needed
Function OrientedView(ByVal baseName As String, ax As Variant, ay As Variant, az As Variant, _
                      ByVal px As Double, ByVal py As Double, ByVal stdName As String, ByVal isIso As Boolean) As Object
    Dim v As Object, attempt As Integer, nm As String
    If Len(stdName) > 0 Then
        Set v = ModelViewAt(stdName, px, py)
        If Not v Is Nothing Then
            If ViewMatches(v, ax, ay) Then Set OrientedView = v: Exit Function
            DeleteView v
        End If
    End If
    For attempt = 1 To 2
        nm = baseName & attempt
        If MakeNamedView(nm, ax, ay, az, attempt) Then
            Set v = ModelViewAt(nm, px, py)
            If Not v Is Nothing Then
                If ViewMatches(v, ax, ay) Then Set OrientedView = v: Exit Function
                DeleteView v
            End If
        End If
    Next
    Set OrientedView = Nothing
End Function

Function ModelViewAt(ByVal nm As String, ByVal px As Double, ByVal py As Double) As Object
    On Error Resume Next
    Set ModelViewAt = swDraw.CreateDrawViewFromModelView3(partPath, nm, px * IN_M, py * IN_M, 0#)
End Function

' the view's screen-right must be ax and screen-up must be ay
Function ViewMatches(v As Object, ax As Variant, ay As Variant) As Boolean
    Dim a As Variant, b As Variant
    a = ViewDir(v, ax): b = ViewDir(v, ay)
    ViewMatches = (a(0) > 0.95 And Abs(a(1)) < 0.05 And Abs(b(0)) < 0.05 And b(1) > 0.95)
End Function

Function FrameOK(v As Object, va As Variant, ByVal ax As Double, ByVal ay As Double, _
                 vb As Variant, ByVal bx As Double, ByVal by As Double) As Boolean
    Dim a As Variant, b As Variant
    a = ViewDir(v, va): b = ViewDir(v, vb)
    FrameOK = (Abs(a(0) - ax) < 0.1 And Abs(a(1) - ay) < 0.1 And Abs(b(0) - bx) < 0.1 And Abs(b(1) - by) < 0.1)
End Function

Function StdFrontName() As String
    If Near(vU, 0, 1, 0) Then
        If Near(vF, 0, 0, 1) Then StdFrontName = "*Front"
        If Near(vF, 0, 0, -1) Then StdFrontName = "*Back"
        If Near(vF, 1, 0, 0) Then StdFrontName = "*Right"
        If Near(vF, -1, 0, 0) Then StdFrontName = "*Left"
    ElseIf Near(vF, 0, 1, 0) And Near(vU, 0, 0, -1) Then
        StdFrontName = "*Top"
    ElseIf Near(vF, 0, -1, 0) And Near(vU, 0, 0, 1) Then
        StdFrontName = "*Bottom"
    End If
End Function

' make a named view in the part with screen axes (ax right, ay up, az toward viewer)
Function MakeNamedView(ByVal nm As String, ax As Variant, ay As Variant, az As Variant, ByVal conv As Integer) As Boolean
    Dim e As Long, mv As Object, a(15) As Double, v As Variant, xf As Object, ok As Boolean
    On Error Resume Next
    swApp.ActivateDoc3 swPart.GetTitle, False, 0, e
    Set mv = swPart.ActiveView
    If conv = 1 Then
        a(0) = ax(0): a(1) = ax(1): a(2) = ax(2)
        a(3) = ay(0): a(4) = ay(1): a(5) = ay(2)
        a(6) = az(0): a(7) = az(1): a(8) = az(2)
    Else
        a(0) = ax(0): a(1) = ay(0): a(2) = az(0)
        a(3) = ax(1): a(4) = ay(1): a(5) = az(1)
        a(6) = ax(2): a(7) = ay(2): a(8) = az(2)
    End If
    a(12) = 1#
    v = a
    Set xf = swMath.CreateTransform((v))
    Err.Clear
    mv.Orientation3 = xf
    If Err.Number <> 0 Then Err.Clear: Set mv.Orientation3 = xf
    ok = (Err.Number = 0)
    swPart.GraphicsRedraw2
    swPart.DeleteNamedView nm
    Err.Clear
    swPart.NameView nm
    ok = ok And (Err.Number = 0)
    ActivateDrawing
    MakeNamedView = ok
End Function

Sub ActivateDrawing()
    Dim e As Long, t As String, d As Object
    On Error Resume Next
    t = swModel.GetTitle
    Set d = swApp.ActivateDoc3(t, False, 0, e)
    If d Is Nothing And InStr(t, " - ") > 0 Then Set d = swApp.ActivateDoc3(Left(t, InStr(t, " - ") - 1), False, 0, e)
End Sub

Sub DeleteView(v As Object)
    On Error Resume Next
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 v.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    swModel.EditDelete
    swModel.ClearSelection2 True
End Sub

Sub SetMode(v As Object, ByVal mode As Long, ByVal label As String)
    Dim m As Long
    On Error Resume Next
    v.SetDisplayMode3 False, mode, False, False
    m = -1
    m = v.GetDisplayMode2
    On Error GoTo 0
    If m <> -1 Then Check (m = mode), label & " view: " & IIf(mode = DM_HIDDEN_VISIBLE, "hidden lines visible", "hidden lines removed")
End Sub

' =====================================================================================
'  Dimensions and GD&T  (text positions measured on the reference, sheet inches from a part point)
' =====================================================================================
' main view: every real-edge dimension, the hole callouts with their position frames, datums B and C
Sub DimMain()
    On Error GoTo Oops
    Dim o As Variant, L As Variant, D As String, z As Double
    D = ChrW(216): z = Z_TOP
    eKey = ""
    CollectEdges viewMain
    o = SP(viewMain, Array(0#, 0#, 0#))
    L = Array(0#, 0.9375, z)                                        ' left end, half way up (datum C)

    ' along the top: 4.000 and the hole row BASIC from the left end (datum C)
    DimT viewMain, "H", L, Array(lenX, 0.9375, z), o(0) + 1.83, o(1) + 2.98, 4#, "", 3, "", False, 0, 0, 0, 3, "Main     4.000"
    DimT viewMain, "H", L, Array(hX(2), hU(2) + 0.25, z), o(0) + 1.53, o(1) + 2.76, 3.375, "", 3, "", False, TOL_BASIC, 0, 0, 3, "Main     3.375 BASIC (boxed)"
    DimT viewMain, "H", L, Array(hX(1), hU(1) + 0.22, z), o(0) + 1.33, o(1) + 2.5, 2.75, "", 3, "", False, TOL_BASIC, 0, 0, 3, "Main     2.750 BASIC (boxed)"
    DimT viewMain, "H", L, Array(hX(0), hU(0) + 0.1285, z), o(0) + 1.08, o(1) + 2.27, 2.125, "", 3, "", False, TOL_BASIC, 0, 0, 3, "Main     2.125 BASIC (boxed)"
    DatumOnEdge viewMain, Array(0#, 0.6, z), o(0) - 0.56, o(1) + 3.15, "C", "Main     datum C on the left end"

    ' heights on the right
    DimT viewMain, "V", Array(3#, lenU, z), Array(3#, 0#, z), o(0) + 5.77, o(1) + 0.94, 1.875, "", 3, "", False, 0, 0, 0, 3, "Main     1.875"
    DimT viewMain, "V", Array(hX(2), hU(2) + 0.25, z), Array(3#, 0#, z), o(0) + 5.22, o(1) + 0.65, 1.25, "", 3, "", False, TOL_BASIC, 0, 0, 3, "Main     1.250 BASIC (boxed)"
    DimT viewMain, "V", Array(hX(3) + 0.1005, hU(3), z), Array(3#, 0#, z), o(0) + 4.85, o(1) + 0.21, 0.5, "", 3, "", False, 0, 0, 0, 3, "Main     .500 (1/4-20 hole)"
    DimT viewMain, "V", Array(hX(4), hU(4) + 0.1005, z), Array(3#, 0#, z), o(0) + 4.4, o(1) - 0.33, 0.25, "", 3, "", False, 0, 0, 0, 3, "Main     .250 (" & D & ".201 hole)"
    DatumOnEdge viewMain, Array(2#, 0#, z), o(0) + 5.93, o(1) - 0.42, "B", "Main     datum B on the bottom edge"

    ' along the bottom
    DimT viewMain, "H", L, Array(hX(3) - 0.1005, hU(3), z), o(0) + 1.27, o(1) - 0.39, 2.5, "", 3, "", False, 0, 0, 0, 3, "Main     2.500 (1/4-20 hole)"
    DimT viewMain, "H", L, Array(hX(4) - 0.1005, hU(4), z), o(0) + 1.87, o(1) - 0.62, 3.75, "", 3, "", False, 0, 0, 0, 3, "Main     3.750 (" & D & ".201 hole)"

    ' radii and the small holes
    DimT viewMain, "R", Array(0.575 - 0.2298, 1.3 + 0.2298, z), Empty, o(0) - 0.75, o(1) + 2.15, 0.325, "4X ", 3, "", False, 0, 0, 0, 3, "Main     4X R.325 (logo pocket)"
    DimT viewMain, "R", Array(0.25 - 0.1768, 0.25 - 0.1768, z), Empty, o(0) - 0.78, o(1) - 0.08, 0.25, "4X ", 3, "", False, 0, 0, 0, 3, "Main     4X R.250 (corners)"
    DimT viewMain, "R", Array(hX(4) - 0.3536, hU(4) + 0.3536, z), Empty, o(0) + 4.56, o(1) - 0.61, 0.5, "", 2, "", False, 0, 0, 0, 2, "Main     R.50 (notch)"
    DimT viewMain, "D", Array(hX(4), hU(4) + 0.1005, z), Empty, o(0) + 2.99, o(1) - 0.39, 0.201, "", 3, "", False, 0, 0, 0, 3, "Main     " & D & ".201"
    NoteAt viewMain, Array(hX(3), hU(3) - 0.1005, z), Array(hX(3) - 0.1005, hU(3), z), "1/4-20 UNC", o(0) + 1.1, o(1) - 0.12, "Main     1/4-20 UNC note on the tapped hole"

    ' hole callouts, each with its position frame (dia .01 | A | B | C) under it
    HoleCall 0, " <HOLE-DEPTH> .38", o(0) + 5.41, o(1) + 3.44, "Main     " & D & ".257 deep .38"
    HoleCall 1, " <HOLE-SPOT> <MOD-DIAM> .44 <HOLE-DEPTH> .25", o(0) + 5.76, o(1) + 2.9, "Main     " & D & ".257 cbore " & D & ".44 deep .25"
    HoleCall 2, " <HOLE-SINK> <MOD-DIAM> .50 X 82<MOD-DEG>", o(0) + 5.76, o(1) + 2.33, "Main     " & D & ".257 csk " & D & ".50 X 82" & ChrW(176)
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' dia .257 hole callout on hole k (a diameter dimension with the reference's symbols after it),
' then its position frame attached under it
Sub HoleCall(ByVal k As Long, ByVal suffix As String, ByVal tx As Double, ByVal ty As Double, ByVal label As String)
    Dim why As String, t As String, dd As Object
    If Not hOK(k) Then Check False, label & " - MISSING (hole not where the reference has it)": Exit Sub
    why = DimP2(viewMain, "D", Array(hX(k), hU(k) + 0.1285, Z_TOP), Empty, tx, ty, 0.257, "", 3, suffix, False)
    If Len(why) = 0 Then
        Set dd = lastDim
        On Error Resume Next                              ' "dia .257" with a space, as on the reference
        t = "": t = dd.GetText(1)
        If InStr(UCase(t), "DIAM") > 0 Or InStr(t, ChrW(216)) > 0 Then dd.SetText 1, "<MOD-DIAM> "
        t = "": t = dd.GetText(2)
        On Error GoTo 0
        If Len(Trim(t)) = 0 Then why = "value OK, but its callout text after the " & ChrW(216) & ".257 did not stick; add it in Dimension Text"
    End If
    ReportW why, label
    If Len(why) = 0 Or InStr(why, "value OK") > 0 Then
        Set lastDim = dd
        MakeFrame3 viewMain, Empty, "<IGTOL-POSI>", True, ".01", "A", "B", "C", tx - 0.4, ty - 0.2, True, _
                   "Main     position " & ChrW(216) & ".01 | A | B | C under " & label
    End If
End Sub

' balloon 3 (note 3) with its leader on the logo
Sub Balloon3()
    On Error GoTo Oops
    Dim ents As Variant, i As Long, p As Variant, q As Variant, best As Long, bd As Double, d As Double, nt As Object
    Dim o As Variant, st As Long, hasLeader As Boolean
    o = SP(viewMain, Array(0#, 0#, 0#))
    ' the logo edge nearest the reference's leader end (the CP symbol's lower tip, x .66, u .85)
    ents = VisibleEdges(viewMain)
    best = -1: bd = 0.2
    If IsArray(ents) Then
        For i = LBound(ents) To UBound(ents)
            p = Empty
            On Error Resume Next
            p = ents(i).GetStartVertex.GetPoint
            On Error GoTo Oops
            If IsArray(p) Then
                q = ToPart(Arr3(p(0), p(1), p(2)))
                If q(2) < Z_FLOOR + 0.0002 And q(2) > Z_FLOOR - 0.02 And q(0) > 0.3 And q(0) < 1.4 And q(1) > 0.3 And q(1) < 1.6 Then
                    d = Sqr((q(0) - 0.66) ^ 2 + (q(1) - 0.85) ^ 2)
                    If d < bd Then bd = d: best = i
                End If
            End If
        Next
    End If
    swModel.ClearSelection2 True
    swDraw.ActivateView viewMain.GetName2
    If best >= 0 Then
        On Error Resume Next
        viewMain.SelectEntity ents(best), False
        On Error GoTo Oops
        hasLeader = (swModel.SelectionManager.GetSelectedObjectCount2(-1) = 1)
    End If
    Set nt = Nothing
    On Error Resume Next
    Set nt = swModel.InsertNote("3")
    If Not nt Is Nothing Then
        nt.SetBalloon 1, 2                                        ' circular, sized for two characters
        nt.GetAnnotation.SetPosition2 (o(0) - 0.64) * IN_M, (o(1) + 0.35) * IN_M, 0#
        st = -1: st = nt.GetBalloonStyle
    End If
    On Error GoTo Oops
    swModel.ClearSelection2 True
    If nt Is Nothing Then
        Check False, "Main     balloon 3 - MISSING; Annotation > Balloon on the logo, text 3"
    Else
        Check (st = 1), "Main     balloon 3 (circle) " & IIf(st = 1, "", "- the note is placed but its circle did not take (style " & st & "); click it and set Border: Circular")
        If Not hasLeader Then ToDo "Balloon 3: no logo edge was found for its leader - click it and drag the leader's end onto the logo"
    End If
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.ClearSelection2 True
End Sub

' SECTION B-B: widths below it, heights on the right, datum A, flatness, parallelism, chamfer note
Sub DimBB()
    On Error GoTo Oops
    Dim a As Variant, u As Double
    If viewBB Is Nothing Then Check False, "SECTION B-B dimensions - no SECTION B-B": Exit Sub
    u = UB
    eKey = ""
    CollectEdges viewBB
    a = SP(viewBB, Array(0#, u, 0#))                                ' left end on the back face

    DimT viewBB, "H", Array(0#, u, 0.15), Array(0.125, u, 0.38), a(0) - 0.485, a(1) - 0.62, 0.125, "", 3, "", False, 0, 0, 0, 3, "B-B      .125 (rim)"
    DimT viewBB, "H", Array(0.125, u, 0.38), Array(0.25, u, 0.42), a(0) - 0.36, a(1) - 0.84, 0.125, "", 3, "", False, 0, 0, 0, 3, "B-B      .125 (to the pocket)"
    DimT viewBB, "H", Array(0.25, u, 0.42), Array(1.4375, u, 0.42), a(0) + 0.9, a(1) - 0.98, 1.1875, "", 4, "", False, 0, 0, 0, 4, "B-B      1.1875 (pocket)"
    If gotGroove Then
        DimT viewBB, "H", Array(0.125, u, 0.38), Array(xGroove, u, zGroove - 0.125), a(0) + 0.86, a(1) - 1.2, xGroove - 0.125, "", 3, "", False, 0, 0, 0, 3, "B-B      1.563 (to the groove's centre)"
    Else
        Check False, "B-B      1.563 - MISSING (groove not found)"
    End If
    DimT viewBB, "H", Array(0.125, u, 0.38), Array(lenX - 0.125, u, 0.38), a(0) + 1.96, a(1) - 1.42, 3.75, "", 3, "", False, 0, 0, 0, 3, "B-B      3.750 (platform)"
    DimT viewBB, "V", Array(lenX - 0.06, u, Z_RIM), Array(3.75, u, Z_TOP), a(0) + 4.14, a(1) - 0.95, 0.2, "", 3, "", False, 0, 0, 0, 3, "B-B      .200 (rim step)"
    DimT viewBB, "V", Array(1#, u, 0#), Array(3.75, u, Z_TOP), a(0) + 4.75, a(1) - 0.95, 0.48, "", 3, "", False, 0, 0, 0, 3, "B-B      .480"

    DatumOnEdge viewBB, Array(0.675, u, 0#), a(0) + 0.675, a(1) + 0.42, "A", "B-B      datum A on the back face"
    MakeFrame3 viewBB, Array(3.9, u, 0#), "<IGTOL-FLAT>", False, ".002", "", "", "", a(0) + 4.89, a(1) + 0.25, False, "B-B      flatness .002 on the back face"
    MakeFrame3 viewBB, Array(3.75, u, Z_TOP), "<IGTOL-PARA>", False, ".002", "A", "", "", a(0) + 4.87, a(1) - 0.5, False, "B-B      parallelism .002 | A on the platform face"
    NoteAt viewBB, Array(lenX - 0.0125, u, 0.0125), Array(lenX, u, 0.1), ".025 X 45<MOD-DEG>", a(0) + 4.73, a(1) + 0.47, "B-B      .025 X 45" & ChrW(176) & " chamfer note"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' SECTION A-A: heights on its left, pocket depth below
Sub DimAA()
    On Error GoTo Oops
    Dim a As Variant, x As Double
    If viewAA Is Nothing Then Check False, "SECTION A-A dimensions - no SECTION A-A": Exit Sub
    x = XA
    eKey = ""
    CollectEdges viewAA
    a = SP(viewAA, Array(x, 0#, 0#))                                ' back face, bottom edge

    DimT viewAA, "V", Array(x, 1.75, 0.38), Array(x, 0.125, 0.38), a(0) - 1.685, a(1) + 0.9, 1.625, "", 3, "", False, 0, 0, 0, 3, "A-A      1.625 (platform)"
    DimT viewAA, "V", Array(x, 1.625, 0.42), Array(x, 0.25, 0.42), a(0) - 1.11, a(1) + 0.9, 1.375, "", 3, "", False, 0, 0, 0, 3, "A-A      1.375 (pocket)"
    DimT viewAA, "V", Array(x, lenU, 0.15), Array(x, 1.75, 0.38), a(0) - 0.66, a(1) + 2.4, 0.125, "2 X ", 3, "", False, 0, 0, 0, 3, "A-A      2 X .125 (rim)"
    DimT viewAA, "V", Array(x, 1.75, 0.38), Array(x, 1.625, 0.42), a(0) - 0.87, a(1) + 2.01, 0.125, "2 X", 3, "", False, 0, 0, 0, 3, "A-A      2 X.125 (to the pocket)"
    DimT viewAA, "H", Array(x, 0.19, Z_TOP), Array(x, 0.3, Z_FLOOR), a(0) - 0.97, a(1) - 0.15, 0.125, "", 3, "", False, 0, 0, 0, 3, "A-A      .125 (pocket depth)"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' a feature control frame with up to three datums. onDim: attached to lastDim (sits under its text);
' else its leader goes to the edge at lp (one click there). Judged on what reads back.
Sub MakeFrame3(v As Object, lp As Variant, ByVal sym As String, ByVal dia As Boolean, ByVal tol1 As String, _
               ByVal d1 As String, ByVal d2 As String, ByVal d3 As String, ByVal tx As Double, ByVal ty As Double, _
               ByVal onDim As Boolean, ByVal label As String)
    Dim g As Object, attached As Boolean, rb As String, want As String
    swModel.ClearSelection2 True
    If onDim Then
        If Not lastDim Is Nothing Then
            On Error Resume Next
            lastDim.GetAnnotation.Select3 False, Nothing
            On Error GoTo 0
        End If
        attached = (swModel.SelectionManager.GetSelectedObjectCount2(-1) = 1)
    ElseIf IsArray(lp) Then
        attached = PickOnly(v, lp, False)
    End If
    Set g = Nothing
    On Error Resume Next
    Set g = swModel.InsertGtol
    If Not g Is Nothing Then
        Err.Clear
        g.SetFrameSymbols2 1, sym, dia, "", False, "", "", "", ""
        If Err.Number <> 0 Then Err.Clear: g.SetFrameSymbols3 1, sym, dia, "", False, "", "", "", ""
        Err.Clear
        g.SetFrameValues2 1, tol1, "", d1, d2, d3
        If Err.Number <> 0 Then Err.Clear: g.SetFrameValues 1, tol1, "", d1, d2, d3
        If Not (onDim And attached) Then g.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
        rb = GtolReadBack(g)
    End If
    On Error GoTo 0
    swModel.ClearSelection2 True
    want = tol1 & IIf(Len(d1) > 0, " | " & d1, "") & IIf(Len(d2) > 0, " | " & d2, "") & IIf(Len(d3) > 0, " | " & d3, "")
    If g Is Nothing Then
        Check False, label & " - MISSING; add it with Annotations > Geometric Tolerance"
        Exit Sub
    End If
    If Len(rb) = 0 Then
        Info label & " placed; SolidWorks does not report its values - look that it reads " & want
    ElseIf InStr(rb, tol1) > 0 And InStr(rb, d1) > 0 And InStr(rb, d2) > 0 And InStr(rb, d3) > 0 Then
        Check True, label & " (reads " & rb & ")"
    Else
        Check False, label & " - reads " & rb & "; double-click it and set " & want
    End If
    Info "  symbol " & sym & IIf(SymInLibrary(sym), " found in your gtol.sym", " not confirmed in gtol.sym - look that the symbol shows")
    If Not attached Then ToDo label & ": the frame is placed without " & IIf(onDim, "being attached to its dimension - drag it under the dimension", "a leader - drag its leader handle to the face as on the reference")
End Sub

' =====================================================================================
'  Picking, dimension, GD&T and note helpers  (from HW#8C)
' =====================================================================================
' a dimension between the edges at part points pa (and pb): edge table first, then single clicks
Function DimP2(v As Object, ByVal kind As String, pa As Variant, pb As Variant, ByVal tx As Double, ByVal ty As Double, _
               ByVal expect As Double, ByVal prefix As String, ByVal prec As Long, ByVal suffix As String, ByVal paren As Boolean) As String
    Dim why As String, need As Long, ok As Boolean
    If IsArray(pb) Then need = 2 Else need = 1
    why = TryDim(v, kind, pa, pb, tx, ty, expect, prefix, prec, suffix, paren, True)
    If Len(why) = 0 Then DimP2 = "": Exit Function
    ok = PickOnly(v, pa, False)
    If need = 2 Then ok = PickOnly(v, pb, True) And ok
    If ok Then
        why = TryDimSel(kind, need, tx, ty, expect, prefix, prec, suffix, paren, True)
    Else
        why = why & "; single-click picks failed too"
    End If
    DimP2 = why
End Function

' DimP2, then a tolerance (tolType 0 = none), then one check line
Sub DimT(v As Object, ByVal kind As String, pa As Variant, pb As Variant, ByVal tx As Double, ByVal ty As Double, _
         ByVal expect As Double, ByVal prefix As String, ByVal prec As Long, ByVal suffix As String, ByVal paren As Boolean, _
         ByVal tolType As Long, ByVal tMin As Double, ByVal tMax As Double, ByVal tolPrec As Long, ByVal label As String)
    Dim why As String
    why = DimP2(v, kind, pa, pb, tx, ty, expect, prefix, prec, suffix, paren)
    If Len(why) = 0 And tolType > 0 Then why = SetTol(lastDim, tolType, tMin, tMax, prec, tolPrec)
    ReportW why, label
End Sub

Sub ReportW(ByVal why As String, ByVal label As String)
    Check (Len(why) = 0), label & IIf(Len(why) = 0, "", "  - " & why)
End Sub

' datum feature symbol on the edge at part point p, its box at (tx, ty)
Sub DatumOnEdge(v As Object, p As Variant, ByVal tx As Double, ByVal ty As Double, ByVal lbl As String, ByVal label As String)
    Dim dt As Object
    Set dt = Nothing
    If PickOnly(v, p, False) Then
        On Error Resume Next
        Set dt = swModel.InsertDatumTag2
        If Not dt Is Nothing Then dt.SetLabel lbl: dt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
        On Error GoTo 0
    End If
    swModel.ClearSelection2 True
    If dt Is Nothing Then
        Check False, label & " - MISSING; Annotations > Datum Feature on that face, label " & lbl
    Else
        Check (UCase(Trim(DatumLabel(dt))) = lbl), label & " (label reads " & DatumLabel(dt) & ")"
    End If
End Sub

Function SkLineP(v As Object, a As Variant, b As Variant, ByVal cand As Integer) As Object
    Dim p As Variant, q As Variant
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    p = ToSk(v, SP(v, a), cand): q = ToSk(v, SP(v, b), cand)
    If Not IsArray(p) Or Not IsArray(q) Then Exit Function
    On Error Resume Next
    swModel.SketchManager.AddToDB = True
    Set SkLineP = swModel.SketchManager.CreateLine(p(0), p(1), 0#, q(0), q(1), 0#)
    swModel.SketchManager.AddToDB = False
    eKey = ""
    swModel.ClearSelection2 True
End Function

' makes one dimension; returns "" when it measured right, otherwise the reason.
' deleteIfWrong: remove a wrong dimension so another way can be tried.
Function TryDim(v As Object, ByVal kind As String, pa As Variant, pb As Variant, ByVal tx As Double, ByVal ty As Double, _
                ByVal expect As Double, ByVal prefix As String, ByVal prec As Long, ByVal suffix As String, _
                ByVal paren As Boolean, ByVal deleteIfWrong As Boolean) As String
    Dim need As Long
    If kind = "H" Or kind = "V" Or kind = "S" Or kind = "A" Then need = 2 Else need = 1
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    SelPt v, pa, False
    If need = 2 Then SelPt v, pb, True
    TryDim = TryDimSel(kind, need, tx, ty, expect, prefix, prec, suffix, paren, deleteIfWrong)
End Function

' makes a dimension from the current selection; returns "" when it measured right, otherwise the reason
Function TryDimSel(ByVal kind As String, ByVal need As Long, ByVal tx As Double, ByVal ty As Double, _
                   ByVal expect As Double, ByVal prefix As String, ByVal prec As Long, ByVal suffix As String, _
                   ByVal paren As Boolean, ByVal deleteIfWrong As Boolean) As String
    Dim dd As Object, got As Double, nSel As Long, txt As String, ok As Boolean
    nSel = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    If nSel <> need Then
        swModel.ClearSelection2 True
        TryDimSel = "MISSING (found " & nSel & " of " & need & " edge(s))"
        Exit Function
    End If
    Set dd = Nothing
    Set lastDim = Nothing
    On Error Resume Next
    Select Case kind
        Case "H": Set dd = swModel.AddHorizontalDimension2(tx * IN_M, ty * IN_M, 0#)
        Case "V", "VL": Set dd = swModel.AddVerticalDimension2(tx * IN_M, ty * IN_M, 0#)
        Case "D": Set dd = swModel.AddDiameterDimension2(tx * IN_M, ty * IN_M, 0#)
        Case "R": Set dd = swModel.AddRadialDimension2(tx * IN_M, ty * IN_M, 0#)
        Case Else: Set dd = swModel.AddDimension2(tx * IN_M, ty * IN_M, 0#)
    End Select
    On Error GoTo 0
    swModel.ClearSelection2 True
    If dd Is Nothing Then TryDimSel = "MISSING (SolidWorks did not create it)": Exit Function
    Set lastDim = dd

    On Error Resume Next
    got = -1
    got = dd.GetDimension2(0).SystemValue
    dd.ShowParenthesis = paren
    If prec >= 0 Then
        dd.SetPrecision3 prec, -1, -1, -1
        If Err.Number <> 0 Then Err.Clear: dd.SetPrecision2 prec, -1, -1, -1
    End If
    If Len(prefix) > 0 Then
        txt = ""
        txt = dd.GetText(1)
        If InStr(txt, prefix) = 0 Then dd.SetText 1, prefix & txt          ' keep any symbol already there
        txt = dd.GetText(1)
    End If
    If Len(suffix) > 0 Then dd.SetText 2, suffix
    On Error GoTo 0

    If kind = "A" Then
        got = got * 180# / 3.14159265358979
        ok = (Abs(got - expect) < 0.3)
    Else
        got = got / IN_M
        ok = (Abs(got - expect) < 0.003)
    End If
    If Not ok Then
        TryDimSel = "WRONG (measured " & Format(got, "0.000") & ")"
        If deleteIfWrong Then                          ' deletion confirmed (HW#8A lesson 2)
            Set lastDim = Nothing
            If Not DeleteDim(dd) Then TryDimSel = TryDimSel & "; that wrong dimension (" & Format(got, "0.000") & _
                ") could not be deleted - click it and press Delete"
        End If
    ElseIf Len(prefix) > 0 And InStr(txt, prefix) = 0 Then
        TryDimSel = "value OK, but the " & IIf(prefix = "<MOD-DIAM>", ChrW(216) & " symbol", "prefix") & " did not stick; add it by hand"
    End If
End Function

' select the edge passing through part point p: first from the view's edge table, else by picking on the sheet
Function SelPt(v As Object, p As Variant, ByVal append As Boolean) As Boolean
    Dim s As Variant, i As Long, best As Long, bd As Double, d As Double, n0 As Long
    s = SheetPt(v, CDbl(p(0)), CDbl(p(1)), CDbl(p(2)))
    best = -1: bd = 0.008
    For i = 0 To eN - 1
        d = EntDist(i, s(0), s(1))
        If d < bd Then bd = d: best = i
    Next
    n0 = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    If best >= 0 Then
        SelEdge v, best, append
    Else
        On Error Resume Next
        swModel.Extension.SelectByID2 "", "EDGE", s(0) * IN_M, s(1) * IN_M, 0#, append, 0, Nothing, 0
        If swModel.SelectionManager.GetSelectedObjectCount2(-1) <= n0 Then
            swModel.Extension.SelectByID2 "", "SILHOUETTE", s(0) * IN_M, s(1) * IN_M, 0#, append, 0, Nothing, 0
        End If
        On Error GoTo 0
        KeepOneIn v, IIf(append, n0, 0)
    End If
    SelPt = (swModel.SelectionManager.GetSelectedObjectCount2(-1) > IIf(append, n0, 0))
End Function

' distance (inches) from a sheet point to table entity i
Function EntDist(ByVal i As Long, ByVal x As Double, ByVal y As Double) As Double
    Dim dx As Double, dy As Double, L2 As Double, t As Double
    If eKind(i) = "C" Then
        EntDist = Abs(Sqr((x - eCX(i)) ^ 2 + (y - eCY(i)) ^ 2) - eR(i))
    Else
        dx = eX2(i) - eX1(i): dy = eY2(i) - eY1(i): L2 = dx * dx + dy * dy
        If L2 < 1E-12 Then EntDist = 1E+30: Exit Function
        t = ((x - eX1(i)) * dx + (y - eY1(i)) * dy) / L2
        If t < 0 Then t = 0
        If t > 1 Then t = 1
        EntDist = Sqr((x - eX1(i) - t * dx) ^ 2 + (y - eY1(i) - t * dy) ^ 2)
    End If
End Function

Sub SelEdge(v As Object, ByVal i As Long, ByVal append As Boolean)
    Dim ok As Boolean, sd As Object, n0 As Long
    On Error Resume Next
    n0 = 0
    If append Then n0 = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    ok = v.SelectEntity(eObj(i), append)
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) <= n0 Then
        Err.Clear
        Set sd = swModel.SelectionManager.CreateSelectData
        Set sd.View = v
        If Err.Number <> 0 Then Err.Clear: sd.View = v
        ok = eObj(i).Select4(append, sd)
    End If
    KeepOneIn v, n0
End Sub

' After a pick, keep exactly one new selection and only one that belongs to view v.
' (An edge shown in a parent view and in its detail view gets selected in both.)
Sub KeepOneIn(v As Object, ByVal n0 As Long)
    Dim sm As Object, n As Long, i As Long, dv As Object, keptOne As Boolean, nm As String
    On Error Resume Next
    Set sm = swModel.SelectionManager
    nm = v.GetName2
    n = sm.GetSelectedObjectCount2(-1)
    For i = n To n0 + 1 Step -1                       ' from the end, so indices stay valid
        Set dv = Nothing
        Set dv = sm.GetSelectedObjectsDrawingView2(i, -1)
        If Not dv Is Nothing Then
            If dv.GetName2 <> nm Then sm.DeSelect2 i, -1
        End If
    Next
    n = sm.GetSelectedObjectCount2(-1)
    For i = n To n0 + 2 Step -1                       ' still more than one new item: keep the first
        sm.DeSelect2 i, -1
    Next
End Sub

' select the edge at part point p by one click at its exact sheet position (no edge-table read);
' True when exactly one new item, in view v, is selected
Function PickOnly(v As Object, p As Variant, ByVal append As Boolean) As Boolean
    Dim s As Variant, n0 As Long
    s = SheetPt(v, CDbl(p(0)), CDbl(p(1)), CDbl(p(2)))
    If Not append Then                                 ' (activating a view can drop the current selection)
        swModel.ClearSelection2 True
        swDraw.ActivateView v.GetName2
    End If
    n0 = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    On Error Resume Next
    swModel.Extension.SelectByID2 "", "EDGE", s(0) * IN_M, s(1) * IN_M, 0#, append, 0, Nothing, 0
    On Error GoTo 0
    KeepOneIn v, n0
    PickOnly = (swModel.SelectionManager.GetSelectedObjectCount2(-1) = n0 + 1)
End Function

' a note with a leader to the edge at p1 (else p2), picked by a click at that exact point (no edge-table read)
Sub NoteAt(v As Object, p1 As Variant, p2 As Variant, ByVal text As String, ByVal tx As Double, ByVal ty As Double, ByVal label As String)
    Dim nt As Object, hasLeader As Boolean
    hasLeader = PickOnly(v, p1, False)
    If Not hasLeader Then hasLeader = PickOnly(v, p2, False)
    Set nt = Nothing
    On Error Resume Next
    Set nt = swModel.InsertNote(text)
    If Not nt Is Nothing Then nt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
    On Error GoTo 0
    swModel.ClearSelection2 True
    If nt Is Nothing Then
        Check False, label & "  - MISSING"
    Else
        Check True, label
        If Not hasLeader Then ToDo label & ": the note is placed without a leader - click it and drag its leader handle to the part as on the reference"
    End If
End Sub

' a sheet note at (x, y) = its top-left; True when it exists and its text reads back
Function PlaceNote(ByVal text As String, ByVal x As Double, ByVal y As Double) As Boolean
    Dim nt As Object, t As String
    On Error Resume Next
    swDraw.ActivateSheet swDraw.GetCurrentSheet.GetName
    swModel.ClearSelection2 True
    Set nt = Nothing
    Set nt = swModel.InsertNote(text)
    If nt Is Nothing Then Exit Function
    nt.GetAnnotation.SetPosition2 x * IN_M, y * IN_M, 0#
    t = nt.GetText
    PlaceNote = (Len(t) > 0 And InStr(UCase(t), UCase(Left(Replace(text, "<MOD-DEG>", ""), 6))) > 0)
End Function

Function DatumLabel(dt As Object) As String
    On Error Resume Next
    DatumLabel = dt.GetLabel
End Function

' what the frame says (tolerance and datum of frame 1), "" if SolidWorks does not tell
Function GtolReadBack(g As Object) As String
    Dim fv As Variant, k As Long, t As String
    On Error Resume Next
    fv = Empty
    fv = g.GetFrameValues2(1)
    If Not IsArray(fv) Then fv = g.GetFrameValues(1)
    If IsArray(fv) Then
        For k = LBound(fv) To UBound(fv)
            If Len(Trim(CStr(fv(k)))) > 0 Then t = t & IIf(Len(t) > 0, " | ", "") & Trim(CStr(fv(k)))
        Next
    End If
    GtolReadBack = t
End Function

' is a symbol code (e.g. <IGTOL-SRUN>) listed in this computer's gtol.sym?
Function SymInLibrary(ByVal sym As String) As Boolean
    Dim base As String, d As String, p As String, f As Integer, t As String, nm As String
    On Error Resume Next
    nm = Mid(sym, InStr(sym, "-") + 1): nm = Left(nm, Len(nm) - 1)
    base = Environ("ProgramData") & "\SOLIDWORKS\"
    d = Dir(base & "SOLIDWORKS 20*", vbDirectory)
    Do While Len(d) > 0 And Len(p) = 0
        If FileExists(base & d & "\lang\english\gtol.sym") Then p = base & d & "\lang\english\gtol.sym"
        d = Dir
    Loop
    If Len(p) = 0 Then Exit Function
    f = FreeFile
    Open p For Input As #f
    t = Input(LOF(f), #f)
    Close #f
    SymInLibrary = (InStr(UCase(t), UCase(nm)) > 0)
End Function

Sub ClearCommand()
    On Error Resume Next
    swModel.ClearSelection2 True
    swDraw.ActivateSheet swDraw.GetCurrentSheet.GetName
    swModel.ClearSelection2 True
End Sub

' sheet point (inches) -> view sketch point (metres).  cand 0: the view sketch's own transform;
' 1: sheet coordinates; 2: from the model origin; 3: from the view centre   (from HW#8A)
Function ToSk(v As Object, p As Variant, ByVal cand As Integer) As Variant
    Dim o As Variant, sc As Double, pos As Variant, sk As Object, xf As Object, a(2) As Double, va As Variant
    Dim q As Object, r As Variant
    sc = 1#
    On Error Resume Next
    sc = v.ScaleDecimal
    If sc <= 0 Then sc = 1#
    Select Case cand
        Case 0
            Set sk = v.GetSketch
            Set xf = sk.ModelToSketchTransform
            If xf Is Nothing Then Exit Function
            a(0) = p(0) * IN_M: a(1) = p(1) * IN_M: a(2) = 0#: va = a
            Set q = swMath.CreatePoint((va))
            Set q = q.MultiplyTransform(xf)
            r = q.ArrayData
            If IsArray(r) Then ToSk = Arr3(r(0), r(1), 0#)
        Case 1
            ToSk = Arr3(p(0) * IN_M, p(1) * IN_M, 0#)
        Case 2
            o = XfPt(v.ModelToViewTransform, 0#, 0#, 0#)
            ToSk = Arr3((p(0) - o(0)) * IN_M / sc, (p(1) - o(1)) * IN_M / sc, 0#)
        Case 3
            pos = v.Position
            ToSk = Arr3((p(0) * IN_M - pos(0)) / sc, (p(1) * IN_M - pos(1)) / sc, 0#)
    End Select
End Function

' horizontal test dimension from sketch item it to the edge through part point pEdge; value in
' inches (-1 if none). The test dimension is deleted and the deletion confirmed (HW#8A lesson 2).
Function TestH(v As Object, it As Object, pEdge As Variant) As Double
    Dim dd As Object, n As Long, tx As Double, ty As Double
    TestH = -1
    tx = SX(v, CDbl(pEdge(0))): ty = SY(v, 0.9)
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    On Error Resume Next
    it.Select4 False, Nothing
    n = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    On Error GoTo 0
    If n <> 1 Then swModel.ClearSelection2 True: Exit Function
    If Not PickOnly(v, pEdge, True) Then
        If Not SelPt(v, pEdge, True) Then swModel.ClearSelection2 True: Exit Function   ' one edge-table read, if needed
    End If
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) <> 2 Then swModel.ClearSelection2 True: Exit Function
    Set dd = Nothing
    On Error Resume Next
    Set dd = swModel.AddHorizontalDimension2(tx * IN_M, ty * IN_M, 0#)
    On Error GoTo 0
    swModel.ClearSelection2 True
    If dd Is Nothing Then Exit Function
    On Error Resume Next
    TestH = dd.GetDimension2(0).SystemValue / IN_M
    On Error GoTo 0
    If Not DeleteDim(dd) Then Check False, "A test dimension (" & Format(TestH, "0.000") & ") could not be deleted - click it and press Delete"
End Function

' delete a dimension and confirm it is gone (HW#8A: Select3 + EditDelete silently left dimensions behind)
Function DeleteDim(dd As Object) As Boolean
    Dim n0 As Long, nm As String
    On Error Resume Next
    n0 = TotalDims()
    nm = dd.GetNameForSelection
    swModel.ClearSelection2 True
    dd.GetAnnotation.Select3 False, Nothing
    swModel.EditDelete
    swModel.ClearSelection2 True
    If TotalDims() < n0 Then DeleteDim = True: Exit Function
    If Len(nm) > 0 Then
        swModel.Extension.SelectByID2 nm, "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0
        swModel.EditDelete
        swModel.ClearSelection2 True
    End If
    DeleteDim = (TotalDims() < n0)
End Function

' every dimension on the sheet (to confirm a deletion)
Function TotalDims() As Long
    Dim v As Object, n As Long
    On Error Resume Next
    Set v = swDraw.GetFirstView
    Do While Not v Is Nothing
        n = n + v.GetDisplayDimensionCount
        Set v = v.GetNextView
    Loop
    TotalDims = n
End Function

Function SetTol(dd As Object, ByVal tolType As Long, ByVal tMin As Double, ByVal tMax As Double, _
                ByVal prec As Long, ByVal tolPrec As Long) As String
    Dim tol As Object, tp As Long, way As Integer, ok As Boolean, cfg As String
    If dd Is Nothing Then SetTol = "value OK, but the dimension was lost before its tolerance": Exit Function
    On Error Resume Next
    Set tol = dd.GetDimension2(0).Tolerance
    If tol Is Nothing Then
        SetTol = "value OK, but its tolerance could not be reached; set it in the PropertyManager"
        Exit Function
    End If
    tol.Type = tolType
    cfg = "": cfg = swPart.ConfigurationManager.ActiveConfiguration.Name
    ' v1 set the values silently to 0, so each way is tried until the values read back right
    For way = 1 To 5
        Err.Clear
        Select Case way
            Case 1: tol.SetValues2 tMin * IN_M, tMax * IN_M, 1, ""                 ' this configuration (SolidWorks' example)
            Case 2: tol.SetValues2 tMin * IN_M, tMax * IN_M, 2, ""                 ' all configurations
            Case 3: tol.SetValues tMin * IN_M, tMax * IN_M
            Case 4: tol.SetValues2 tMin * IN_M, tMax * IN_M, 3, Array(cfg)         ' the part's configuration by name
            Case 5: tol.Type = tolType: tol.SetValues2 tMin * IN_M, tMax * IN_M, 1, Empty
        End Select
        If TolReads(tol, tolType, tMin, tMax) Then ok = True: Exit For
    Next
    Err.Clear
    dd.SetPrecision3 prec, -1, tolPrec, -1
    If Err.Number <> 0 Then Err.Clear: dd.SetPrecision2 prec, -1, tolPrec, -1
    tp = -1: tp = tol.Type
    If Not ok And tolType = TOL_SYM Then                  ' last resort for +-: write it as text after the value
        tol.Type = 0
        dd.SetText 2, "<MOD-PM>" & Mid(Format(tMax, "0.000"), 2)
        ok = True: tp = tolType: way = 6
    End If
    swModel.GraphicsRedraw2
    On Error GoTo 0
    If tp <> tolType Then
        SetTol = "value OK, but the tolerance type did not stick; set it in the PropertyManager"
    ElseIf Not ok Then
        SetTol = "value OK, but the tolerance values would not stick (they read " & TolText(tol) & "); set them in the PropertyManager"
    ElseIf way = 6 Then
        Info "  (tolerance written as text after the value: SolidWorks would not take the tolerance values)"
    End If
End Function

Function TolReads(tol As Object, ByVal tolType As Long, ByVal tMin As Double, ByVal tMax As Double) As Boolean
    Dim mn As Double, mx As Double
    On Error Resume Next
    mn = -99: mx = -99
    mx = tol.GetMaxValue / IN_M
    mn = tol.GetMinValue / IN_M
    TolReads = (Abs(mx - tMax) < 0.00005) And (tolType = TOL_SYM Or Abs(mn - tMin) < 0.00005)
End Function

Function TolText(tol As Object) As String
    Dim mn As Double, mx As Double
    On Error Resume Next
    mn = -99: mx = -99
    mx = tol.GetMaxValue / IN_M
    mn = tol.GetMinValue / IN_M
    TolText = "+" & Format(mx, "0.000") & " / " & Format(mn, "0.000")
End Function

' NOTES block, top left: each number and each text is its own note so the columns line up exactly
Sub AddNotes()
    On Error GoTo Oops
    Dim lines As Variant, i As Long, y As Double, ok As Boolean
    lines = Array("USE THE IN-LAB ASSIGNMENT #8D PART MODEL ON CANVAS TO RE-CREATE THE ENGINEERING DRAWING", _
                  "MAKE SURE YOUR DRAWING'S TOLERANCES MATCH THE DRAWING TOLERANCES SHOWN BELOW.", _
                  "ENGRAVE THE MANUFACTURING ENGINEERING LOGO ON THE PART .003"" DEEP")
    ok = PlaceNote("NOTES:", NOTE_X, NOTE_Y)
    For i = 0 To 2
        y = NOTE_Y - (i + 1) * NOTE_DY
        ok = PlaceNote((i + 1) & ".", NOTE_X, y) And ok
        ok = PlaceNote(CStr(lines(i)), NOTE_X + NOTE_TAB, y) And ok
    Next
    Check ok, "NOTES 1-3 top left (each read back)"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.ClearSelection2 True
End Sub

' =====================================================================================
'  Edge table  (from HW#8C)
' =====================================================================================
Sub CollectEdges(v As Object)
    Dim ents As Variant, i As Long, xf As Object, sc As Double, pos As Variant, key As String, m As Long
    ' read a view's edges once; read again only if the view moved, changed scale or display (v2: speed)
    On Error Resume Next
    pos = v.Position: m = -1: m = v.GetDisplayMode2
    key = v.GetName2 & "|" & Format(pos(0), "0.000000") & "|" & Format(pos(1), "0.000000") & "|" & v.ScaleDecimal & "|" & m
    On Error GoTo 0
    If key = eKey And eN > 0 Then Exit Sub
    eKey = key
    eN = 0
    ReDim eObj(0 To EMAX - 1): ReDim eKind(0 To EMAX - 1)
    ReDim eX1(0 To EMAX - 1): ReDim eY1(0 To EMAX - 1): ReDim eX2(0 To EMAX - 1): ReDim eY2(0 To EMAX - 1)
    ReDim eCX(0 To EMAX - 1): ReDim eCY(0 To EMAX - 1): ReDim eR(0 To EMAX - 1): ReDim eFull(0 To EMAX - 1)
    xMin = 1E+30: xMax = -1E+30: yMin = 1E+30: yMax = -1E+30

    Set xf = v.ModelToViewTransform
    sc = 1#
    On Error Resume Next
    sc = v.ScaleDecimal
    On Error GoTo 0
    If sc <= 0 Then sc = 1#

    ents = VisibleEdges(v)
    If IsArray(ents) Then
        For i = LBound(ents) To UBound(ents)
            If eN < EMAX Then AddEdge ents(i), xf, sc
        Next
    End If
    report = report & "info     " & v.GetName2 & ": " & eN & " edges read" & vbCrLf
End Sub

Function VisibleEdges(v As Object) As Variant
    Dim tp As Variant, r As Variant, best As Variant, i As Long, n As Long, bestN As Long, c As Object
    For Each tp In Array(2, 1, 3, 4)                  ' swViewEntityType_e: try each, keep the one with edges
        r = Empty
        On Error Resume Next
        r = v.GetVisibleEntities2(Nothing, CLng(tp))
        If Not IsArray(r) Then r = v.GetVisibleEntities(Nothing, CLng(tp))
        On Error GoTo 0
        If IsArray(r) Then
            n = 0                                     ' v2: judge each type on a sample of 40 (speed)
            For i = LBound(r) To IIf(UBound(r) > LBound(r) + 39, LBound(r) + 39, UBound(r))
                Set c = Nothing
                On Error Resume Next
                Set c = r(i).GetCurve
                On Error GoTo 0
                If Not c Is Nothing Then n = n + 1
            Next
            n = n * (UBound(r) - LBound(r) + 1) / IIf(UBound(r) - LBound(r) + 1 > 40, 40, UBound(r) - LBound(r) + 1)
            If n > bestN Then bestN = n: best = r
        End If
    Next
    VisibleEdges = best
End Function

Sub Extend(p As Variant)
    If p(0) < xMin Then xMin = p(0)
    If p(0) > xMax Then xMax = p(0)
    If p(1) < yMin Then yMin = p(1)
    If p(1) > yMax Then yMax = p(1)
End Sub

Sub AddEdge(ByVal ent As Object, xf As Object, ByVal sc As Double)
    ' Extents come only from exact geometry: line end points, and circle/arc centre +- radius.
    Dim crv As Object, cp As Variant, a As Variant, b As Variant, p1 As Variant, p2 As Variant
    Dim isL As Boolean, isC As Boolean, cc As Variant, axv As Variant, ctr As Variant, L As Double, rr As Double

    On Error Resume Next
    Set crv = ent.GetCurve
    isL = crv.IsLine
    isC = crv.IsCircle
    If Not (isL Or isC) Then Exit Sub                ' helix / spline edges (modelled threads) are never picked: skip them at once
    cp = ent.GetCurveParams2                          ' start point (0-2), end point (3-5)
    p1 = Empty: p2 = Empty
    p1 = ent.GetStartVertex.GetPoint                  ' vertices are the most reliable end points
    p2 = ent.GetEndVertex.GetPoint
    On Error GoTo 0
    If crv Is Nothing Then Exit Sub
    If Not IsArray(p1) Or Not IsArray(p2) Then
        If Not IsArray(cp) Then Exit Sub
        p1 = Arr3(cp(0), cp(1), cp(2))
        p2 = Arr3(cp(3), cp(4), cp(5))
    End If
    a = XfPt(xf, p1(0), p1(1), p1(2))
    b = XfPt(xf, p2(0), p2(1), p2(2))

    If isL Then
        L = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
        If L < 0.003 Then Exit Sub                    ' edge seen end-on
        Extend a: Extend b
        Set eObj(eN) = ent: eKind(eN) = "L"
        eX1(eN) = a(0): eY1(eN) = a(1): eX2(eN) = b(0): eY2(eN) = b(1)
        eN = eN + 1
    ElseIf isC Then
        On Error Resume Next
        cc = crv.CircleParams                         ' centre(3) axis(3) radius
        On Error GoTo 0
        If Not IsArray(cc) Then Exit Sub
        axv = XfVec(xf, Arr3(cc(3), cc(4), cc(5)))
        ctr = XfPt(xf, cc(0), cc(1), cc(2))
        rr = cc(6) / IN_M * sc
        If Abs(axv(2)) < 0.1 Then
            ' circle or arc seen exactly side-on: on the sheet it is a straight line
            Dim dx As Double, dy As Double, dl As Double
            If Abs(a(0) - b(0)) + Abs(a(1) - b(1)) < 0.0005 Then    ' full circle: centre +- radius
                dx = -axv(1): dy = axv(0): dl = Sqr(dx * dx + dy * dy)
                If dl < 0.000001 Then Exit Sub
                a = Arr3(ctr(0) - rr * dx / dl, ctr(1) - rr * dy / dl, 0#)
                b = Arr3(ctr(0) + rr * dx / dl, ctr(1) + rr * dy / dl, 0#)
            End If
            L = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
            If L < 0.003 Then Exit Sub
            Extend a: Extend b
            Set eObj(eN) = ent: eKind(eN) = "L"
            eX1(eN) = a(0): eY1(eN) = a(1): eX2(eN) = b(0): eY2(eN) = b(1)
            eN = eN + 1
            Exit Sub
        End If
        If Abs(axv(2)) < 0.9 Then Exit Sub            ' seen at an angle (isometric): not used
        Extend a: Extend b
        Extend Arr3(ctr(0) - rr, ctr(1) - rr, 0#)
        Extend Arr3(ctr(0) + rr, ctr(1) + rr, 0#)
        Set eObj(eN) = ent: eKind(eN) = "C"
        eCX(eN) = ctr(0): eCY(eN) = ctr(1): eR(eN) = rr
        eX1(eN) = a(0): eY1(eN) = a(1): eX2(eN) = b(0): eY2(eN) = b(1)
        eFull(eN) = (Abs(a(0) - b(0)) + Abs(a(1) - b(1)) < 0.0005)
        eN = eN + 1
    End If
End Sub

' =====================================================================================
'  Title block, finish, save  (from HW#8C)
' =====================================================================================
Sub TitleBlock()
    Dim pass As Integer, sv As Object, nt As Object, t As String, pos As Variant
    Dim mX As Double, mY As Double, dX As Double, dY As Double, tX As Double, tY As Double
    Dim gotM As Boolean, gotD As Boolean, gotT As Boolean, sheetDone As Boolean, sheetSeen As Boolean, scaleDone As Boolean
    Dim twoOK As Boolean, threeOK As Boolean, tolSeen As Boolean

    For pass = 1 To 2
        On Error Resume Next
        If pass = 2 Then swDraw.EditTemplate
        Set sv = swDraw.GetFirstView                        ' the sheet itself
        Set nt = sv.GetFirstNote
        Do While Not nt Is Nothing
            t = UCase(Trim(nt.GetText))
            pos = nt.GetAnnotation.GetPosition
            If IsArray(pos) Then
                If Left(t, 9) = "MATERIAL:" And Not gotM Then mX = pos(0) / IN_M: mY = pos(1) / IN_M: gotM = True
                If Left(t, 5) = "DATE:" And Not gotD Then dX = pos(0) / IN_M: dY = pos(1) / IN_M: gotD = True
                If Left(t, 6) = "TITLE:" And Not gotT Then tX = pos(0) / IN_M: tY = pos(1) / IN_M: gotT = True
            End If
            If Left(t, 5) = "SHEET" And InStr(t, " OF ") > 0 And Not sheetDone Then
                sheetSeen = True
                nt.SetText "SHEET 4 OF 4"
                If UCase(Trim(nt.GetText)) <> "SHEET 4 OF 4" Then nt.PropertyLinkedText = "SHEET 4 OF 4"
                sheetDone = (UCase(Trim(nt.GetText)) = "SHEET 4 OF 4")
            End If
            If Left(t, 6) = "SCALE:" And Not scaleDone Then              ' as on the reference
                If InStr(t, "1:1") = 0 Then nt.SetText "SCALE: 1:1"
                If InStr(UCase(nt.GetText), "1:1") = 0 Then nt.PropertyLinkedText = "SCALE: 1:1"
                scaleDone = (InStr(UCase(nt.GetText), "1:1") > 0)
            End If
            If InStr(t, "PLACE DECIMAL") > 0 Then                         ' note 2: tolerances as on the reference
                tolSeen = True
                FixTol nt, twoOK, threeOK
            End If
            Set nt = nt.GetNext
        Loop
        If pass = 2 Then swDraw.EditSheet
        On Error GoTo 0
        If gotM And gotD And gotT And sheetDone And scaleDone And twoOK And threeOK Then Exit For
    Next
    On Error Resume Next
    swDraw.EditSheet                                    ' make sure we are NOT left in Edit Sheet Format mode
    On Error GoTo 0

    If Not gotM Then mX = 5.07: mY = 0.94
    If Not gotD Then dX = 3.17: dY = 0.48
    If Not gotT Then tX = 7.2: tY = 0.92
    Check gotM And gotD And gotT, "Title block labels found (MATERIAL, DATE, TITLE)" & _
          IIf(gotM And gotD And gotT, "", " - used the reference positions instead")

    AddSheetNote "6061-T6 ALUMINUM", mX + 0.06, mY - 0.17, "Material: 6061-T6 ALUMINUM"
    AddSheetNote Format(Date, "mm-dd-yyyy"), dX + 0.4, dY - 0.18, "Date: " & Format(Date, "mm-dd-yyyy")
    AddSheetNote "IME 1140 ENGINEERING DRAWINGS II", tX + 0.44, tY, "Title line 1"
    AddSheetNote "IN-LAB ASSIGNMENT #8D", tX + 0.44, tY - 0.17, "Title line 2"
    Check scaleDone, "SCALE: 1:1" & IIf(scaleDone, "", " - edit the scale note by hand")
    Check sheetDone, "SHEET 4 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
    Check (twoOK And threeOK), "Tolerance block reads TWO PLACE " & ChrW(177) & ".01, THREE PLACE " & ChrW(177) & ".005 (note 2, as on the reference)" & _
          IIf(twoOK And threeOK, "", IIf(tolSeen, " - edit it by hand (Edit Sheet Format)", " - tolerance note not found; check it by hand"))
End Sub

' make the TWO / THREE PLACE DECIMAL lines read +-.01 / +-.005 (the reference) and read them back
Sub FixTol(nt As Object, twoOK As Boolean, threeOK As Boolean)
    Dim t As String, u As String
    On Error Resume Next
    t = nt.GetText
    If InStr(UCase(LineOf(UCase(t), "TWO PLACE DECIMAL")), ".01") = 0 Or InStr(UCase(LineOf(UCase(t), "THREE PLACE DECIMAL")), ".005") = 0 Then
        t = NewLine(t, "TWO PLACE DECIMAL", "TWO PLACE DECIMAL <MOD-PM> .01")
        t = NewLine(t, "THREE PLACE DECIMAL", "THREE PLACE DECIMAL  <MOD-PM> .005")
        nt.SetText t
    End If
    u = UCase(nt.GetText)
    If InStr(u, "TWO PLACE DECIMAL") > 0 Then twoOK = (InStr(LineOf(u, "TWO PLACE DECIMAL"), ".01") > 0 And InStr(LineOf(u, "TWO PLACE DECIMAL"), ".02") = 0)
    If InStr(u, "THREE PLACE DECIMAL") > 0 Then threeOK = (InStr(LineOf(u, "THREE PLACE DECIMAL"), ".005") > 0)
End Sub

' replace the whole line that contains key (case-insensitive) with newText
Function NewLine(ByVal t As String, ByVal key As String, ByVal newText As String) As String
    Dim p As Long, e As Long, e2 As Long
    p = InStr(UCase(t), UCase(key))
    If p = 0 Then NewLine = t: Exit Function
    e = InStr(p, t, vbCr): e2 = InStr(p, t, vbLf)
    If e = 0 Or (e2 > 0 And e2 < e) Then e = e2
    If e = 0 Then e = Len(t) + 1
    NewLine = Left(t, p - 1) & newText & Mid(t, e)
End Function

Function LineOf(ByVal t As String, ByVal key As String) As String
    Dim p As Long, e As Long, e2 As Long
    p = InStr(t, key)
    If p = 0 Then Exit Function
    e = InStr(p, t, vbCr): e2 = InStr(p, t, vbLf)
    If e = 0 Or (e2 > 0 And e2 < e) Then e = e2
    If e = 0 Then e = Len(t) + 1
    LineOf = Mid(t, p, e - p)
End Function

Sub AddSheetNote(ByVal text As String, ByVal x As Double, ByVal y As Double, ByVal label As String)
    Dim nt As Object
    On Error Resume Next
    swDraw.ActivateSheet swDraw.GetCurrentSheet.GetName
    swModel.ClearSelection2 True
    Set nt = swModel.InsertNote(text)
    If Not nt Is Nothing Then nt.GetAnnotation.SetPosition2 x * IN_M, y * IN_M, 0#
    On Error GoTo 0
    Check Not nt Is Nothing, label
End Sub

Sub FinishSheet()
    ' leave the drawing active, in normal sheet mode, with every view shown and redrawn
    Dim vs As Variant, n As Long, i As Long, v As Variant
    On Error Resume Next
    ActivateDrawing
    swDraw.EditSheet
    swModel.ClearSelection2 True
    For Each v In Array(viewMain, viewBB, viewAA, viewIso)
        If Not v Is Nothing Then
            If Not v.GetVisible Then v.SetVisible True, True
        End If
    Next
    swModel.ForceRebuild3 False
    swModel.GraphicsRedraw2
    swModel.ViewZoomtofit2
    vs = swDraw.GetCurrentSheet.GetViews
    n = -1
    If IsArray(vs) Then n = UBound(vs) - LBound(vs) + 1
    On Error GoTo 0
    If n >= 0 Then Check (n >= 4), n & " views on the sheet (expected 4)"
End Sub

Sub SaveAll()
    Dim lastName As String, folder As String, drw As String, pdf As String, ok As Boolean, okPdf As Boolean
    lastName = Trim(InputBox("Last name for the file names  " & DRAW_NAME & "<LastName>.SLDDRW  and  .PDF" & vbCrLf & _
                             "(leave empty to skip saving)", "Save " & DRAW_NAME, "Perhach"))
    If Len(lastName) = 0 Then WriteReport SaveFolder() & "\": Exit Sub
    folder = Left(partPath, InStrRev(partPath, "\"))
    drw = folder & DRAW_NAME & lastName & ".SLDDRW"
    ok = TrySave(drw)
    If Not ok Then
        folder = SaveFolder() & "\"
        drw = folder & DRAW_NAME & lastName & ".SLDDRW"
        ok = TrySave(drw)
    End If
    If Not ok Then
        MsgBox "Could not save the drawing automatically. Use File > Save As and name it  " & DRAW_NAME & lastName & _
               ", then File > Save As > PDF with the same name.", vbExclamation, DRAW_NAME
        Exit Sub
    End If
    pdf = folder & DRAW_NAME & lastName & ".PDF"
    FinishSheet                                         ' views visible + drawing active before the PDF export
    okPdf = TrySavePdf(pdf)
    WriteReport folder
    MsgBox "Saved:" & vbCrLf & drw & vbCrLf & IIf(okPdf, pdf, "PDF NOT saved - use File > Save As > PDF"), _
           IIf(okPdf, vbInformation, vbExclamation), DRAW_NAME
End Sub

Sub WriteReport(ByVal folder As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open folder & FILE_TAG & "_checks.txt" For Output As #f
    Print #f, DRAW_NAME & " drawing checks  (" & Now & ")"
    Print #f, "Passed: " & nOK & "   Problems: " & nBad & "   To do by hand: " & nTodo
    Print #f, ""
    Print #f, report
    Close #f
End Sub

Function TrySave(ByVal path As String) As Boolean
    Dim errs As Long, warns As Long, ok As Boolean
    On Error Resume Next
    ok = swModel.Extension.SaveAs(path, 0, 1, Nothing, errs, warns)       ' current version, silent
    If Err.Number <> 0 Or Not ok Then Err.Clear: ok = (swModel.SaveAs3(path, 0, 1) = 0)
    TrySave = ok And FileExists(path)
End Function

Function TrySavePdf(ByVal path As String) As Boolean
    Dim errs As Long, warns As Long, ok As Boolean, ex As Object
    On Error Resume Next
    ok = swModel.Extension.SaveAs(path, 0, 1, Nothing, errs, warns)
    If Err.Number <> 0 Or Not ok Then
        Err.Clear
        Set ex = swApp.GetExportFileData(1)                                ' 1 = swExportPdfData
        ok = swModel.Extension.SaveAs(path, 0, 1, ex, errs, warns)
    End If
    TrySavePdf = ok And FileExists(path)
End Function

Function SaveFolder() As String
    Dim sh As Object, p As String
    On Error Resume Next
    Set sh = CreateObject("WScript.Shell")
    p = sh.SpecialFolders("Desktop")
    If Len(p) = 0 Or Len(Dir(p, vbDirectory)) = 0 Then p = sh.SpecialFolders("MyDocuments")
    If Len(p) = 0 Or Len(Dir(p, vbDirectory)) = 0 Then p = Environ("USERPROFILE")
    SaveFolder = p
End Function

Sub Check(ByVal ok As Boolean, ByVal label As String)
    If ok Then
        nOK = nOK + 1
        report = report & "OK       " & label & vbCrLf
    Else
        nBad = nBad + 1
        report = report & "PROBLEM  " & label & vbCrLf
        problems = problems & "- " & label & vbCrLf
    End If
End Sub

' =====================================================================================
'  Geometry helpers  (from HW#8C)
' =====================================================================================
' model point (metres) -> sheet point (inches)
Function XfPt(xf As Object, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Variant
    Dim a(2) As Double, v As Variant, p As Object, r As Variant, o(2) As Double, k As Integer
    a(0) = x: a(1) = y: a(2) = z
    v = a
    On Error Resume Next
    Set p = swMath.CreatePoint((v))                 ' double parentheses = pass by value (required)
    If Not p Is Nothing Then Set p = p.MultiplyTransform(xf)
    If Not p Is Nothing Then r = p.ArrayData
    On Error GoTo 0
    If IsArray(r) Then
        If Abs(r(0)) + Abs(r(1)) + Abs(r(2)) > 0 Or Abs(x) + Abs(y) + Abs(z) = 0 Then
            For k = 0 To 2: o(k) = r(k) / IN_M: Next
            XfPt = o
            Exit Function
        End If
    End If
    r = ManualXf(xf, v, True)
    For k = 0 To 2: o(k) = r(k) / IN_M: Next
    XfPt = o
End Function

' model direction -> sheet direction (unit vector)
Function XfVec(xf As Object, vec As Variant) As Variant
    Dim a(2) As Double, v As Variant, p As Object, r As Variant
    a(0) = vec(0): a(1) = vec(1): a(2) = vec(2)
    v = a
    On Error Resume Next
    Set p = swMath.CreateVector((v))
    If Not p Is Nothing Then Set p = p.MultiplyTransform(xf)
    If Not p Is Nothing Then r = p.ArrayData
    On Error GoTo 0
    If IsArray(r) Then
        If Abs(r(0)) + Abs(r(1)) + Abs(r(2)) > 0 Then XfVec = Unit(r): Exit Function
    End If
    XfVec = Unit(ManualXf(xf, v, False))
End Function

Function ManualXf(xf As Object, a As Variant, ByVal withTranslation As Boolean) As Variant
    Dim t As Variant, o(2) As Double, k As Integer
    t = xf.ArrayData                               ' 9 rotation terms, 3 translation, 1 scale
    For k = 0 To 2
        o(k) = (a(0) * t(k) + a(1) * t(3 + k) + a(2) * t(6 + k)) * t(12)
        If withTranslation Then o(k) = o(k) + t(9 + k)
    Next
    ManualXf = o
End Function

Function ViewDir(v As Object, vec As Variant) As Variant
    ViewDir = XfVec(v.ModelToViewTransform, vec)
End Function

Function Arr3(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Variant
    Dim o(2) As Double
    o(0) = x: o(1) = y: o(2) = z
    Arr3 = o
End Function

Function Unit(v As Variant) As Variant
    Dim L As Double
    L = Sqr(v(0) * v(0) + v(1) * v(1) + v(2) * v(2))
    If L < 1E-12 Then L = 1
    Unit = Arr3(v(0) / L, v(1) / L, v(2) / L)
End Function

Function Dot(a As Variant, b As Variant) As Double
    Dot = a(0) * b(0) + a(1) * b(1) + a(2) * b(2)
End Function

Function Cross(a As Variant, b As Variant) As Variant
    Cross = Arr3(a(1) * b(2) - a(2) * b(1), a(2) * b(0) - a(0) * b(2), a(0) * b(1) - a(1) * b(0))
End Function

Function Scl(a As Variant, ByVal s As Double) As Variant
    Scl = Arr3(a(0) * s, a(1) * s, a(2) * s)
End Function

Function Add3(a As Variant, b As Variant) As Variant
    Add3 = Arr3(a(0) + b(0), a(1) + b(1), a(2) + b(2))
End Function

Function Sub3(a As Variant, b As Variant) As Variant
    Sub3 = Arr3(a(0) - b(0), a(1) - b(1), a(2) - b(2))
End Function

Function Near(v As Variant, ByVal x As Double, ByVal y As Double, ByVal z As Double) As Boolean
    Near = (Abs(v(0) - x) < 0.01 And Abs(v(1) - y) < 0.01 And Abs(v(2) - z) < 0.01)
End Function

Function VecTxt(v As Variant) As String
    Dim n(2) As String, k As Integer
    For k = 0 To 2
        If Abs(v(k)) < 0.01 Then n(k) = "0" Else n(k) = Format(v(k), "0.##")
    Next
    VecTxt = "(" & n(0) & ", " & n(1) & ", " & n(2) & ")"
End Function
