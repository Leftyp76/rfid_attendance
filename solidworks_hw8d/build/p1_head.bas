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
