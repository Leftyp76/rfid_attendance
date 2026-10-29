' =====================================================================================
'  HW#8A  -  IME 1140 Engineering Drawings II, In-Lab Assignment #8A  -  DRAWING macro
'            (inches, sheet 1:1, third angle)
'
'  RUN THE ENGRAVE MACRO FIRST
'   HW8A_Engrave_Logo_Macro.bas sets the logo to .003 deep (note 3) and saves the part.
'   This drawing macro only checks that depth; it never changes or saves the part.
'
'  HOW TO RUN
'   1. Open "IME 1140 Engineering Drawings II In-Lab Assignment Part 1.SLDPRT" in SolidWorks.
'   2. Tools > Macro > New...  save it as  HW8A_Drawing.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main  (Tools > Macro > Run, select "main", Run).
'   5. At the end enter your last name; it saves  HW#8A<LastName>.SLDDRW  and  .PDF  next to
'      the part, plus HW8A_checks.txt with every check.
'
'  WHAT IT BUILDS (matches the reference sheet)
'   - Front view: pocket and logo toward you, counterbore hole at the top, hidden lines removed
'   - SECTION A-A to the right of it, cut down the vertical centre line, level with the front view
'   - Isometric view top right, shaded with edges, scale 1:3
'   - Front:   3.875+-.003, (.6875), dia 2.50 bolt circle, 6 X 60 deg, 16 X R.375, 4X .75 X 45 deg,
'              hole callouts: dia .257 cbore dia .44 deep .25 / dia .257 deep .10 / 1/4-20 UNC /
'              dia .257 THRU / dia .257 deep .35 / dia .257 csk dia .50 X 90 deg, centre marks
'   - Section: 3.875+-.003, 3.630/3.620, dia 3.50 +.01/-.00, .480+-.005, .085/.075, .250,
'              .025 X 45 deg TYP, datum A on the back face, parallelism .002 to A, hole centrelines
'   - Title block: 6061-T6 ALUMINUM, today's date, IME 1140 ENGINEERING DRAWINGS II /
'     IN-LAB ASSIGNMENT #8A, SHEET 1 OF 4, SCALE 1:1, tolerance block +-.02 and +-.008 (note 2)
'   - No NOTES block (left out as asked).  The SECTION A-A label is hidden, as on the reference.
'
'  ALIGNMENT
'   - Every view is moved by a real part point to its measured position, to within .001 in.
'   - SECTION A-A stays projected from the front view; the report checks both centres are level.
'   - The dimensions above the section are stacked .24 in apart, with their text in one column.
'
'  v2: views forced to 1:1 (SolidWorks had auto-scaled them to 1:2); bolt circle and hole circles
'      scale-aware; hole depths measured exactly from their edges, blind holes told apart by depth;
'      tolerance values tried several ways until they read back; counterbore <HOLE-SPOT>,
'      countersink <HOLE-SINK>, parallelism <IGTOL-PARA>; every hole listed in the check file
'  v1: first version
'
'  SAFETY
'   - Orientation read from the part: the dia 3.50 pocket faces you, the counterbore hole is up.
'   - The six holes are recognised by shape (counterbore, countersink, depth) and checked
'     against the reference layout; each callout goes on the hole that really has that shape.
'   - Every view is verified, every dimension is read back and checked, and the sketched
'     cutting line and bolt circle are measured from the part edges before they are used.
'   - The part file is never saved or changed by this macro.
' =====================================================================================
Option Explicit

Dim swApp As Object, swMath As Object, swPart As Object, swModel As Object, swDraw As Object
Dim partPath As String, curStep As String, report As String, problems As String
Dim nOK As Long, nBad As Long
Dim lastDim As Object                      ' the dimension TryDimSel made last (tolerances go on it)
Dim oldInputDim As Boolean, savedInputDim As Boolean
Dim sketchCand As Integer                  ' how the front view's sketch coordinates work (-1 = not proven yet)
Dim segBolt As Object, segRay1 As Object, segRay2 As Object

Const IN_M = 0.0254
Const DRAW_NAME = "HW#8A"
Const FILE_TAG = "HW8A"
Const SHEET_DEN = 1#             ' sheet scale 1 : 1
Const EMAX = 8000                ' edge table size (the engraved logo adds thousands of edges)
Const PLACE_TOL = 0.001          ' views are placed to within .001 in
Const DEG = 0.0174532925199433   ' radians per degree

' ---- sheet layout (sheet inches, measured from the reference: 108 px per inch) ----
Const FRONT_X = 3.86             ' part centre in the front view
Const FRONT_Y = 3.85
Const SEC_X = 9.65               ' back face (datum A) in SECTION A-A
Const ISO_X = 9.92               ' part centre in the isometric view
Const ISO_Y = 7.47
Const ISO_DEN = 3#               ' isometric view scale 1 : 3

Const DM_HIDDEN_VISIBLE = 1
Const DM_HIDDEN_REMOVED = 2
Const DM_SHADED = 3

Const TOL_BILAT = 2              ' swTolType_e
Const TOL_LIMIT = 3
Const TOL_SYM = 4

' ---- part sizes (inches) ----
Const HALF = 1.9375              ' half of 3.875 (outside of the plate)
Const PAD = 1.8125               ' half of 3.625 (raised pad)
Const PKT = 1.75                 ' pocket radius (dia 3.50)
Const BC = 1.25                  ' bolt circle radius (dia 2.50)
Const R257 = 0.1285              ' hole radius (dia .257)
Const F_RIM = 0.25               ' rim front face, measured from the back face
Const F_FLOOR = 0.4              ' pocket floor
Const F_PAD = 0.48               ' pad front face

' ---- the part's directions (model coordinates, unit vectors) ----
Dim vR As Variant                ' right in the front view
Dim vU As Variant                ' up in the front view (centre -> counterbore hole)
Dim vF As Variant                ' toward you in the front view (back face -> pocket)
Dim pOrg As Variant              ' model point: part centre on the back face (datum A)

' ---- the six holes on the bolt circle ----
Dim hN As Long
Dim hD(5) As Variant             ' offset from the part centre (model vector, metres)
Dim hCb(5) As Boolean, hCsk(5) As Boolean, hSpan(5) As Double, hRmin(5) As Double
Dim hR(5) As Double, hU(5) As Double, hKind(5) As String

Dim viewFront As Object, viewSec As Object, viewIso As Object

Dim eN As Long
Dim eObj() As Object, eKind() As String
Dim eX1() As Double, eY1() As Double, eX2() As Double, eY2() As Double
Dim eCX() As Double, eCY() As Double, eR() As Double, eFull() As Boolean
Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double

Sub main()
    On Error GoTo Fail
    Dim tmpl As String
    Set swApp = Application.SldWorks
    Set swMath = swApp.GetMathUtility
    report = "": problems = "": nOK = 0: nBad = 0: sketchCand = -1
    Set viewFront = Nothing: Set viewSec = Nothing: Set viewIso = Nothing      ' VBA keeps these between runs
    Set segBolt = Nothing: Set segRay1 = Nothing: Set segRay2 = Nothing: Set lastDim = Nothing
    NoDimDialog True

    curStep = "Step 1 - find the open part":               FindPart
    curStep = "Step 2 - read the part and its holes":      FindOrientation
    curStep = "Step 2 - logo depth":                       CheckLogo
    curStep = "Step 3 - find your template":               tmpl = FindTemplate()
    curStep = "Step 4 - new drawing (scale 1:1)":          NewDrawing tmpl
    curStep = "Step 5 - front view":                       MakeFront
    curStep = "Step 6 - SECTION A-A":                      MakeSectionAA
    curStep = "Step 7 - isometric view":                   MakeIso
    curStep = "Step 8 - centre marks and bolt circle":     ClearCommand: CentreMarks
    curStep = "Step 9 - front view dimensions":            ClearCommand: DimFront
    curStep = "Step 10 - SECTION A-A dimensions":          ClearCommand: DimSection
    curStep = "Step 11 - datum A and parallelism":         ClearCommand: GtolAndDatum
    curStep = "Step 12 - section label":                   HideSectionLabel
    curStep = "Step 13 - title block":                     ClearCommand: TitleBlock
    curStep = "Step 14 - centrelines":                     CentreLines viewSec: ClearCommand
    curStep = "Step 15 - alignment check":                 CheckAlignment
    curStep = "Step 16 - finish":                          FinishSheet
    NoDimDialog False

    MsgBox DRAW_NAME & " drawing built." & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & vbCrLf & vbCrLf & _
           IIf(nBad = 0, "Every view, dimension and title block item checked out.", "Problems:" & vbCrLf & Left(problems, 700)) & _
           vbCrLf & vbCrLf & "The full check list is saved as " & FILE_TAG & "_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), DRAW_NAME

    curStep = "Step 17 - save drawing and PDF":            SaveAll
    Exit Sub
Fail:
    NoDimDialog False
    MsgBox "The macro stopped at: " & curStep & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Checks passed so far: " & nOK & vbCrLf & _
           IIf(nBad = 0, "", "Problems so far:" & vbCrLf & Left(problems, 500)), vbCritical, DRAW_NAME & " macro"
End Sub

' Dimensions on sketch entities (bolt circle, radial lines) would pop up the Modify box;
' "Input dimension value" is switched off while the macro runs and put back afterwards.
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

' leave any half-finished command and selection, back to plain sheet editing
Sub ClearCommand()
    On Error Resume Next
    swModel.ClearSelection2 True
    swDraw.ActivateSheet swDraw.GetCurrentSheet.GetName
    swModel.ClearSelection2 True
    swModel.GraphicsRedraw2
End Sub

' =====================================================================================
'  Part, orientation, holes, logo
' =====================================================================================
Function PartMatches(ByVal t As String) As Boolean
    t = UCase(Replace(Replace(t, "_", " "), "#", ""))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    PartMatches = (InStr(t, "DRAWINGS II") > 0 And InStr(t, "PART 1") > 0) Or InStr(t, "IME COASTER") > 0
End Function

Sub FindPart()
    ' Use the open part named "...Drawings II In-Lab Assignment Part 1". Otherwise the active part,
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
        "Could not tell which open part is the In-Lab #8A part. Close the other parts, click into its window, and run again."
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then Err.Raise vbObjectError + 2, , "The part has never been saved. Save it first."
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

Sub FindOrientation()
    ' dia 3.50 pocket: its axis is front-to-back and passes through the part centre; the pocket is
    ' on the front.  Counterbore hole (dia .44): straight up from the centre.
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, k As Long, g As Long
    Dim s As Object, cp As Variant, bx As Variant, ax As Variant, r As Double
    Dim p0 As Variant, pMid As Variant, gotP As Boolean, bb As Variant, sp As Variant, e1 As Variant
    Dim fMin As Double, wR As Double, wU As Double, wF As Double, cb As Long, w As Double

    bodies = swPart.GetBodies2(0, True)
    If Not IsArray(bodies) Then Err.Raise vbObjectError + 3, , "The part has no solid body."
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder And Not gotP Then
                cp = s.CylinderParams
                If Abs(cp(6) / IN_M - PKT) < 0.004 Then
                    ax = Unit(Arr3(cp(3), cp(4), cp(5)))
                    p0 = Arr3(cp(0), cp(1), cp(2))
                    bx = faces(i).GetBox
                    pMid = Arr3((bx(0) + bx(3)) / 2, (bx(1) + bx(4)) / 2, (bx(2) + bx(5)) / 2)
                    gotP = True
                End If
            End If
        Next
    Next
    If Not gotP Then Err.Raise vbObjectError + 4, , _
        "This part has no dia 3.50 pocket, so it is not the In-Lab #8A part. Click into the right part window and run again."
    bb = bodies(LBound(bodies)).GetBodyBox
    vF = ax                                                         ' the pocket is on the front
    If Dot(Sub3(pMid, BoxMid(bb)), vF) < 0 Then vF = Scl(vF, -1)

    ' ---- the holes on the dia 2.50 bolt circle, grouped by position ----
    hN = 0
    For k = 0 To 5
        hD(k) = Empty: hCb(k) = False: hCsk(k) = False: hSpan(k) = 0: hKind(k) = "": hRmin(k) = 9#
    Next
    If Abs(vF(0)) < 0.9 Then e1 = Unit(Cross(vF, Arr3(1, 0, 0))) Else e1 = Unit(Cross(vF, Arr3(0, 1, 0)))
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder Then
                cp = s.CylinderParams
                r = cp(6) / IN_M
                If r < 0.3 And Abs(Dot(Unit(Arr3(cp(3), cp(4), cp(5))), vF)) > 0.999 Then
                    g = HoleGroup(InPlane(Arr3(cp(0), cp(1), cp(2)), p0))
                    If g >= 0 Then
                        If r < hRmin(g) Then hRmin(g) = r
                        If Abs(r - 0.22) < 0.004 Then hCb(g) = True
                        If Abs(r - R257) < 0.003 Then
                            w = CylSpan(faces(i))                          ' exact, from its circular edges
                            If w < 0 Then sp = SpanAlong(faces(i).GetBox, vF): w = (sp(1) - sp(0)) / IN_M
                            If w > hSpan(g) Then hSpan(g) = w
                        End If
                    End If
                End If
            ElseIf s.IsCone Then
                bx = faces(i).GetBox
                g = HoleGroup(InPlane(Arr3((bx(0) + bx(3)) / 2, (bx(1) + bx(4)) / 2, (bx(2) + bx(5)) / 2), p0))
                sp = SpanAlong(bx, e1)
                If g >= 0 And (sp(1) - sp(0)) / IN_M > 0.4 Then hCsk(g) = True       ' a drill point is only .257 wide
            End If
        Next
    Next
    cb = -1
    For k = 0 To hN - 1
        If hCb(k) And cb < 0 Then cb = k
    Next
    If cb < 0 Then Err.Raise vbObjectError + 5, , "Could not find the counterbore hole (dia .44 on the dia 2.50 bolt circle)."
    vU = Unit(hD(cb))
    vR = Cross(vU, vF)

    ' ---- origin: part centre on the back face; size check ----
    sp = SpanAlong(bb, vF): fMin = sp(0): wF = (sp(1) - sp(0)) / IN_M
    sp = SpanAlong(bb, vR): wR = (sp(1) - sp(0)) / IN_M
    sp = SpanAlong(bb, vU): wU = (sp(1) - sp(0)) / IN_M
    pOrg = Sub3(p0, Scl(vF, Dot(p0, vF) - fMin))
    Check (Abs(wR - 3.875) < 0.01 And Abs(wU - 3.875) < 0.01 And Abs(wF - 0.48) < 0.005), _
          "Part size 3.875 x 3.875 x .480 (measured " & Format(wR, "0.000") & " x " & Format(wU, "0.000") & " x " & Format(wF, "0.000") & ")"

    ' ---- name each hole by its shape ----
    Dim b1 As Long, b2 As Long
    b1 = -1: b2 = -1
    For k = 0 To hN - 1
        hR(k) = Dot(hD(k), vR) / IN_M: hU(k) = Dot(hD(k), vU) / IN_M
        If hCb(k) Then
            hKind(k) = "CBORE"
        ElseIf hCsk(k) Then
            hKind(k) = "CSK"
        ElseIf hSpan(k) <= 0 Then
            hKind(k) = "TAP"                                   ' no dia .257 bore: the tap drill hole
        ElseIf hSpan(k) > 0.39 Then
            hKind(k) = "THRU"
        Else
            hKind(k) = "BLIND"
            If b1 < 0 Then b1 = k Else b2 = k
        End If
    Next
    ' the two blind holes: the shallower one is the .10 deep, the deeper one the .35 deep
    If b1 >= 0 And b2 >= 0 Then
        If hSpan(b1) > hSpan(b2) Then k = b1: b1 = b2: b2 = k
        hKind(b1) = "D10": hKind(b2) = "D35"
    ElseIf b1 >= 0 Then
        If Abs(hSpan(b1) - 0.1) < Abs(hSpan(b1) - 0.35) Then hKind(b1) = "D10" Else hKind(b1) = "D35"
    End If
    For k = 0 To hN - 1
        Info "Hole at " & Format(AngDeg(hR(k), hU(k)), "0") & " deg: " & hKind(k) & "  (smallest bore " & ChrW(216) & _
             Format(2 * hRmin(k), "0.000") & IIf(hSpan(k) > 0, ", " & ChrW(216) & ".257 bore " & Format(hSpan(k), "0.000") & " long", "") & ")"
    Next
    Check (hN = 6), hN & " holes found on the " & ChrW(216) & "2.50 bolt circle (reference: 6)"
    HoleLayout
End Sub

' the reference layout: counterbore 90 deg (top), THRU 30, .10 deep 150, 1/4-20 210, countersink 270, .35 deep 330
Sub HoleLayout()
    Dim kinds As Variant, angs As Variant, i As Long, k As Long, a As Double, dA As Double, msg As String
    kinds = Array("CBORE", "THRU", "D10", "TAP", "CSK", "D35")
    angs = Array(90#, 30#, 150#, 210#, 270#, 330#)
    For i = 0 To 5
        k = HoleIdx(CStr(kinds(i)))
        If k < 0 Then
            msg = msg & " " & kinds(i) & " hole not found;"
        Else
            a = AngDeg(hR(k), hU(k))
            dA = a - angs(i)
            Do While dA > 180#: dA = dA - 360#: Loop
            Do While dA < -180#: dA = dA + 360#: Loop
            If Abs(dA) > 2# Then msg = msg & " " & kinds(i) & " at " & Format(a, "0") & " deg (reference " & angs(i) & ");"
        End If
    Next
    Check (Len(msg) = 0), "Hole layout matches the reference (cbore top, csk bottom, .10 upper left, 1/4-20 lower left, THRU upper right, .35 lower right)" & _
          IIf(Len(msg) = 0, "", " - DIFFERENT:" & msg & " callouts still go on the matching holes")
End Sub

' which hole group a bolt-circle position belongs to (adds a new group); -1 if not on the bolt circle
Function HoleGroup(d As Variant) As Long
    Dim k As Long, L As Double
    HoleGroup = -1
    L = Sqr(Dot(d, d)) / IN_M
    If Abs(L - BC) > 0.03 Then Exit Function
    For k = 0 To hN - 1
        If Sqr(Dot(Sub3(d, hD(k)), Sub3(d, hD(k)))) / IN_M < 0.02 Then HoleGroup = k: Exit Function
    Next
    If hN < 6 Then
        hD(hN) = d
        HoleGroup = hN
        hN = hN + 1
    End If
End Function

' offset of model point q from the pocket axis point p0, in the plane of the plate
Function InPlane(q As Variant, p0 As Variant) As Variant
    Dim d As Variant
    d = Sub3(q, p0)
    InPlane = Sub3(d, Scl(vF, Dot(d, vF)))
End Function

Function HoleIdx(ByVal kind As String) As Long
    Dim k As Long
    HoleIdx = -1
    For k = 0 To hN - 1
        If hKind(k) = kind Then HoleIdx = k: Exit Function
    Next
End Function

' a point on hole k's edge circle of radius rad, in direction (dx, dy) from its centre (part inches)
Function HolePt(ByVal k As Long, ByVal rad As Double, ByVal dx As Double, ByVal dy As Double) As Variant
    HolePt = Array(hR(k) + rad * dx, hU(k) + rad * dy, F_FLOOR)
End Function

Function AngDeg(ByVal x As Double, ByVal y As Double) As Double
    Dim a As Double
    If Abs(x) < 0.000000001 Then
        If y >= 0 Then a = 90# Else a = 270#
    Else
        a = Atn(y / x) / DEG
        If x < 0 Then a = a + 180#
        If a < 0 Then a = a + 360#
    End If
    AngDeg = a
End Function

Function BoxMid(bx As Variant) As Variant
    BoxMid = Arr3((bx(0) + bx(3)) / 2, (bx(1) + bx(4)) / 2, (bx(2) + bx(5)) / 2)
End Function

' length of a cylinder face along vF, from the centres of its circular edges (exact); -1 if unknown
Function CylSpan(fc As Object) As Double
    Dim eds As Variant, j As Long, c As Object, cc As Variant, t As Double, lo As Double, hi As Double, n As Long
    CylSpan = -1
    lo = 1E+30: hi = -1E+30
    On Error Resume Next
    eds = fc.GetEdges
    If Not IsArray(eds) Then Exit Function
    For j = LBound(eds) To UBound(eds)
        Set c = Nothing
        Set c = eds(j).GetCurve
        If Not c Is Nothing Then
            If c.IsCircle Then
                cc = Empty
                cc = c.CircleParams
                If IsArray(cc) Then
                    t = Dot(Arr3(cc(0), cc(1), cc(2)), vF)
                    If t < lo Then lo = t
                    If t > hi Then hi = t
                    n = n + 1
                End If
            End If
        End If
    Next
    If n >= 2 Then CylSpan = (hi - lo) / IN_M
End Function

' a line in the check file that is information only (not counted as OK or PROBLEM)
Sub Info(ByVal label As String)
    report = report & "INFO     " & label & vbCrLf
End Sub

' smallest and largest value of a box's 8 corners along direction "along" (metres)
Function SpanAlong(bx As Variant, along As Variant) As Variant
    Dim ix As Integer, iy As Integer, iz As Integer, t As Double, lo As Double, hi As Double
    lo = 1E+30: hi = -1E+30
    For ix = 0 To 1
        For iy = 0 To 1
            For iz = 0 To 1
                t = Dot(Arr3(bx(3 * ix), bx(1 + 3 * iy), bx(2 + 3 * iz)), along)
                If t < lo Then lo = t
                If t > hi Then hi = t
            Next
        Next
    Next
    SpanAlong = Array(lo, hi)
End Function

' note 3: the logo's floor faces must sit .003 below the pocket floor
Sub CheckLogo()
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, s As Object, pp As Variant, n As Variant
    Dim f As Double, dMin As Double, dMax As Double, cnt As Long, f0 As Double
    f0 = Dot(pOrg, vF)
    dMin = 1#: dMax = 0#
    bodies = swPart.GetBodies2(0, True)
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsPlane Then
                pp = s.PlaneParams
                n = Unit(Arr3(pp(0), pp(1), pp(2)))
                If Abs(Dot(n, vF)) > 0.999 Then
                    f = (Dot(Arr3(pp(3), pp(4), pp(5)), vF) - f0) / IN_M
                    If f > F_FLOOR - 0.02 And f < F_FLOOR - 0.0005 Then
                        cnt = cnt + 1
                        If F_FLOOR - f < dMin Then dMin = F_FLOOR - f
                        If F_FLOOR - f > dMax Then dMax = F_FLOOR - f
                    End If
                End If
            End If
        Next
    Next
    If cnt = 0 Then
        Check False, "No engraved logo found on the pocket floor - run HW8A_Engrave_Logo_Macro.bas first (note 3)"
    ElseIf Abs(dMin - 0.003) < 0.0003 And Abs(dMax - 0.003) < 0.0003 Then
        Check True, "Logo engraved .003 deep (" & cnt & " engraved faces) - note 3"
    Else
        Check False, "Logo is engraved " & Format(dMax, ".0000") & " deep, note 3 asks .003 - run HW8A_Engrave_Logo_Macro.bas, then this macro again"
    End If
End Sub

' part coordinates (inches): r right, u up (from the part centre), f toward you from the back face
Function PartPt(ByVal r As Double, ByVal u As Double, ByVal f As Double) As Variant
    PartPt = Add3(pOrg, Add3(Add3(Scl(vR, r * IN_M), Scl(vU, u * IN_M)), Scl(vF, f * IN_M)))
End Function

Function SheetPt(v As Object, ByVal r As Double, ByVal u As Double, ByVal f As Double) As Variant
    Dim p As Variant
    p = PartPt(r, u, f)
    SheetPt = XfPt(v.ModelToViewTransform, p(0), p(1), p(2))
End Function

Function SX(v As Object, ByVal r As Double, ByVal u As Double, ByVal f As Double) As Double
    SX = SheetPt(v, r, u, f)(0)
End Function

Function SY(v As Object, ByVal r As Double, ByVal u As Double, ByVal f As Double) As Double
    SY = SheetPt(v, r, u, f)(1)
End Function

' =====================================================================================
'  Template, new drawing  (from HW#7C)
' =====================================================================================
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
Sub MakeFront()
    Dim sc As Double
    Set viewFront = OrientedView("HW8A_Front", vR, vU, vF, FRONT_X, FRONT_Y, StdFrontName(), False)
    If viewFront Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented front view."
    Check True, "Front view: pocket and logo toward you, counterbore hole at the top"
    Check SetViewScale(viewFront, 1#), "Front view scale 1:1 (" & Format(ViewSc(viewFront), "0.###") & ")"
    SetMode viewFront, DM_HIDDEN_REMOVED, "Front"
    PlaceView viewFront, 0#, 0#, 0#, FRONT_X, FRONT_Y, True, True, _
              "Front view placed (part centre at " & Format(FRONT_X, "0.00") & ", " & Format(FRONT_Y, "0.00") & ")"
End Sub

' cutting line A-A down the vertical centre line, then the section to the right of the front view
Sub MakeSectionAA()
    Dim a As Variant, b As Variant, seg As Object, cand As Variant, opt As Integer, sv As Object, sec As Object
    Dim d As Double, tried As String, before As String, ok As Boolean

    CollectEdges viewFront
    a = SheetPt(viewFront, 0#, 2.15, F_PAD)
    b = SheetPt(viewFront, 0#, -2.15, F_PAD)
    ' The view sketch's coordinate system is not documented clearly, so each way is tried. A line is
    ' kept only when a test dimension measures it 1.9375 from the plate's left edge (the centre line).
    For Each cand In CandOrder()
        Set seg = DrawSeg(viewFront, a, b, CInt(cand))
        If seg Is Nothing Then
            tried = tried & "m" & cand & ": no line. "
        Else
            d = TestDist(seg, True, Array(-HALF, 0.3, F_RIM))
            If Abs(d - HALF) < 0.005 Then
                sketchCand = CInt(cand)
                Exit For
            End If
            tried = tried & "m" & cand & ": " & IIf(d < 0, "not measurable", "off by " & Format(Abs(d - HALF), "0.000")) & ". "
            DeleteSeg seg
            Set seg = Nothing
        End If
    Next
    If seg Is Nothing Then
        Check False, "Cutting line A-A: no method put it on the centre line (" & Trim(tried) & ") - check its position"
        Set seg = DrawSeg(viewFront, a, b, 2)                      ' as in HW#7B, without the position check
        If seg Is Nothing Then Err.Raise vbObjectError + 10, , "Could not draw the cutting line A-A."
    Else
        Check True, "Cutting line A-A on the vertical centre line (method " & sketchCand & ", measured 1.9375 from the left edge)"
    End If

    For opt = 0 To 1
        before = ViewNames()
        swModel.ClearSelection2 True
        On Error Resume Next
        seg.Select4 False, Nothing                                 ' select the line itself (no screen picking)
        Set sv = Nothing
        Set sv = swDraw.CreateSectionViewAt5((SEC_X - 0.24) * IN_M, FRONT_Y * IN_M, 0#, "A", opt, Nothing, 0#)
        If sv Is Nothing Then Err.Clear: Set sv = swDraw.CreateSectionViewAt4((SEC_X - 0.24) * IN_M, FRONT_Y * IN_M, 0#, "A", opt, Nothing)
        Err.Clear
        On Error GoTo 0
        If sv Is Nothing Then Set sv = NewViewSince(before)        ' built but not handed back
        If Not sv Is Nothing Then Exit For
    Next
    swModel.ClearSelection2 True
    If sv Is Nothing Then Err.Raise vbObjectError + 11, , "SolidWorks would not make SECTION A-A from the cutting line."

    ' looking left (arrows point left): the pocket (front face) shows on the left, up stays up
    If Not FrameOK(sv, vF, -1, 0, vU, 0, 1) Then
        On Error Resume Next
        Set sec = sv.GetSection
        sec.SetReversedCutDirection Not sec.GetReversedCutDirection
        swModel.ForceRebuild3 False
        On Error GoTo 0
    End If
    ok = FrameOK(sv, vF, -1, 0, vU, 0, 1)
    Check ok, "SECTION A-A: arrows point left, pocket on the left of the section" & IIf(ok, "", " - click the cutting line, Flip Direction")
    Set viewSec = sv
    Check SetViewScale(viewSec, 1#), "SECTION A-A scale 1:1 (" & Format(ViewSc(viewSec), "0.###") & ")"
    SetMode viewSec, DM_HIDDEN_REMOVED, "SECTION A-A"
    PlaceView viewSec, 0#, 0#, 0#, SEC_X, FRONT_Y, True, False, "SECTION A-A placed (back face at x = " & Format(SEC_X, "0.00") & ")"
End Sub

Sub MakeIso()
    Dim dv As Variant, upI As Variant, rtI As Variant, m As Long
    dv = Unit(Add3(Add3(vR, vU), vF))
    upI = Unit(Sub3(vU, Scl(dv, Dot(vU, dv))))
    rtI = Cross(upI, dv)
    Set viewIso = OrientedView("HW8A_Iso", rtI, upI, dv, ISO_X, ISO_Y, StdIsoName(), True)
    If viewIso Is Nothing Then
        Check False, "Isometric view - could not create it; add it by hand (View Palette > Isometric, scale 1:3)"
        Exit Sub
    End If
    Check True, "Isometric view"

    ' scale 1:3
    If SetViewScale(viewIso, ISO_DEN) Then
        Check True, "Isometric view scale 1:3"
    Else
        Check False, "Isometric view scale is " & Format(ViewSc(viewIso), "0.###") & " - set Scale > Use custom scale 1:3 by hand"
    End If

    ' shaded with edges
    On Error Resume Next
    viewIso.SetDisplayMode3 False, DM_SHADED, False, True
    m = -1: m = viewIso.GetDisplayMode2
    On Error GoTo 0
    Check (m = DM_SHADED), "Isometric view: shaded with edges" & IIf(m = DM_SHADED, "", " - pick Shaded With Edges by hand")
    PlaceView viewIso, 0#, 0#, F_PAD / 2, ISO_X, ISO_Y, True, True, "Isometric view placed (part centre)"
End Sub

' set view v to scale 1:den (SolidWorks may auto-scale new views); True when it reads back right
Function SetViewScale(v As Object, ByVal den As Double) As Boolean
    Dim sr(1) As Double, va As Variant
    If Abs(ViewSc(v) - 1# / den) > 0.0005 Then
        sr(0) = 1#: sr(1) = den: va = sr
        On Error Resume Next
        v.UseSheetScale = False
        v.ScaleRatio = va
        If Abs(ViewSc(v) - 1# / den) > 0.0005 Then Err.Clear: v.ScaleDecimal = 1# / den
        swModel.ForceRebuild3 False
        On Error GoTo 0
    End If
    SetViewScale = (Abs(ViewSc(v) - 1# / den) < 0.0005)
End Function

Function ViewSc(v As Object) As Double
    Dim sc As Double
    On Error Resume Next
    sc = 0
    sc = v.ScaleDecimal
    If sc <= 0 Then sc = 1#
    ViewSc = sc
End Function

' =====================================================================================
'  Sketching in the front view (cutting line, bolt circle, radial centre lines)
' =====================================================================================
Function CandOrder() As Variant
    If sketchCand >= 0 Then
        CandOrder = Array(sketchCand, 2, 3, 1, 0)
    Else
        CandOrder = Array(2, 3, 1, 0)
    End If
End Function

' sheet point (inches) -> view sketch point (metres).  cand 0: the view sketch's own transform;
' 1: sheet coordinates; 2: from the model origin; 3: from the view centre
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

Function DrawSeg(v As Object, a As Variant, b As Variant, ByVal cand As Integer) As Object
    Dim p As Variant, q As Variant
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    p = ToSk(v, a, cand): q = ToSk(v, b, cand)
    If Not IsArray(p) Or Not IsArray(q) Then Exit Function
    On Error Resume Next
    swModel.SketchManager.AddToDB = True
    Set DrawSeg = swModel.SketchManager.CreateLine(p(0), p(1), 0#, q(0), q(1), 0#)
    swModel.SketchManager.AddToDB = False
    swModel.GraphicsRedraw2
    swModel.ClearSelection2 True
End Function

' c = centre (sheet inches), rad = radius on the sheet (inches)
Function DrawCirc(v As Object, c As Variant, ByVal rad As Double, ByVal cand As Integer) As Object
    Dim p As Variant, sc As Double, rr As Double
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    p = ToSk(v, c, cand)
    If Not IsArray(p) Then Exit Function
    sc = 1#
    On Error Resume Next
    sc = v.ScaleDecimal
    If sc <= 0 Then sc = 1#
    If cand = 1 Then rr = rad * IN_M Else rr = rad * IN_M / sc
    swModel.SketchManager.AddToDB = True
    Set DrawCirc = swModel.SketchManager.CreateCircle(p(0), p(1), 0#, p(0) + rr, p(1), 0#)
    swModel.SketchManager.AddToDB = False
    swModel.GraphicsRedraw2
    swModel.ClearSelection2 True
End Function

' a test dimension from sketch item seg to the part edge through pEdge (front view); returns its
' value in inches (-1 if it could not be made).  The test dimension is deleted again.
Function TestDist(seg As Object, ByVal horiz As Boolean, pEdge As Variant) As Double
    Dim dd As Object, tp As Variant, n As Long
    TestDist = -1
    tp = SheetPt(viewFront, -1#, 2.7, 0#)                    ' text off the part, top left
    swModel.ClearSelection2 True
    swDraw.ActivateView viewFront.GetName2
    On Error Resume Next
    seg.Select4 False, Nothing
    n = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    On Error GoTo 0
    If n <> 1 Then swModel.ClearSelection2 True: Exit Function
    If Not SelPt(viewFront, pEdge, True) Then swModel.ClearSelection2 True: Exit Function
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) <> 2 Then swModel.ClearSelection2 True: Exit Function
    Set dd = Nothing
    On Error Resume Next
    If horiz Then
        Set dd = swModel.AddHorizontalDimension2(tp(0) * IN_M, tp(1) * IN_M, 0#)
    Else
        Set dd = swModel.AddVerticalDimension2(tp(0) * IN_M, tp(1) * IN_M, 0#)
    End If
    swModel.ClearSelection2 True
    If Not dd Is Nothing Then
        TestDist = dd.GetDimension2(0).SystemValue / IN_M
        dd.GetAnnotation.Select3 False, Nothing
        swModel.EditDelete
        swModel.ClearSelection2 True
    End If
    On Error GoTo 0
End Function

Sub CentreStyle(seg As Object)
    Dim ok As Boolean
    On Error Resume Next
    swModel.ClearSelection2 True
    seg.Select4 False, Nothing
    ok = swModel.SetLineStyle("Center")
    If Not ok Then ok = swModel.SetLineStyle("CENTER")
    swModel.ClearSelection2 True
End Sub

' =====================================================================================
'  Centre marks, bolt circle, radial centre lines
' =====================================================================================
Sub CentreMarks()
    Dim k As Long, idx As Long, n As Long, nMarks As Long, cm As Object, c As Variant, seg As Object
    Dim d1 As Double, d2 As Double, cand As Variant, tried As String, a As Variant, b As Variant, i As Integer
    Dim cs As Double, sn As Double

    CollectEdges viewFront
    ' 1. dia 3.50 pocket: centre mark with its lines extended (the long horizontal centre line)
    swModel.ClearSelection2 True
    swDraw.ActivateView viewFront.GetName2
    Set cm = Nothing
    If SelPt(viewFront, Array(PKT * 0.7071, PKT * 0.7071, F_FLOOR), False) Then
        On Error Resume Next
        Set cm = swDraw.InsertCenterMark3(0, False, True)
        If Not cm Is Nothing Then cm.ShowLines = True
        On Error GoTo 0
    End If
    swModel.ClearSelection2 True
    Check Not cm Is Nothing, "Front: centre mark with extended lines on the " & ChrW(216) & "3.50 pocket"

    ' 2. the six holes: one circular group (marks turned toward the centre, as on the reference)
    swModel.ClearSelection2 True
    swDraw.ActivateView viewFront.GetName2
    n = 0
    For k = 0 To hN - 1
        idx = HoleCircle(k)
        If idx >= 0 Then
            SelEdge viewFront, idx, (n > 0)
            n = n + 1
        End If
    Next
    Set cm = Nothing
    If n = hN And hN > 0 And swModel.SelectionManager.GetSelectedObjectCount2(-1) = hN Then
        On Error Resume Next
        Set cm = swDraw.InsertCenterMark3(2, False, False)       ' 2 = circular group (the bolt circle is drawn below)
        On Error GoTo 0
    End If
    swModel.ClearSelection2 True
    If cm Is Nothing Then
        For k = 0 To hN - 1
            idx = HoleCircle(k)
            If idx >= 0 Then
                swModel.ClearSelection2 True
                SelEdge viewFront, idx, False
                Set cm = Nothing
                On Error Resume Next
                Set cm = swDraw.InsertCenterMark3(0, False, True)
                On Error GoTo 0
                If Not cm Is Nothing Then nMarks = nMarks + 1
            End If
        Next
        swModel.ClearSelection2 True
        Check (nMarks = hN And hN > 0), "Front: centre marks on the " & nMarks & " holes (single marks - the reference turns them toward the centre)"
    Else
        Check True, "Front: centre marks on the six holes (circular group)"
    End If

    ' 3. dia 2.50 bolt circle, drawn in the view and measured from the left and top edges
    c = SheetPt(viewFront, 0#, 0#, F_FLOOR)
    Set seg = Nothing
    For Each cand In CandOrder()
        Set seg = DrawCirc(viewFront, c, BC * ViewSc(viewFront), CInt(cand))     ' radius on the sheet
        If Not seg Is Nothing Then
            d1 = TestDist(seg, True, Array(-HALF, 0.3, F_RIM))
            d2 = TestDist(seg, False, Array(-0.5, HALF, F_RIM))
            If Abs(d1 - HALF) < 0.005 And Abs(d2 - HALF) < 0.005 Then
                sketchCand = CInt(cand)
                Exit For
            End If
            tried = tried & "m" & cand & ": " & Format(d1, "0.000") & "/" & Format(d2, "0.000") & ". "
            DeleteSeg seg
            Set seg = Nothing
        End If
    Next
    If seg Is Nothing Then
        Check False, "Front: " & ChrW(216) & "2.50 bolt circle could not be put round the part centre (" & Trim(tried) & ")"
        Exit Sub
    End If
    Set segBolt = seg
    CentreStyle segBolt
    Check True, "Front: " & ChrW(216) & "2.50 bolt circle drawn round the part centre (measured from the left and top edges)"

    ' 4. radial centre lines through the 30 and 330 deg holes, for 6 X 60 deg
    cs = Sqr(3#) / 2: sn = 0.5
    For i = 1 To 2
        If i = 2 Then sn = -0.5
        a = SheetPt(viewFront, 1.45 * cs, 1.45 * sn, F_FLOOR)
        b = SheetPt(viewFront, 2.45 * cs, 2.45 * sn, F_FLOOR)
        Set seg = DrawSeg(viewFront, a, b, sketchCand)
        If Not seg Is Nothing Then CentreStyle seg
        If i = 1 Then Set segRay1 = seg Else Set segRay2 = seg
    Next
    Check (Not segRay1 Is Nothing And Not segRay2 Is Nothing), "Front: radial centre lines through the 30 and 330 deg holes"
End Sub

' edge-table circle of hole k: the dia .257 one if there is one, else the smallest circle there
Function HoleCircle(ByVal k As Long) As Long
    Dim i As Long, c As Variant, best As Long, bestR As Double, sc As Double
    HoleCircle = -1
    If k < 0 Then Exit Function
    sc = ViewSc(viewFront)                                ' edge-table radii are sheet sizes
    c = SheetPt(viewFront, hR(k), hU(k), F_FLOOR)
    best = -1: bestR = 1E+30
    For i = 0 To eN - 1
        If eKind(i) = "C" Then
            If Sqr((eCX(i) - c(0)) ^ 2 + (eCY(i) - c(1)) ^ 2) < 0.01 Then
                If Abs(eR(i) / sc - R257) < 0.004 Then HoleCircle = i: Exit Function
                If eR(i) < bestR Then bestR = eR(i): best = i
            End If
        End If
    Next
    HoleCircle = best
End Function

' =====================================================================================
'  View helpers  (from HW#7C)
' =====================================================================================
Function StdIsoName() As String
    If Near(vR, 1, 0, 0) And Near(vU, 0, 1, 0) And Near(vF, 0, 0, 1) Then StdIsoName = "*Isometric"
End Function

' names of all views on the sheet, separated by |
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

' a view that exists now but was not in the list taken earlier
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
        swModel.ForceRebuild3 False
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

Sub CentreLines(v As Object)
    On Error Resume Next
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 v.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    swDraw.InsertCenterLine2
    swModel.ClearSelection2 True
End Sub

Sub DeleteSeg(seg As Object)
    On Error Resume Next
    If seg Is Nothing Then Exit Sub
    swModel.ClearSelection2 True
    seg.Select4 False, Nothing
    swModel.EditDelete
    swModel.ClearSelection2 True
End Sub

' =====================================================================================
'  Dimensions
'  Text positions are offsets from the part, measured on the reference (inches on the sheet).
' =====================================================================================
Sub DimFront()
    Dim cx As Double, cy As Double, why As String, k As Long, D As String
    D = ChrW(216)
    CollectEdges viewFront
    cx = SX(viewFront, 0, 0, 0): cy = SY(viewFront, 0, 0, 0)

    MakeDimT viewFront, "H", Array(-HALF, 0.3, F_RIM), Array(HALF, 0.3, F_RIM), cx, cy - 2.53, 3.875, _
             "Front    3.875" & ChrW(177) & ".003", "", 3, "", False, TOL_SYM, -0.003, 0.003, 3
    k = HoleIdx("CBORE")
    If k >= 0 Then
        MakeDimP viewFront, "V", Array(-0.5, HALF, F_RIM), HolePt(k, 0.22, -0.7071, -0.7071), cx + 2.09, cy + 2.1, 0.6875, _
                 "Front    (.6875)", "", 4, "", True
    Else
        Check False, "Front    (.6875)  - MISSING (counterbore hole not found)"
    End If

    ' dia 2.50 on the sketched bolt circle
    why = "MISSING (bolt circle not drawn)"
    If Not segBolt Is Nothing Then
        swModel.ClearSelection2 True
        swDraw.ActivateView viewFront.GetName2
        On Error Resume Next
        segBolt.Select4 False, Nothing
        On Error GoTo 0
        why = TryDimSel("D", 1, cx - 2.63, cy + 2.09, 2.5, "", 2, "", False, False)
    End If
    Check (Len(why) = 0), "Front    " & D & "2.50 bolt circle" & IIf(Len(why) = 0, "", "  - " & why)

    ' 6 X 60 deg between the radial centre lines
    why = "MISSING (radial centre lines not drawn)"
    If Not segRay1 Is Nothing And Not segRay2 Is Nothing Then
        swModel.ClearSelection2 True
        swDraw.ActivateView viewFront.GetName2
        On Error Resume Next
        segRay1.Select4 False, Nothing
        segRay2.Select4 True, Nothing
        On Error GoTo 0
        why = TryDimSel("A", 2, cx + 2.3, cy, 60#, "6 X ", 0, "", False, False)
    End If
    Check (Len(why) = 0), "Front    6 X 60" & ChrW(176) & IIf(Len(why) = 0, "", "  - " & why)

    ' hole callouts, each on the hole that has that shape
    HoleDim "CBORE", " <HOLE-SPOT> <MOD-DIAM> .44 <HOLE-DEPTH> .25", -2.49, 2.51, "Front    " & D & " .257 cbore " & D & " .44 deep .25"
    HoleDim "D10", " <HOLE-DEPTH> .10", -2.88, 1.66, "Front    " & D & " .257 deep .10"
    HoleDim "THRU", " THRU", 2.35, 2.49, "Front    " & D & " .257 THRU"
    HoleDim "D35", " <HOLE-DEPTH> .35", 2.93, -2.1, "Front    " & D & " .257 deep .35"
    HoleDim "CSK", " <HOLE-SINK> <MOD-DIAM> .50 X 90<MOD-DEG>", 3.35, -2.48, "Front    " & D & " .257 csk " & D & " .50 X 90" & ChrW(176)
    LeaderNoteI viewFront, HoleCircle(HoleIdx("TAP")), "1/4-20 UNC", cx - 3.37, cy + 1.4, "Front    1/4-20 UNC note on the tapped hole"

    ' outside corners: fillet at the lower-left corner (left side meets the chamfer), chamfer midpoint
    MakeDimP viewFront, "R", Array(-1.909, -1.1757, F_RIM), Empty, cx - 2.93, cy - 1.54, 0.375, "Front    16 X R.375", "16 X ", 3
    LeaderNote2 viewFront, Array(-1.5625, -1.5625, F_RIM), Array(-1.45, -1.675, F_RIM), "4X .75 X 45<MOD-DEG>", _
                cx - 3.35, cy - 1.92, "Front    4X .75 X 45" & ChrW(176) & " chamfer note"
End Sub

' dia .257 callout on hole "kind": a diameter dimension with the reference's symbols as suffix
Sub HoleDim(ByVal kind As String, ByVal suffix As String, ByVal dx As Double, ByVal dy As Double, ByVal label As String)
    Dim idx As Long, why As String, t As String, cx As Double, cy As Double
    idx = HoleCircle(HoleIdx(kind))
    If idx < 0 Then Check False, label & "  - MISSING (hole not found)": Exit Sub
    cx = SX(viewFront, 0, 0, 0): cy = SY(viewFront, 0, 0, 0)
    swModel.ClearSelection2 True
    swDraw.ActivateView viewFront.GetName2
    SelEdge viewFront, idx, False
    why = TryDimSel("D", 1, cx + dx, cy + dy, 2 * R257, "", 3, suffix, False, False)
    If Len(why) = 0 Then
        On Error Resume Next                              ' hole-callout spacing, as on the reference
        t = "": t = lastDim.GetText(1)
        If InStr(UCase(t), "DIAM") > 0 Or InStr(t, ChrW(216)) > 0 Then lastDim.SetText 1, "<MOD-DIAM> "
        On Error GoTo 0
    End If
    Check (Len(why) = 0), label & IIf(Len(why) = 0, "", "  - " & why)
End Sub

' note with a leader to edge-table entity idx (-1: note without a leader)
Sub LeaderNoteI(v As Object, ByVal idx As Long, ByVal text As String, ByVal tx As Double, ByVal ty As Double, ByVal label As String)
    Dim nt As Object
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    If idx >= 0 Then SelEdge v, idx, False
    Set nt = Nothing
    On Error Resume Next
    Set nt = swModel.InsertNote(text)
    If Not nt Is Nothing Then nt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
    On Error GoTo 0
    swModel.ClearSelection2 True
    If nt Is Nothing Then
        Check False, label & "  - MISSING"
    ElseIf idx < 0 Then
        Check False, label & "  - placed, but the hole edge was not found, so it has no leader; drag one to the hole"
    Else
        Check True, label
    End If
End Sub

Sub DimSection()
    Dim bx As Double, cy As Double, yT As Double, D As String
    D = ChrW(216)
    CollectEdges viewSec
    bx = SX(viewSec, 0, 0, 0): cy = SY(viewSec, 0, 0, 0)        ' back face, part centre
    yT = SY(viewSec, 0, HALF, 0)                                  ' top of the plate

    ' heights, left of the section (pick points on the cut, away from the .025 chamfers)
    MakeDimT viewSec, "V", Array(0#, HALF, 0.12), Array(0#, -HALF, 0.12), bx - 2.46, cy, 3.875, _
             "Section  3.875" & ChrW(177) & ".003", "", 3, "", False, TOL_SYM, -0.003, 0.003, 3
    MakeDimT viewSec, "V", Array(0#, PAD, 0.36), Array(0#, -PAD, 0.36), bx - 1.7, cy, 3.625, _
             "Section  3.630/3.620", "", 3, "", False, TOL_LIMIT, -0.005, 0.005, 3
    MakeDimT viewSec, "V", Array(0#, PKT, 0.43), Array(0#, -PKT, 0.43), bx - 0.98, cy, 3.5, _
             "Section  " & D & "3.50 +.01/-.00", "<MOD-DIAM>", 2, "", False, TOL_BILAT, 0#, 0.01, 2

    ' thicknesses, stacked .24 apart above the section; .480 and .250 text in one column
    MakeDimT viewSec, "H", Array(0#, 0.5, 0#), Array(0#, 1.78125, F_PAD), bx + 0.58, yT + 0.72, 0.48, _
             "Section  .480" & ChrW(177) & ".005", "", 3, "", False, TOL_SYM, -0.005, 0.005, 3
    MakeDimT viewSec, "H", Array(0#, 1.78125, F_PAD), Array(0#, 1.61, F_FLOOR), bx - 0.93, yT + 0.48, 0.08, _
             "Section  .085/.075 (pocket depth)", "", 3, "", False, TOL_LIMIT, -0.005, 0.005, 3
    MakeDimP viewSec, "H", Array(0#, 0.5, 0#), Array(0#, 1.86, F_RIM), bx + 0.58, yT + 0.24, 0.25, "Section  .250", "", 3

    LeaderNote2 viewSec, Array(0#, -1.925, 0.2375), Array(0#, -1.9375, 0.12), ".025 X 45<MOD-DEG> TYP", _
                bx - 1.99, cy - 2.12, "Section  .025 X 45" & ChrW(176) & " TYP note"
End Sub

' a dimension with a tolerance: tolType TOL_SYM / TOL_LIMIT / TOL_BILAT, tMin / tMax in inches
Sub MakeDimT(v As Object, ByVal kind As String, pa As Variant, pb As Variant, ByVal tx As Double, ByVal ty As Double, _
             ByVal expect As Double, ByVal label As String, ByVal prefix As String, ByVal prec As Long, _
             ByVal suffix As String, ByVal paren As Boolean, ByVal tolType As Long, ByVal tMin As Double, _
             ByVal tMax As Double, ByVal tolPrec As Long)
    Dim why As String
    why = TryDim(v, kind, pa, pb, tx, ty, expect, prefix, prec, suffix, paren, False)
    If Len(why) = 0 Then why = SetTol(lastDim, tolType, tMin, tMax, prec, tolPrec)
    Check (Len(why) = 0), label & IIf(Len(why) = 0, "", "  - " & why)
End Sub

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

' datum A on the back face, and the // .002 A frame above the pad's front face
Sub GtolAndDatum()
    Dim bx As Double, cy As Double, yT As Double, xP As Double, dt As Object, g As Object
    Dim okS As Boolean, okV As Boolean
    CollectEdges viewSec
    bx = SX(viewSec, 0, 0, 0): cy = SY(viewSec, 0, 0, 0)
    yT = SY(viewSec, 0, HALF, 0): xP = SX(viewSec, 0, 0, F_PAD)

    swModel.ClearSelection2 True
    swDraw.ActivateView viewSec.GetName2
    Set dt = Nothing
    If SelPt(viewSec, Array(0#, 0.185, 0#), False) Then
        On Error Resume Next
        Set dt = swModel.InsertDatumTag2
        If Not dt Is Nothing Then
            dt.SetLabel "A"
            dt.GetAnnotation.SetPosition2 (bx + 0.49) * IN_M, (cy + 0.185) * IN_M, 0#
        End If
        On Error GoTo 0
    End If
    swModel.ClearSelection2 True
    If dt Is Nothing Then
        Check False, "Section  datum A - MISSING; add it with Annotations > Datum Feature on the back face"
    Else
        Check True, "Section  datum A on the back face"
    End If

    swModel.ClearSelection2 True
    swDraw.ActivateView viewSec.GetName2
    Set g = Nothing
    On Error Resume Next
    Set g = swModel.InsertGtol
    If Not g Is Nothing Then
        okS = g.SetFrameSymbols2(1, "<IGTOL-PARA>", False, "", False, "", "", "", "")
        If Err.Number <> 0 Or Not okS Then Err.Clear: okS = g.SetFrameSymbols3(1, "<IGTOL-PARA>", False, "", False, "", "", "", "")
        Err.Clear
        okV = g.SetFrameValues2(1, ".002", "", "A", "", "")
        If Err.Number <> 0 Or Not okV Then Err.Clear: okV = g.SetFrameValues(1, ".002", "", "A", "", "")
        Err.Clear
        g.GetAnnotation.SetPosition2 (xP - 0.89) * IN_M, (yT + 1.06) * IN_M, 0#
    End If
    On Error GoTo 0
    swModel.ClearSelection2 True
    If g Is Nothing Then
        Check False, "Section  // .002 A - MISSING; add it with Annotations > Geometric Tolerance"
    ElseIf Not (okS And okV) Then
        Check False, "Section  // .002 A frame placed, but its " & IIf(okS, "values", "symbol") & " did not take; double-click it and set Parallelism, .002, A"
    Else
        Check True, "Section  // .002 A parallelism frame (look that the // symbol shows)"
    End If
End Sub

' the reference shows no "SECTION A-A" label under the section, so it is hidden
Sub HideSectionLabel()
    Dim v As Object, anns As Variant, i As Long, t As String, found As Boolean, vis As Long, orig As Long, want As Long
    On Error Resume Next
    swModel.ForceRebuild3 False
    Set v = swDraw.GetFirstView                         ' the sheet, then each drawing view
    Do While Not v Is Nothing And Not found
        anns = Empty
        anns = v.GetAnnotations
        If IsArray(anns) Then
            For i = LBound(anns) To UBound(anns)
                t = ""
                t = UCase(anns(i).GetSpecificAnnotation.GetText)
                If Len(t) = 0 Then t = UCase(anns(i).GetTextAtIndex(0))
                If InStr(t, "SECTION") > 0 Then
                    found = True
                    orig = -1: orig = anns(i).Visible   ' the label is shown now: this is the "visible" value
                    want = IIf(orig = 1, 0, 1)          ' swAnnotationHidden is 1 (visible 0); never re-set orig
                    anns(i).Visible = want
                    vis = -1: vis = anns(i).Visible
                    Exit For
                End If
            Next
        End If
        Set v = v.GetNextView
    Loop
    swModel.GraphicsRedraw2
    On Error GoTo 0
    If Not found Then
        Check True, "SECTION A-A label: none shown (as on the reference)"
    Else
        Check (vis = want And vis <> orig), "SECTION A-A label hidden (as on the reference; state " & orig & " -> " & vis & ")" & _
              IIf(vis = want And vis <> orig, " - if you still see it, right-click it > Hide", " - right-click the label under the section > Hide")
    End If
End Sub

' "perfectly aligned": every view where it was measured, the section level with the front view
Sub CheckAlignment()
    Dim a As Variant, b As Variant, c As Variant, ok As Boolean
    swModel.ForceRebuild3 False
    a = SheetPt(viewFront, 0, 0, 0)
    ok = (Abs(a(0) - FRONT_X) < 0.002 And Abs(a(1) - FRONT_Y) < 0.002)
    Check ok, "Alignment: front view centre at " & Format(a(0), "0.000") & ", " & Format(a(1), "0.000") & _
              " (wanted " & Format(FRONT_X, "0.000") & ", " & Format(FRONT_Y, "0.000") & ")"
    If Not viewSec Is Nothing Then
        b = SheetPt(viewSec, 0, 0, 0)
        ok = (Abs(b(1) - a(1)) < 0.001)
        Check ok, "Alignment: SECTION A-A level with the front view (centres at y " & Format(a(1), "0.000") & " / " & Format(b(1), "0.000") & ")" & _
                  IIf(ok, "", " - right-click the section > Alignment > Default alignment")
        ok = (Abs(b(0) - SEC_X) < 0.002)
        Check ok, "Alignment: SECTION A-A back face at x " & Format(b(0), "0.000") & " (wanted " & Format(SEC_X, "0.000") & ")"
    End If
    If Not viewIso Is Nothing Then
        c = SheetPt(viewIso, 0, 0, F_PAD / 2)
        ok = (Abs(c(0) - ISO_X) < 0.002 And Abs(c(1) - ISO_Y) < 0.002)
        Check ok, "Alignment: isometric view centre at " & Format(c(0), "0.000") & ", " & Format(c(1), "0.000")
    End If
End Sub

' =====================================================================================
'  Title block, finish, save
' =====================================================================================
Sub TitleBlock()
    Dim pass As Integer, sv As Object, nt As Object, t As String, pos As Variant
    Dim mX As Double, mY As Double, dX As Double, dY As Double, tX As Double, tY As Double
    Dim gotM As Boolean, gotD As Boolean, gotT As Boolean, sheetDone As Boolean, sheetSeen As Boolean, scaleDone As Boolean
    Dim twoDone As Boolean, threeDone As Boolean, tolSeen As Boolean

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
                nt.SetText "SHEET 1 OF 4"
                If UCase(Trim(nt.GetText)) <> "SHEET 1 OF 4" Then nt.PropertyLinkedText = "SHEET 1 OF 4"
                sheetDone = (UCase(Trim(nt.GetText)) = "SHEET 1 OF 4")
            End If
            If Left(t, 6) = "SCALE:" And Not scaleDone Then              ' sheet is 1:1
                If InStr(t, "1:1") = 0 Then nt.SetText "SCALE: 1:1"
                If InStr(UCase(nt.GetText), "1:1") = 0 Then nt.PropertyLinkedText = "SCALE: 1:1"
                scaleDone = (InStr(UCase(nt.GetText), "1:1") > 0)
            End If
            If InStr(t, "PLACE DECIMAL") > 0 Then                         ' note 2: tolerance block
                tolSeen = True
                FixTol nt, twoDone, threeDone
            End If
            Set nt = nt.GetNext
        Loop
        If pass = 2 Then swDraw.EditSheet
        On Error GoTo 0
        If gotM And gotD And gotT And sheetDone And scaleDone And twoDone And threeDone Then Exit For
    Next
    On Error Resume Next
    swDraw.EditSheet                                    ' make sure we are NOT left in Edit Sheet Format mode
    On Error GoTo 0

    ' fall back to the reference layout if a label was not found
    If Not gotM Then mX = 5.07: mY = 0.94
    If Not gotD Then dX = 3.17: dY = 0.48
    If Not gotT Then tX = 7.2: tY = 0.92
    Check gotM And gotD And gotT, "Title block labels found (MATERIAL, DATE, TITLE)" & _
          IIf(gotM And gotD And gotT, "", " - used the reference positions instead")

    AddSheetNote "6061-T6 ALUMINUM", mX + 0.06, mY - 0.17, "Material: 6061-T6 ALUMINUM"
    AddSheetNote Format(Date, "mm-dd-yyyy"), dX + 0.4, dY - 0.18, "Date: " & Format(Date, "mm-dd-yyyy")
    AddSheetNote "IME 1140 ENGINEERING DRAWINGS II", tX + 0.44, tY, "Title line 1"
    AddSheetNote "IN-LAB ASSIGNMENT #8A", tX + 0.44, tY - 0.17, "Title line 2"
    Check scaleDone, "SCALE: 1:1" & IIf(scaleDone, "", " - edit the scale note by hand")
    Check sheetDone, "SHEET 1 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
    Check (twoDone And threeDone), "Tolerance block: TWO PLACE DECIMAL " & ChrW(177) & ".02, THREE PLACE DECIMAL " & ChrW(177) & ".008 (note 2)" & _
          IIf(twoDone And threeDone, "", IIf(tolSeen, " - the note would not change; edit it by hand (Edit Sheet Format)", " - tolerance note not found; edit it by hand"))
End Sub

' change the TWO / THREE PLACE DECIMAL lines of a tolerance note to the reference values
Sub FixTol(nt As Object, twoDone As Boolean, threeDone As Boolean)
    Dim t As String, u As String
    On Error Resume Next
    t = nt.GetText
    t = NewLine(t, "TWO PLACE DECIMAL", "TWO PLACE DECIMAL <MOD-PM> .02")
    t = NewLine(t, "THREE PLACE DECIMAL", "THREE PLACE DECIMAL  <MOD-PM> .008")
    nt.SetText t
    u = UCase(nt.GetText)
    If InStr(u, "TWO PLACE DECIMAL") > 0 Then twoDone = (InStr(LineOf(u, "TWO PLACE DECIMAL"), ".02") > 0 And InStr(LineOf(u, "TWO PLACE DECIMAL"), ".01") = 0)
    If InStr(u, "THREE PLACE DECIMAL") > 0 Then threeDone = (InStr(LineOf(u, "THREE PLACE DECIMAL"), ".008") > 0)
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

' =====================================================================================
'  Dimension and selection helpers  (from HW#7D)
' =====================================================================================
Sub LeaderNote2(v As Object, p1 As Variant, p2 As Variant, ByVal text As String, ByVal tx As Double, ByVal ty As Double, ByVal label As String)
    Dim nt As Object, hasLeader As Boolean
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    hasLeader = SelPt(v, p1, False)
    If Not hasLeader Then hasLeader = SelPt(v, p2, False)
    Set nt = Nothing
    On Error Resume Next
    Set nt = swModel.InsertNote(text)
    If Not nt Is Nothing Then nt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
    On Error GoTo 0
    swModel.ClearSelection2 True
    If nt Is Nothing Then
        Check False, label & "  - MISSING"
    ElseIf Not hasLeader Then
        Check False, label & "  - placed, but no edge found for its leader; drag the leader to the chamfer"
    Else
        Check True, label
    End If
End Sub

' kind: H horizontal, V vertical (2 edges), VL vertical length of 1 edge, S aligned, A angle (degrees), D diameter, R radius
Sub MakeDimP(v As Object, ByVal kind As String, pa As Variant, pb As Variant, ByVal tx As Double, ByVal ty As Double, _
             ByVal expect As Double, ByVal label As String, Optional ByVal prefix As String = "", Optional ByVal prec As Long = -1, _
             Optional ByVal suffix As String = "", Optional ByVal paren As Boolean = False)
    Dim why As String
    why = TryDim(v, kind, pa, pb, tx, ty, expect, prefix, prec, suffix, paren, False)
    Check (Len(why) = 0), label & IIf(Len(why) = 0, "", "  - " & why)
End Sub

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
        If deleteIfWrong Then
            On Error Resume Next
            swModel.ClearSelection2 True
            dd.GetAnnotation.Select3 False, Nothing
            swModel.EditDelete
            swModel.ClearSelection2 True
            On Error GoTo 0
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

' =====================================================================================
'  Edge table: every visible edge of a view, in sheet inches  (from HW#7C)
' =====================================================================================
Sub CollectEdges(v As Object)
    Dim ents As Variant, i As Long, xf As Object, sc As Double
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
    If Not IsArray(ents) Then Exit Sub
    For i = LBound(ents) To UBound(ents)
        If eN < EMAX Then AddEdge ents(i), xf, sc
    Next
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
            n = 0
            For i = LBound(r) To UBound(r)
                Set c = Nothing
                On Error Resume Next
                Set c = r(i).GetCurve
                On Error GoTo 0
                If Not c Is Nothing Then n = n + 1
            Next
            If n > bestN Then bestN = n: best = r
        End If
    Next
    VisibleEdges = best
End Function

Sub AddEdge(ByVal ent As Object, xf As Object, ByVal sc As Double)
    ' Extents come only from exact geometry: line end points, and circle/arc centre +- radius.
    ' (For this part every arc's full circle stays inside the outline, so centre +- radius is exact.)
    Dim crv As Object, cp As Variant, a As Variant, b As Variant, p1 As Variant, p2 As Variant
    Dim isL As Boolean, isC As Boolean, cc As Variant, axv As Variant, ctr As Variant, L As Double, rr As Double

    On Error Resume Next
    Set crv = ent.GetCurve
    isL = crv.IsLine
    isC = crv.IsCircle
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

Sub Extend(p As Variant)
    If p(0) < xMin Then xMin = p(0)
    If p(0) > xMax Then xMax = p(0)
    If p(1) < yMin Then yMin = p(1)
    If p(1) > yMax Then yMax = p(1)
End Sub

' =====================================================================================
'  Finish and save  (from HW#7C)
' =====================================================================================
Sub FinishSheet()
    ' leave the drawing active, in normal sheet mode, with all three views shown and redrawn
    Dim vs As Variant, n As Long, i As Long, v As Variant
    On Error Resume Next
    ActivateDrawing
    swDraw.EditSheet
    swModel.ClearSelection2 True
    For Each v In Array(viewFront, viewSec, viewIso)
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
    If n >= 0 Then Check (n >= 3), n & " views on the sheet (expected 3)"
End Sub

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
    Print #f, "Passed: " & nOK & "   Problems: " & nBad
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
'  Geometry helpers  (from HW#7C)
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

' make the largest component positive, so parallel axes compare equal
Function Canon(v As Variant) As Variant
    Dim k As Integer, m As Integer
    m = 0
    For k = 1 To 2
        If Abs(v(k)) > Abs(v(m)) Then m = k
    Next
    If v(m) < 0 Then Canon = Scl(v, -1) Else Canon = v
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
