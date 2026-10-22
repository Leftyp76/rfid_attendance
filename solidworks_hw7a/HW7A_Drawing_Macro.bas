' =====================================================================================
'  HW#7A  -  IME 1140 In-Lab Assignment #7A  -  SolidWorks DRAWING macro   (inches)
'
'  HOW TO RUN
'   1. Open "IME 1140 Engineering Drawings I In-Lab Assignment Part #1.SLDPRT" in SolidWorks.
'      (It must be a saved file on your computer, not a preview.)
'   2. Tools > Macro > New...  save it as  HW7A_Drawing.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main  (Tools > Macro > Run, or the Macros dialog with "main" selected).
'   5. The macro builds the drawing on your template, checks every dimension, asks for your
'      last name, and saves  HW#7A<LastName>.SLDDRW  and  HW#7A<LastName>.PDF  next to the part.
'
'  WHAT IT BUILDS (matches the reference sheet)
'   - Front view (face of the upright, flange at the bottom), Top and Right projected from it
'     (third angle), hidden lines visible.  Isometric view, hidden lines removed.  Scale 1:1.
'   - Centre marks on the holes, centrelines.
'   - Top:   2 X dia.50, 1.00, .50, 2.00, .50, 1.50
'   - Front: R1.00, dia1.00, 1.50, 1.00
'   - Right: .25, 3.00, .25, R.25, 2 X R.125, .025 X 45 deg TYP
'   - Title block: 316 STAINLESS STEEL, today's date, title, SHEET 1 OF 4.
'
'  SAFETY
'   - The part's own orientation is read from its holes, so the front view is always the face
'     of the upright with the flange at the bottom, however the part was modelled.
'   - Every view's orientation is verified after it is created.
'   - Every dimension is read back and compared with the drawing value; the final message
'     lists each one as OK / WRONG / MISSING.
'   - The part file is never saved or changed by this macro.
' =====================================================================================
Option Explicit

Dim swApp As Object              ' SldWorks
Dim swMath As Object             ' MathUtility
Dim swPart As Object             ' the part (ModelDoc2)
Dim partPath As String
Dim swModel As Object            ' the drawing (ModelDoc2)
Dim swDraw As Object             ' the drawing (DrawingDoc)
Dim curStep As String
Dim report As String             ' one line per check
Dim problems As String           ' only the PROBLEM lines (MsgBox can show ~1000 characters)
Dim nOK As Long, nBad As Long

Const IN_M = 0.0254              ' metres per inch (the API always works in metres)

' ---- sheet layout, view centres in inches from the sheet's lower-left corner ----
Const FRONT_X = 2.46
Const FRONT_Y = 3.23
Const TOP_Y = 6.8
Const RIGHT_X = 6.66
Const ISO_X = 9.46
Const ISO_Y = 6.51

' ---- display modes (swDisplayMode_e) ----
Const DM_HIDDEN_VISIBLE = 1      ' swHIDDEN_GREYED  = hidden lines visible (dashed)
Const DM_HIDDEN_REMOVED = 2      ' swHIDDEN         = hidden lines removed

' ---- the part's own directions (model coordinates, unit vectors) ----
Dim vR As Variant                ' right  in the front view
Dim vU As Variant                ' up     in the front view (bottom of flange -> top of upright)
Dim vF As Variant                ' toward the viewer in the front view (upright -> flange tip)

Dim viewFront As Object, viewTop As Object, viewRight As Object, viewIso As Object

' ---- edge table of the view being dimensioned (sheet inches) ----
Dim eN As Long
Dim eObj() As Object, eKind() As String
Dim eX1() As Double, eY1() As Double, eX2() As Double, eY2() As Double
Dim eCX() As Double, eCY() As Double, eR() As Double, eFull() As Boolean
Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double

' =====================================================================================
Sub main()
    On Error GoTo Fail
    Dim tmpl As String

    Set swApp = Application.SldWorks
    Set swMath = swApp.GetMathUtility
    report = "": problems = "": nOK = 0: nBad = 0

    curStep = "Step 1 - find the open part":          FindPart
    curStep = "Step 2 - read the part orientation":   FindOrientation
    curStep = "Step 3 - find your template":          tmpl = FindTemplate()
    curStep = "Step 4 - new drawing":                 NewDrawing tmpl
    curStep = "Step 5 - views":                       MakeViews
    curStep = "Step 6 - centre marks and centrelines": CentreMarksAndLines
    curStep = "Step 7 - Top view dimensions":         DimTop
    curStep = "Step 8 - Front view dimensions":       DimFront
    curStep = "Step 9 - Right view dimensions":       DimRight
    curStep = "Step 10 - title block":                TitleBlock

    curStep = "Step 11 - finish"
    swModel.ClearSelection2 True
    swModel.ForceRebuild3 False
    swModel.ViewZoomtofit2

    MsgBox "HW#7A drawing built." & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & vbCrLf & vbCrLf & _
           IIf(nBad = 0, "Every view, dimension and title block item checked out.", "Problems:" & vbCrLf & Left(problems, 700)) & _
           vbCrLf & vbCrLf & "The full check list is saved as HW7A_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), "HW#7A"

    curStep = "Step 12 - save drawing and PDF":       SaveAll
    Exit Sub

Fail:
    MsgBox "The macro stopped at: " & curStep & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Checks passed so far: " & nOK & vbCrLf & _
           IIf(nBad = 0, "", "Problems so far:" & vbCrLf & Left(problems, 500)), vbCritical, "HW#7A macro"
End Sub

' =====================================================================================
'  Step 1 - the part
' =====================================================================================
Sub FindPart()
    Dim d As Object, docs As Variant, i As Long
    Set swPart = Nothing
    Set d = swApp.ActiveDoc
    If Not d Is Nothing Then
        If d.GetType = 1 Then Set swPart = d                       ' 1 = swDocPART
    End If
    If swPart Is Nothing Then
        On Error Resume Next
        docs = swApp.GetDocuments
        On Error GoTo 0
        If IsArray(docs) Then
            For i = LBound(docs) To UBound(docs)
                If docs(i).GetType = 1 Then Set swPart = docs(i): Exit For
            Next
        End If
    End If
    If swPart Is Nothing Then Err.Raise vbObjectError + 1, , _
        "No part is open. Open the In-Lab Assignment Part #1 file first, then run the macro again."
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then Err.Raise vbObjectError + 2, , _
        "The part has never been saved, so a drawing cannot reference it. Save the part first."
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

' =====================================================================================
'  Step 2 - which way does the part face?  Read it from the holes.
'   Big hole  (dia 1.00): its axis is the front-view direction.
'   Small holes (dia .50): their axis is "up".
' =====================================================================================
Sub FindOrientation()
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, k As Long
    Dim s As Object, cp As Variant, bx As Variant, ax As Variant, r As Double
    Dim bigAx As Variant, bigC(2) As Double, bigN As Long
    Dim gAx(19) As Variant, gCnt(19) As Long, gC(19, 2) As Double, gN As Long, g As Long, hit As Long
    Dim smC(2) As Double, best As Long, dd As Double

    bodies = swPart.GetBodies2(0, True)                             ' 0 = swSolidBody
    If Not IsArray(bodies) Then Err.Raise vbObjectError + 3, , "The part has no solid body."
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder Then
                cp = s.CylinderParams                               ' origin(3) axis(3) radius
                r = cp(6) / IN_M
                ax = Canon(Unit(Arr3(cp(3), cp(4), cp(5))))
                bx = faces(i).GetBox
                If Abs(r - 0.5) < 0.005 Then                        ' dia 1.00 hole
                    bigAx = ax
                    For k = 0 To 2: bigC(k) = bigC(k) + (bx(k) + bx(k + 3)) / 2: Next
                    bigN = bigN + 1
                ElseIf Abs(r - 0.25) < 0.005 Then                   ' dia .50 holes (and the R.25 bend)
                    hit = -1
                    For g = 0 To gN - 1
                        If Abs(Dot(gAx(g), ax)) > 0.999 Then hit = g: Exit For
                    Next
                    If hit < 0 And gN < 20 Then hit = gN: gAx(gN) = ax: gN = gN + 1
                    If hit >= 0 Then
                        gCnt(hit) = gCnt(hit) + 1
                        For k = 0 To 2: gC(hit, k) = gC(hit, k) + (bx(k) + bx(k + 3)) / 2: Next
                    End If
                End If
            End If
        Next
    Next
    If bigN = 0 Then Err.Raise vbObjectError + 4, , "Could not find the dia 1.00 hole in the part."

    best = -1
    For g = 0 To gN - 1
        If Abs(Dot(gAx(g), bigAx)) < 0.05 Then                      ' must be at 90 deg to the big hole
            If best < 0 Then
                best = g
            ElseIf gCnt(g) > gCnt(best) Then
                best = g
            End If
        End If
    Next
    If best < 0 Then Err.Raise vbObjectError + 5, , "Could not find the two dia .50 holes in the part."

    For k = 0 To 2
        bigC(k) = bigC(k) / bigN
        smC(k) = gC(best, k) / gCnt(best)
    Next
    vF = bigAx
    vU = gAx(best)
    ' signs: the flange is in front of the upright; the big hole is above the flange
    dd = (smC(0) - bigC(0)) * vF(0) + (smC(1) - bigC(1)) * vF(1) + (smC(2) - bigC(2)) * vF(2)
    If dd < 0 Then vF = Scl(vF, -1)
    dd = (bigC(0) - smC(0)) * vU(0) + (bigC(1) - smC(1)) * vU(1) + (bigC(2) - smC(2)) * vU(2)
    If dd < 0 Then vU = Scl(vU, -1)
    vR = Cross(vU, vF)
    Check True, "Part orientation: up = " & VecTxt(vU) & ", front = " & VecTxt(vF)
End Sub

' =====================================================================================
'  Step 3 - template
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
                 "then paste it here:", "HW#7A - drawing template")
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

' =====================================================================================
'  Step 4 - new drawing, scale 1:1, third angle
' =====================================================================================
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
        sh.SetProperties2 CLng(p(0)), CLng(p(1)), 1#, 1#, False, CDbl(p(5)), CDbl(p(6)), True
    End If
    p = sh.GetProperties2
    On Error GoTo 0
    If IsArray(p) Then
        Check (Abs(p(2) - 1) < 0.001 And Abs(p(3) - 1) < 0.001), "Sheet scale 1:1"
        Check (p(4) = 0 Or p(4) = False), "Sheet projection: third angle"
    End If
End Sub

' =====================================================================================
'  Step 5 - views
' =====================================================================================
Sub MakeViews()
    Dim d As Variant, upI As Variant, rtI As Variant

    ' Front
    Set viewFront = OrientedView("HW7A_Front", vR, vU, vF, FRONT_X, FRONT_Y, StdFrontName(), False)
    If viewFront Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented front view."
    Check True, "Front view: face of the upright, flange at the bottom"
    SetMode viewFront, DM_HIDDEN_VISIBLE, "Front"

    ' Top and Right, projected from the front view
    Set viewTop = Projected(viewFront, FRONT_X, TOP_Y)
    If viewTop Is Nothing Then Err.Raise vbObjectError + 10, , "Could not create the projected top view."
    Check FrameOK(viewTop, vR, 1, 0, vF, 0, -1), "Top view: flange toward the bottom of the sheet"
    SetMode viewTop, DM_HIDDEN_VISIBLE, "Top"

    Set viewRight = Projected(viewFront, RIGHT_X, FRONT_Y)
    If viewRight Is Nothing Then Err.Raise vbObjectError + 11, , "Could not create the projected right view."
    Check FrameOK(viewRight, vF, -1, 0, vU, 0, 1), "Right view: flange pointing left (third angle)"
    SetMode viewRight, DM_HIDDEN_VISIBLE, "Right"

    ' Isometric: looking from front-right-top of the part
    d = Unit(Add3(Add3(vR, vU), vF))
    upI = Unit(Sub3(vU, Scl(d, Dot(vU, d))))
    rtI = Cross(upI, d)
    Set viewIso = OrientedView("HW7A_Iso", rtI, upI, d, ISO_X, ISO_Y, StdIsoName(), True)
    If viewIso Is Nothing Then Err.Raise vbObjectError + 12, , "Could not create a correctly oriented isometric view."
    Check True, "Isometric view"
    SetMode viewIso, DM_HIDDEN_REMOVED, "Isometric"
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

Function StdIsoName() As String
    If Near(vR, 1, 0, 0) And Near(vU, 0, 1, 0) And Near(vF, 0, 0, 1) Then StdIsoName = "*Isometric"
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

Function Projected(parent As Object, ByVal px As Double, ByVal py As Double) As Object
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 parent.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    On Error Resume Next
    Set Projected = swDraw.CreateUnfoldedViewAt3(px * IN_M, py * IN_M, 0#, False)
    On Error GoTo 0
    swModel.ClearSelection2 True
End Function

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
'  Step 6 - centre marks on the holes (Top and Front) and centrelines (Top, Front, Right)
' =====================================================================================
Sub CentreMarksAndLines()
    CentreMarks viewTop, "Top", 2
    CentreMarks viewFront, "Front", 1
    CentreLines viewTop
    CentreLines viewFront
    CentreLines viewRight
End Sub

Sub CentreMarks(v As Object, ByVal label As String, ByVal want As Long)
    Dim i As Long, j As Long, n As Long, dup As Boolean, cm As Object, existing As Long
    Dim doneX(20) As Double, doneY(20) As Double, nd As Long
    On Error Resume Next
    existing = 0
    existing = v.GetCenterMarkCount
    On Error GoTo 0
    If existing >= want Then
        Check True, label & " view: centre marks (already added by the template)"
        Exit Sub
    End If
    CollectEdges v
    For i = 0 To eN - 1
        If eKind(i) = "C" And eFull(i) And (Abs(eR(i) - 0.25) < 0.01 Or Abs(eR(i) - 0.5) < 0.01) Then
            dup = False
            For j = 0 To nd - 1
                If Abs(doneX(j) - eCX(i)) < 0.01 And Abs(doneY(j) - eCY(i)) < 0.01 Then dup = True
            Next
            If Not dup And nd < 20 Then
                swModel.ClearSelection2 True
                swDraw.ActivateView v.GetName2
                SelEdge v, i, False
                Set cm = Nothing
                On Error Resume Next
                Set cm = swDraw.InsertCenterMark3(0, False, True)
                On Error GoTo 0
                If Not cm Is Nothing Then n = n + 1
                doneX(nd) = eCX(i): doneY(nd) = eCY(i): nd = nd + 1
            End If
        End If
    Next
    swModel.ClearSelection2 True
    Check (n >= want), label & " view: " & n & " centre mark(s), expected " & want
End Sub

Sub CentreLines(v As Object)
    On Error Resume Next
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 v.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    swDraw.InsertCenterLine2
    swModel.ClearSelection2 True
End Sub

' =====================================================================================
'  Steps 7-9 - dimensions
' =====================================================================================
Sub DimTop()
    Dim hL As Long, hR As Long, eLeft As Long, eRight As Long, eFront As Long, eBack As Long
    Dim r1 As Double, r2 As Double
    CollectEdges viewTop
    hL = PickCircle(0.25, 0.01, "minx", 1)
    hR = PickCircle(0.25, 0.01, "maxx", 1)
    eLeft = PickLine(True, xMin, 0.01)
    eRight = PickLine(True, xMax, 0.01)
    eFront = PickLine(False, yMin, 0.01)
    eBack = PickLine(False, yMax, 0.01)
    r1 = yMin - 0.3: r2 = yMin - 0.7
    MakeDim viewTop, "D", hL, -1, xMin - 0.55, yMin - 0.25, 0.5, "Top    2 X " & ChrW(216) & ".50", "2 X "
    MakeDim viewTop, "H", hL, hR, (CX(hL) + CX(hR)) / 2, r1, 1#, "Top    1.00  (hole to hole)"
    MakeDim viewTop, "H", hR, eRight, xMax + 0.35, r1, 0.5, "Top    .50  (hole to side)"
    MakeDim viewTop, "H", eLeft, eRight, (xMin + xMax) / 2, r2, 2#, "Top    2.00"
    MakeDim viewTop, "V", hR, eFront, xMax + 0.3, (yMin + CY(hR)) / 2, 0.5, "Top    .50  (hole to front edge)"
    MakeDim viewTop, "V", eFront, eBack, xMax + 0.65, (yMin + yMax) / 2, 1.5, "Top    1.50"
End Sub

Sub DimFront()
    Dim hole As Long, arc As Long, eBot As Long, eRt As Long
    CollectEdges viewFront
    hole = PickCircle(0.5, 0.01, "any", 1)
    arc = PickCircle(1#, 0.01, "maxy", 0)
    eBot = PickLine(False, yMin, 0.01)
    eRt = PickLine(True, xMax, 0.01)
    MakeDim viewFront, "R", arc, -1, xMax + 0.6, yMax + 0.15, 1#, "Front  R1.00"
    MakeDim viewFront, "D", hole, -1, xMax + 0.6, yMax - 0.25, 1#, "Front  " & ChrW(216) & "1.00"
    MakeDim viewFront, "V", hole, eBot, xMax + 0.4, (yMin + CY(hole)) / 2, 1.5, "Front  1.50  (hole height)"
    MakeDim viewFront, "H", hole, eRt, (CX(hole) + xMax) / 2, yMin - 0.4, 1#, "Front  1.00  (hole to side)"
End Sub

Sub DimRight()
    Dim eOut As Long, eIn As Long, eTop As Long, eBot As Long, eFl As Long, a25 As Long, a125 As Long, ch As Long
    CollectEdges viewRight
    eOut = PickLine(True, xMax, 0.01)                  ' back face of the upright
    eIn = PickLine(True, xMax - 0.25, 0.015)           ' front face of the upright
    eTop = PickLine(False, yMax, 0.01)
    eBot = PickLine(False, yMin, 0.01)
    eFl = PickLine(False, yMin + 0.25, 0.015)          ' top of the flange
    a25 = PickCircle(0.25, 0.01, "any", 0)             ' inside bend
    a125 = PickCircle(0.125, 0.01, "maxdiag", 0)       ' outside bend (lower right)
    ch = PickChamfer()
    MakeDim viewRight, "H", eIn, eOut, xMax + 0.35, yMax + 0.25, 0.25, "Right  .25  (upright)"
    MakeDim viewRight, "V", eTop, eBot, xMax + 0.7, (yMin + yMax) / 2, 3#, "Right  3.00"
    MakeDim viewRight, "V", eFl, eBot, xMin - 0.4, yMin - 0.3, 0.25, "Right  .25  (flange)"
    MakeDim viewRight, "R", a25, -1, xMin + 0.35, yMin + 0.8, 0.25, "Right  R.25"
    MakeDim viewRight, "R", a125, -1, xMax + 0.9, yMin - 0.3, 0.125, "Right  2 X R.125", "2 X ", 3
    LeaderNote viewRight, ch, ".025 X 45<MOD-DEG> TYP", xMin - 0.95, yMin + 0.8, "Right  .025 X 45" & ChrW(176) & " TYP"
End Sub

' kind: H horizontal, V vertical, D diameter, R radius.  ia/ib = edge indices (ib = -1 for D and R)
Sub MakeDim(v As Object, ByVal kind As String, ByVal ia As Long, ByVal ib As Long, ByVal tx As Double, ByVal ty As Double, _
            ByVal expect As Double, ByVal label As String, Optional ByVal prefix As String = "", Optional ByVal prec As Long = -1)
    Dim dd As Object, got As Double, nSel As Long, need As Long, txt As String

    If kind = "H" Or kind = "V" Then need = 2 Else need = 1
    If ia < 0 Or (need = 2 And ib < 0) Then
        Check False, label & "  - MISSING (could not find the edge in the view)"
        Exit Sub
    End If

    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    SelEdge v, ia, False
    If need = 2 Then SelEdge v, ib, True
    nSel = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    If nSel <> need Then
        Check False, label & "  - MISSING (selected " & nSel & " edge(s), needed " & need & ")"
        swModel.ClearSelection2 True
        Exit Sub
    End If

    Set dd = Nothing
    On Error Resume Next
    Select Case kind
        Case "H": Set dd = swModel.AddHorizontalDimension2(tx * IN_M, ty * IN_M, 0#)
        Case "V": Set dd = swModel.AddVerticalDimension2(tx * IN_M, ty * IN_M, 0#)
        Case "D": Set dd = swModel.AddDiameterDimension2(tx * IN_M, ty * IN_M, 0#)
        Case "R": Set dd = swModel.AddRadialDimension2(tx * IN_M, ty * IN_M, 0#)
    End Select
    On Error GoTo 0
    swModel.ClearSelection2 True
    If dd Is Nothing Then
        Check False, label & "  - MISSING (SolidWorks did not create it)"
        Exit Sub
    End If

    On Error Resume Next
    got = -1
    got = dd.GetDimension2(0).SystemValue / IN_M
    dd.ShowParenthesis = False
    If prec >= 0 Then
        dd.SetPrecision3 prec, -1, -1, -1
        If Err.Number <> 0 Then Err.Clear: dd.SetPrecision2 prec, -1, -1, -1
    End If
    If Len(prefix) > 0 Then
        dd.SetText 1, prefix                                    ' 1 = swDimensionTextPrefix
        txt = dd.GetText(1)
    End If
    On Error GoTo 0

    If Abs(got - expect) > 0.003 Then
        Check False, label & "  - WRONG (measured " & Format(got, "0.000") & ")"
    ElseIf Len(prefix) > 0 And Trim(txt) <> Trim(prefix) Then
        Check False, label & "  - value OK, but the """ & prefix & """ prefix did not stick; add it by hand"
    Else
        Check True, label
    End If
End Sub

Sub LeaderNote(v As Object, ByVal i As Long, ByVal text As String, ByVal tx As Double, ByVal ty As Double, ByVal label As String)
    Dim nt As Object
    If i < 0 Then Check False, label & "  - MISSING (chamfer edge not found)": Exit Sub
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    SelEdge v, i, False
    Set nt = Nothing
    On Error Resume Next
    Set nt = swModel.InsertNote(text)
    If Not nt Is Nothing Then nt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
    On Error GoTo 0
    swModel.ClearSelection2 True
    Check Not nt Is Nothing, label
End Sub

Sub SelEdge(v As Object, ByVal i As Long, ByVal append As Boolean)
    Dim ok As Boolean, sd As Object
    On Error Resume Next
    ok = v.SelectEntity(eObj(i), append)
    If Not ok Then
        Err.Clear
        Set sd = swModel.SelectionManager.CreateSelectData
        Set sd.View = v
        If Err.Number <> 0 Then Err.Clear: sd.View = v
        ok = eObj(i).Select4(append, sd)
    End If
End Sub

' =====================================================================================
'  Edge table: every visible edge of a view, in sheet inches
' =====================================================================================
Sub CollectEdges(v As Object)
    Dim ents As Variant, i As Long, xf As Object, sc As Double
    eN = 0
    ReDim eObj(0 To 999): ReDim eKind(0 To 999)
    ReDim eX1(0 To 999): ReDim eY1(0 To 999): ReDim eX2(0 To 999): ReDim eY2(0 To 999)
    ReDim eCX(0 To 999): ReDim eCY(0 To 999): ReDim eR(0 To 999): ReDim eFull(0 To 999)
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
        If eN < 1000 Then AddEdge ents(i), xf, sc
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
    Dim crv As Object, cp As Variant, pt As Variant, a As Variant, b As Variant, k As Long, t As Double
    Dim isL As Boolean, isC As Boolean, cc As Variant, axv As Variant, ctr As Variant, L As Double

    On Error Resume Next
    Set crv = ent.GetCurve
    cp = ent.GetCurveParams2                          ' start(0-2) end(3-5) tStart(6) tEnd(7)
    isL = crv.IsLine
    isC = crv.IsCircle
    On Error GoTo 0
    If crv Is Nothing Or Not IsArray(cp) Then Exit Sub

    ' extents: sample the edge
    For k = 0 To 8
        pt = Empty
        On Error Resume Next
        t = cp(6) + (cp(7) - cp(6)) * k / 8
        pt = crv.Evaluate2(t, 0)
        On Error GoTo 0
        If IsArray(pt) Then Extend XfPt(xf, pt(0), pt(1), pt(2))
    Next
    a = XfPt(xf, cp(0), cp(1), cp(2))
    b = XfPt(xf, cp(3), cp(4), cp(5))
    Extend a: Extend b

    If isL Then
        L = Sqr((b(0) - a(0)) ^ 2 + (b(1) - a(1)) ^ 2)
        If L < 0.003 Then Exit Sub                    ' edge seen end-on
        Set eObj(eN) = ent: eKind(eN) = "L"
        eX1(eN) = a(0): eY1(eN) = a(1): eX2(eN) = b(0): eY2(eN) = b(1)
        eN = eN + 1
    ElseIf isC Then
        On Error Resume Next
        cc = crv.CircleParams                         ' centre(3) axis(3) radius
        On Error GoTo 0
        If Not IsArray(cc) Then Exit Sub
        axv = XfVec(xf, Arr3(cc(3), cc(4), cc(5)))
        If Abs(axv(2)) < 0.9 Then Exit Sub            ' circle seen edge-on
        ctr = XfPt(xf, cc(0), cc(1), cc(2))
        Set eObj(eN) = ent: eKind(eN) = "C"
        eCX(eN) = ctr(0): eCY(eN) = ctr(1): eR(eN) = cc(6) / IN_M * sc
        eX1(eN) = a(0): eY1(eN) = a(1): eX2(eN) = b(0): eY2(eN) = b(1)
        eFull(eN) = (Abs(cp(0) - cp(3)) + Abs(cp(1) - cp(4)) + Abs(cp(2) - cp(5)) < 0.000001)
        eN = eN + 1
    End If
End Sub

Sub Extend(p As Variant)
    If p(0) < xMin Then xMin = p(0)
    If p(0) > xMax Then xMax = p(0)
    If p(1) < yMin Then yMin = p(1)
    If p(1) > yMax Then yMax = p(1)
End Sub

' longest vertical (isVert) or horizontal line whose position is within tol of target
Function PickLine(ByVal isVert As Boolean, ByVal target As Double, ByVal tol As Double) As Long
    Dim i As Long, best As Long, bl As Double, dx As Double, dy As Double, pos As Double, L As Double
    best = -1: bl = 0
    For i = 0 To eN - 1
        If eKind(i) = "L" Then
            dx = Abs(eX2(i) - eX1(i)): dy = Abs(eY2(i) - eY1(i))
            If isVert Then
                pos = (eX1(i) + eX2(i)) / 2: L = dy
                If dx < 0.002 And L > 0.01 And Abs(pos - target) <= tol And L > bl Then best = i: bl = L
            Else
                pos = (eY1(i) + eY2(i)) / 2: L = dx
                If dy < 0.002 And L > 0.01 And Abs(pos - target) <= tol And L > bl Then best = i: bl = L
            End If
        End If
    Next
    PickLine = best
End Function

' fullMode: 1 = full circles only, 0 = arcs only, -1 = either
Function PickCircle(ByVal r As Double, ByVal tol As Double, ByVal mode As String, ByVal fullMode As Integer) As Long
    Dim i As Long, best As Long, bs As Double, s As Double
    best = -1: bs = -1E+30
    For i = 0 To eN - 1
        If eKind(i) = "C" And Abs(eR(i) - r) <= tol Then
            If (fullMode = 1 And Not eFull(i)) Or (fullMode = 0 And eFull(i)) Then GoTo NextI
            Select Case mode
                Case "minx": s = -eCX(i)
                Case "maxx": s = eCX(i)
                Case "maxy": s = eCY(i)
                Case "maxdiag": s = eCX(i) - eCY(i)
                Case Else: s = 0
            End Select
            If s > bs + 0.0001 Then best = i: bs = s
        End If
NextI:
    Next
    PickCircle = best
End Function

' the short 45 deg line of the chamfer at the flange tip (leftmost, then highest)
Function PickChamfer() As Long
    Dim i As Long, best As Long, dx As Double, dy As Double, L As Double, bx As Double, by As Double, mx As Double, my As Double
    best = -1: bx = 1E+30: by = -1E+30
    For i = 0 To eN - 1
        If eKind(i) = "L" Then
            dx = Abs(eX2(i) - eX1(i)): dy = Abs(eY2(i) - eY1(i)): L = Sqr(dx * dx + dy * dy)
            If L > 0.02 And L < 0.06 And Abs(dx - dy) < 0.01 Then
                mx = (eX1(i) + eX2(i)) / 2: my = (eY1(i) + eY2(i)) / 2
                If mx < bx - 0.01 Or (Abs(mx - bx) <= 0.01 And my > by) Then best = i: bx = mx: by = my
            End If
        End If
    Next
    PickChamfer = best
End Function

Function CX(ByVal i As Long) As Double
    If i >= 0 Then CX = eCX(i)
End Function

Function CY(ByVal i As Long) As Double
    If i >= 0 Then CY = eCY(i)
End Function

' =====================================================================================
'  Step 10 - title block
' =====================================================================================
Sub TitleBlock()
    Dim pass As Integer, sv As Object, nt As Object, t As String, pos As Variant
    Dim mX As Double, mY As Double, dX As Double, dY As Double, tX As Double, tY As Double
    Dim gotM As Boolean, gotD As Boolean, gotT As Boolean, sheetDone As Boolean, sheetSeen As Boolean

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
            Set nt = nt.GetNext
        Loop
        If pass = 2 Then swDraw.EditSheet
        On Error GoTo 0
        If gotM And gotD And gotT And sheetDone Then Exit For
    Next

    ' fall back to the reference layout if a label was not found
    If Not gotM Then mX = 5.07: mY = 0.94
    If Not gotD Then dX = 3.17: dY = 0.48
    If Not gotT Then tX = 7.2: tY = 0.92
    Check gotM And gotD And gotT, "Title block labels found (MATERIAL, DATE, TITLE)" & _
          IIf(gotM And gotD And gotT, "", " - used the reference positions instead")

    AddSheetNote "316 STAINLESS STEEL", mX + 0.06, mY - 0.17, "Material: 316 STAINLESS STEEL"
    AddSheetNote Format(Date, "mm-dd-yyyy"), dX + 0.4, dY - 0.18, "Date: " & Format(Date, "mm-dd-yyyy")
    AddSheetNote "IME 1140 ENGINEERING DRAWINGS I", tX + 0.44, tY, "Title line 1"
    AddSheetNote "IN-LAB ASSIGNMENT #7A", tX + 0.44, tY - 0.17, "Title line 2"
    Check sheetDone, "SHEET 1 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
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

' =====================================================================================
'  Step 12 - save the drawing and the PDF
' =====================================================================================
Sub SaveAll()
    Dim lastName As String, folder As String, drw As String, pdf As String, ok As Boolean, okPdf As Boolean
    lastName = Trim(InputBox("Last name for the file names  HW#7A<LastName>.SLDDRW  and  .PDF" & vbCrLf & _
                             "(leave empty to skip saving)", "Save HW#7A", "Perhach"))
    If Len(lastName) = 0 Then WriteReport SaveFolder() & "\": Exit Sub
    folder = Left(partPath, InStrRev(partPath, "\"))
    drw = folder & "HW#7A" & lastName & ".SLDDRW"
    ok = TrySave(drw)
    If Not ok Then
        folder = SaveFolder() & "\"
        drw = folder & "HW#7A" & lastName & ".SLDDRW"
        ok = TrySave(drw)
    End If
    If Not ok Then
        MsgBox "Could not save the drawing automatically. Use File > Save As and name it  HW#7A" & lastName & _
               ", then File > Save As > PDF with the same name.", vbExclamation, "HW#7A"
        Exit Sub
    End If
    pdf = folder & "HW#7A" & lastName & ".PDF"
    okPdf = TrySavePdf(pdf)
    WriteReport folder
    MsgBox "Saved:" & vbCrLf & drw & vbCrLf & IIf(okPdf, pdf, "PDF NOT saved - use File > Save As > PDF"), _
           IIf(okPdf, vbInformation, vbExclamation), "HW#7A"
End Sub

Sub WriteReport(ByVal folder As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open folder & "HW7A_checks.txt" For Output As #f
    Print #f, "HW#7A drawing checks  (" & Now & ")"
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

' =====================================================================================
'  Report
' =====================================================================================
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
'  Geometry helpers
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
