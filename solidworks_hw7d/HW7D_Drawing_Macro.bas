' =====================================================================================
'  HW#7D  -  IME 1140 In-Lab Assignment #7D  -  SolidWorks DRAWING macro   (inches, sheet 1:1)
'
'  HOW TO RUN
'   1. Open "IME 1140 Engineering Drawings I In-Lab Assignment Part #4.SLDPRT" in SolidWorks.
'   2. Tools > Macro > New...  save it as  HW7D_Drawing.swp
'   3. Select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main.  At the end enter your last name; it saves  HW#7D<LastName>.SLDDRW  and  .PDF.
'
'  WHAT IT BUILDS (matches the reference sheet)
'   - Main view of the pin (big body on the left), DETAIL E at 2:1 around the groove, isometric view
'   - Main:     .075 X 45 deg, dia .925, 2.207, .685, 4.000, dia .625, dia .75
'   - Detail E: .07, .16, 30 deg, dia .650, dia .492, dia .625, .100, .050
'   - Title block: C36000 BRASS, today's date, title, SHEET 4 OF 4
'
'  SAFETY
'   - Part orientation read from the dia .925 body and the dia .625 pin.
'   - Edges picked from exact positions; cylinder outlines first, side-on circle edges as a fallback.
'   - Every dimension read back and checked (wrong ones are deleted and retried another way).
'   - The part file is never saved or changed by this macro.
' =====================================================================================
Option Explicit

Dim swApp As Object, swMath As Object, swPart As Object, swModel As Object, swDraw As Object
Dim partPath As String, curStep As String, report As String, problems As String
Dim nOK As Long, nBad As Long

Const IN_M = 0.0254
Const DRAW_NAME = "HW#7D"
Const PART_TAG = "PART #4"
Const SHEET_DEN = 1#             ' sheet scale 1 : 1

' ---- layout (sheet inches, from the reference) ----
Const MAIN_X = 2.61              ' left end of the pin
Const MAIN_Y = 3.13              ' centre line of the pin
Const DET_X = 3.81               ' detail E: end of the dia .75 neck
Const DET_Y = 5.72               ' detail E: centre line
Const ISO_X = 8.87
Const ISO_Y = 5.64

Const DM_HIDDEN_VISIBLE = 1
Const DM_HIDDEN_REMOVED = 2

' ---- axial positions along the pin (inches from the left end) ----
Const X_SHOULDER = 2.207         ' dia .925 -> dia .75
Const X_NECK = 2.892             ' end of the dia .75 neck (2.207 + .685), start of the groove
Const X_RINGL = 2.942            ' groove .050 wide
Const X_RINGR = 2.992            ' ring .050 wide
Const X_UC1 = 3.062              ' .07 from the ring: start of the dia .492 undercut
Const X_UC2 = 3.152              ' .16 from the ring: end of the undercut, start of the 30 deg chamfer
Const X_PIN = 3.1904             ' chamfer reaches dia .625

Dim vR As Variant                ' along the pin axis, body -> pin end   (right in the main view)
Dim vU As Variant                ' a radial direction                     (up in the main view)
Dim vF As Variant                ' toward you in the main view
Dim pOrg As Variant              ' model point on the axis at the left end

Dim viewMain As Object, viewDet As Object, viewIso As Object

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
    report = "": problems = "": nOK = 0: nBad = 0

    curStep = "Step 1 - find the open part":           FindPart
    curStep = "Step 2 - read the part orientation":    FindOrientation
    curStep = "Step 3 - find your template":           tmpl = FindTemplate()
    curStep = "Step 4 - new drawing":                  NewDrawing tmpl
    curStep = "Step 5 - views and DETAIL E":           MakeViews
    curStep = "Step 6 - centrelines":                  CentreLines viewMain: CentreLines viewDet
    curStep = "Step 7 - main view dimensions":         DimMain
    curStep = "Step 8 - DETAIL E dimensions":          DimDetail
    curStep = "Step 9 - title block":                  TitleBlock
    curStep = "Step 10 - finish":                      FinishSheet

    MsgBox DRAW_NAME & " drawing built." & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & vbCrLf & vbCrLf & _
           IIf(nBad = 0, "Every view, dimension and title block item checked out.", "Problems:" & vbCrLf & Left(problems, 700)) & _
           vbCrLf & vbCrLf & "The full check list is saved as HW7D_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), DRAW_NAME

    curStep = "Step 11 - save drawing and PDF":        SaveAll
    Exit Sub
Fail:
    MsgBox "The macro stopped at: " & curStep & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Checks passed so far: " & nOK & vbCrLf & _
           IIf(nBad = 0, "", "Problems so far:" & vbCrLf & Left(problems, 500)), vbCritical, DRAW_NAME & " macro"
End Sub

' =====================================================================================
'  Part, template, new drawing
' =====================================================================================
Sub FindPart()
    ' Use the open part whose name says Part #4. Otherwise the active part, otherwise the only part.
    Dim d As Object, docs As Variant, i As Long, t As String, nParts As Long, firstPart As Object
    Set swPart = Nothing
    On Error Resume Next
    docs = swApp.GetDocuments
    On Error GoTo 0
    If IsArray(docs) Then
        For i = LBound(docs) To UBound(docs)
            If docs(i).GetType = 1 Then
                nParts = nParts + 1
                If firstPart Is Nothing Then Set firstPart = docs(i)
                t = UCase(docs(i).GetTitle & " " & docs(i).GetPathName)
                If InStr(t, PART_TAG) > 0 Or InStr(t, Replace(PART_TAG, " #", "_")) > 0 Or InStr(t, Replace(PART_TAG, "#", "")) > 0 Then
                    Set swPart = docs(i)
                End If
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
        "Could not tell which open part is In-Lab Assignment Part #4. Close the other parts, click into the Part #4 window, and run again."
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then Err.Raise vbObjectError + 2, , "The part has never been saved. Save it first."
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

Sub FindOrientation()
    ' dia .925 body (r .4625): its axis is the pin axis.  dia .625 pin (r .3125) tells which end is right.
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, k As Long
    Dim s As Object, cp As Variant, bx As Variant, r As Double
    Dim bAx As Variant, bC(2) As Double, bN As Long, a0 As Variant, pC(2) As Double, pN As Long
    Dim bb As Variant, c As Variant, ix As Integer, iy As Integer, iz As Integer, t As Double
    Dim rMin As Double, rMax As Double, cand As Variant

    bodies = swPart.GetBodies2(0, True)
    If Not IsArray(bodies) Then Err.Raise vbObjectError + 3, , "The part has no solid body."
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder Then
                cp = s.CylinderParams
                r = cp(6) / IN_M
                bx = faces(i).GetBox
                If Abs(r - 0.4625) < 0.004 Then
                    bAx = Canon(Unit(Arr3(cp(3), cp(4), cp(5)))): a0 = Arr3(cp(0), cp(1), cp(2)): bN = bN + 1
                    For k = 0 To 2: bC(k) = bC(k) + (bx(k) + bx(k + 3)) / 2: Next
                ElseIf Abs(r - 0.3125) < 0.004 Then
                    pN = pN + 1
                    For k = 0 To 2: pC(k) = pC(k) + (bx(k) + bx(k + 3)) / 2: Next
                End If
            End If
        Next
    Next
    If bN = 0 Then Err.Raise vbObjectError + 4, , "This part has no dia .925 body, so it is not In-Lab Assignment Part #4. Click into the Part #4 window and run again."
    If pN = 0 Then Err.Raise vbObjectError + 4, , "Could not find the dia .625 pin in the part."
    For k = 0 To 2: bC(k) = bC(k) / bN: pC(k) = pC(k) / pN: Next

    vR = bAx                                                      ' pin end is to the right of the body
    If (pC(0) - bC(0)) * vR(0) + (pC(1) - bC(1)) * vR(1) + (pC(2) - bC(2)) * vR(2) < 0 Then vR = Scl(vR, -1)
    ' "up": a model axis square to the pin axis (prefers +Y, then +Z, then +X) so standard views can be used
    vU = Empty
    For Each cand In Array(Arr3(0, 1, 0), Arr3(0, 0, 1), Arr3(1, 0, 0))
        If IsEmpty(vU) Then
            If Abs(Dot(cand, vR)) < 0.01 Then vU = cand
        End If
    Next
    If IsEmpty(vU) Then vU = Unit(Cross(vR, Arr3(0, 0, 1)))
    vF = Cross(vR, vU)

    bb = bodies(LBound(bodies)).GetBodyBox
    rMin = 1E+30: rMax = -1E+30
    For ix = 0 To 1
        For iy = 0 To 1
            For iz = 0 To 1
                c = Arr3(bb(3 * ix), bb(1 + 3 * iy), bb(2 + 3 * iz))
                t = Dot(c, vR): If t < rMin Then rMin = t
                If t > rMax Then rMax = t
            Next
        Next
    Next
    pOrg = Add3(a0, Scl(vR, rMin - Dot(a0, vR)))                  ' on the axis, at the left end
    Check (Abs((rMax - rMin) / IN_M - 4#) < 0.01), "Part length 4.000 (measured " & Format((rMax - rMin) / IN_M, "0.000") & ")"
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
                 "then paste it here:", "HW#7D - drawing template")
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
'  Views and DETAIL E
' =====================================================================================
Sub MakeViews()
    Dim dv As Variant, upI As Variant, rtI As Variant

    Set viewMain = OrientedView("HW7D_Main", vR, vU, vF, MAIN_X + 2, MAIN_Y, StdFrontName(), False)
    If viewMain Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented main view."
    Check True, "Main view: body on the left, pin end on the right"
    SetMode viewMain, DM_HIDDEN_REMOVED, "Main"
    PlaceView viewMain, 0#, 0#, 0#, MAIN_X, MAIN_Y, True, True, "Main view placed (left end, centre line)"

    Set viewDet = MakeDetail()
    If viewDet Is Nothing Then Err.Raise vbObjectError + 11, , "Could not create DETAIL E."
    SetMode viewDet, DM_HIDDEN_REMOVED, "DETAIL E"
    PlaceView viewDet, X_NECK, 0#, 0#, DET_X, DET_Y, True, True, "DETAIL E placed (groove, centre line)"

    dv = Unit(Add3(Add3(vR, vU), vF))
    upI = Unit(Sub3(vU, Scl(dv, Dot(vU, dv))))
    rtI = Cross(upI, dv)
    Set viewIso = OrientedView("HW7D_Iso", rtI, upI, dv, ISO_X, ISO_Y, StdIsoName(), True)
    If viewIso Is Nothing Then
        Check False, "Isometric view - could not create it; add it by hand (View Palette > Isometric)"
    Else
        Check True, "Isometric view"
        SetMode viewIso, DM_HIDDEN_REMOVED, "Isometric"
    End If
End Sub

Function StdIsoName() As String
    If Near(vR, 1, 0, 0) And Near(vU, 0, 1, 0) And Near(vF, 0, 0, 1) Then StdIsoName = "*Isometric"
End Function

' DETAIL E: sketch a circle round the groove in the main view, then make the 2:1 detail from it
Function MakeDetail() As Object
    Dim c As Variant, rad As Double, cand As Variant, seg As Object, dv As Object, before As String, sc As Double

    c = SheetPt(viewMain, 3.06, 0#, 0#)                            ' circle centre (groove), sheet inches
    sc = 1#: On Error Resume Next: sc = viewMain.ScaleDecimal: On Error GoTo 0
    If sc <= 0 Then sc = 1#
    rad = 0.56 * sc                                                ' circle radius on the sheet
    before = ViewNames()
    For Each cand In Array(2, 3, 1)                                ' the view sketch's coordinates: try each way
        Set seg = DrawViewCircle(viewMain, c, rad, CInt(cand))
        If Not seg Is Nothing Then
            swModel.ClearSelection2 True
            On Error Resume Next
            seg.Select4 False, Nothing
            Set dv = Nothing
            Set dv = swDraw.CreateDetailViewAt4(DET_X * IN_M + 1.5 * IN_M, DET_Y * IN_M, 0#, 0, 2#, 1#, "E", 1, False, False, False, 0)
            If dv Is Nothing Then Err.Clear: Set dv = swDraw.CreateDetailViewAt3(DET_X * IN_M + 1.5 * IN_M, DET_Y * IN_M, 0#, 0, 2#, 1#, "E", 1, False)
            On Error GoTo 0
            If dv Is Nothing Then Set dv = NewViewSince(before)           ' built but not handed back
            If Not dv Is Nothing Then Exit For
            DeleteSeg seg
        End If
    Next
    swModel.ClearSelection2 True
    If dv Is Nothing Then
        Check False, "DETAIL E: SolidWorks would not make the detail view"
        Exit Function
    End If
    Check True, "DETAIL E created from a circle round the groove (method " & cand & ")"
    sc = 0: On Error Resume Next: sc = dv.ScaleDecimal: On Error GoTo 0
    Check (Abs(sc - 2#) < 0.01), "DETAIL E scale 2:1" & IIf(Abs(sc - 2#) < 0.01, "", " (it is " & Format(sc, "0.##") & ") - set the view scale to 2:1 by hand")
    Set MakeDetail = dv
End Function

' sketch a circle in a drawing view; c = centre (sheet inches), rad = radius (sheet inches)
Function DrawViewCircle(v As Object, c As Variant, ByVal rad As Double, ByVal cand As Integer) As Object
    Dim x As Double, y As Double, rr As Double, o As Variant, sc As Double, pos As Variant
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    sc = 1#
    On Error Resume Next
    sc = v.ScaleDecimal
    If sc <= 0 Then sc = 1#
    Select Case cand
        Case 1
            x = c(0) * IN_M: y = c(1) * IN_M: rr = rad * IN_M
        Case 2
            o = XfPt(v.ModelToViewTransform, 0#, 0#, 0#)
            x = (c(0) - o(0)) * IN_M / sc: y = (c(1) - o(1)) * IN_M / sc: rr = rad * IN_M / sc
        Case 3
            pos = v.Position
            x = (c(0) * IN_M - pos(0)) / sc: y = (c(1) * IN_M - pos(1)) / sc: rr = rad * IN_M / sc
    End Select
    swModel.SketchManager.AddToDB = True
    Set DrawViewCircle = swModel.SketchManager.CreateCircleByRadius(x, y, 0#, rr)
    swModel.SketchManager.AddToDB = False
    swModel.GraphicsRedraw2
    swModel.ClearSelection2 True
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
    For k = 1 To 3
        cur = SheetPt(v, r, u, f)
        ok = (Not doX Or Abs(cur(0) - tx) < 0.02) And (Not doY Or Abs(cur(1) - ty) < 0.02)
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
    ok = (Not doX Or Abs(cur(0) - tx) < 0.02) And (Not doY Or Abs(cur(1) - ty) < 0.02)
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
'  Centrelines and dimensions
' =====================================================================================
Sub CentreLines(v As Object)
    On Error Resume Next
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 v.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    swDraw.InsertCenterLine2
    swModel.ClearSelection2 True
End Sub

Sub DimMain()
    Dim xL As Double, xR As Double, yA As Double, xN As Double
    CollectEdges viewMain
    xL = SX(viewMain, 0): xR = SX(viewMain, 4#): xN = SX(viewMain, X_NECK): yA = SY(viewMain, 0)
    ' .075 X 45 deg: leader to the chamfer outline, else to the left end edge
    LeaderNote2 viewMain, Array(0.0375, 0.425, 0#), Array(0#, 0.3, 0#), ".075 X 45<MOD-DEG>", xL - 1.28, yA + 0.76, "Main     .075 X 45" & ChrW(176)
    DiaDim viewMain, 1.1, 0.4625, X_SHOULDER, 0.42, xL - 0.65, yA, 0.925, "Main     " & ChrW(216) & ".925", 3
    MakeDimP viewMain, "H", Array(0#, 0.1, 0#), Array(X_SHOULDER, 0.42, 0#), (xL + SX(viewMain, X_SHOULDER)) / 2, yA - 0.8, 2.207, "Main     2.207", "", 3
    MakeDimP viewMain, "H", Array(X_SHOULDER, 0.42, 0#), Array(X_NECK, 0.36, 0#), xN + 0.38, yA - 0.8, 0.685, "Main     .685", "", 3
    MakeDimP viewMain, "H", Array(0#, 0.1, 0#), Array(4#, 0.1, 0#), (xL + xR) / 2, yA - 1.15, 4#, "Main     4.000", "", 3
    DiaDim viewMain, 3.6, 0.3125, X_PIN, 0.2, xR + 0.64, yA, 0.625, "Main     " & ChrW(216) & ".625", 3
    DiaDim viewMain, 2.5, 0.375, X_NECK, 0.36, xR + 1.19, yA, 0.75, "Main     " & ChrW(216) & ".75", 2
End Sub

Sub DimDetail()
    Dim xRR As Double, xN As Double, xU2 As Double, yA As Double
    CollectEdges viewDet
    xRR = SX(viewDet, X_RINGR): xN = SX(viewDet, X_NECK): xU2 = SX(viewDet, X_UC2): yA = SY(viewDet, 0)
    MakeDimP viewDet, "H", Array(X_RINGR, 0.36, 0#), Array(X_UC1, 0.28, 0#), xRR - 0.42, yA + 1.47, 0.07, "Detail E .07", "", 2
    MakeDimP viewDet, "H", Array(X_RINGR, 0.36, 0#), Array(X_UC2, 0.2, 0#), xRR - 0.44, yA + 1.14, 0.16, "Detail E .16", "", 2
    MakeDimP viewDet, "A", Array(X_RINGR, 0.36, 0#), Array((X_UC2 + X_PIN) / 2, 0.279, 0#), xRR + 0.55, yA + 1.72, 30#, "Detail E 30" & ChrW(176)
    DiaDim viewDet, (X_NECK + X_RINGL) / 2, 0.325, -1, 0, xN - 1.64, yA, 0.65, "Detail E " & ChrW(216) & ".650", 3
    DiaDim viewDet, (X_UC1 + X_UC2) / 2, 0.246, X_UC2, 0.1, xU2 + 1.89, yA, 0.492, "Detail E " & ChrW(216) & ".492", 3
    DiaDim viewDet, X_PIN + 0.15, 0.3125, X_PIN, 0.2, xU2 + 2.6, yA, 0.625, "Detail E " & ChrW(216) & ".625", 3
    MakeDimP viewDet, "H", Array(X_NECK, 0.36, 0#), Array(X_RINGR, 0.36, 0#), xN - 0.46, yA - 0.99, 0.1, "Detail E .100", "", 3
    MakeDimP viewDet, "H", Array(X_NECK, 0.36, 0#), Array(X_RINGL, 0.36, 0#), xN - 0.46, yA - 1.2, 0.05, "Detail E .050", "", 3
End Sub

' diameter: first the two cylinder outlines at axial position xs (radius rr); if that does not measure
' right, the circle edge seen side-on at xe (picked at radius re) - when xe >= 0
Sub DiaDim(v As Object, ByVal xs As Double, ByVal rr As Double, ByVal xe As Double, ByVal re As Double, _
           ByVal tx As Double, ByVal ty As Double, ByVal expect As Double, ByVal label As String, ByVal prec As Long)
    Dim why As String
    why = TryDim(v, "V", Array(xs, rr, 0#), Array(xs, -rr, 0#), tx, ty, expect, "<MOD-DIAM>", prec, "", False, True)
    If Len(why) > 0 And xe >= 0 Then
        why = TryDim(v, "VL", Array(xe, re, 0#), Empty, tx, ty, expect, "<MOD-DIAM>", prec, "", False, True)
    End If
    Check (Len(why) = 0), label & IIf(Len(why) = 0, "", "  - " & why)
End Sub

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
    Dim dd As Object, got As Double, nSel As Long, need As Long, txt As String, ok As Boolean
    If kind = "H" Or kind = "V" Or kind = "S" Or kind = "A" Then need = 2 Else need = 1
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    SelPt v, pa, False
    If need = 2 Then SelPt v, pb, True
    nSel = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    If nSel <> need Then
        swModel.ClearSelection2 True
        TryDim = "MISSING (found " & nSel & " of " & need & " edge(s))"
        Exit Function
    End If
    Set dd = Nothing
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
    If dd Is Nothing Then TryDim = "MISSING (SolidWorks did not create it)": Exit Function

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
        TryDim = "WRONG (measured " & Format(got, "0.000") & ")"
        If deleteIfWrong Then
            On Error Resume Next
            swModel.ClearSelection2 True
            dd.GetAnnotation.Select3 False, Nothing
            swModel.EditDelete
            swModel.ClearSelection2 True
            On Error GoTo 0
        End If
    ElseIf Len(prefix) > 0 And InStr(txt, prefix) = 0 Then
        TryDim = "value OK, but the " & IIf(prefix = "<MOD-DIAM>", ChrW(216) & " symbol", "prefix") & " did not stick; add it by hand"
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
'  Title block, finish, save
' =====================================================================================
Sub TitleBlock()
    Dim pass As Integer, sv As Object, nt As Object, t As String, pos As Variant
    Dim mX As Double, mY As Double, dX As Double, dY As Double, tX As Double, tY As Double
    Dim gotM As Boolean, gotD As Boolean, gotT As Boolean, sheetDone As Boolean, sheetSeen As Boolean, scaleDone As Boolean

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
            If Left(t, 6) = "SCALE:" And Not scaleDone Then              ' sheet is 1:1
                If InStr(t, "1:1") = 0 Then nt.SetText "SCALE: 1:1"
                If InStr(UCase(nt.GetText), "1:1") = 0 Then nt.PropertyLinkedText = "SCALE: 1:1"
                scaleDone = (InStr(UCase(nt.GetText), "1:1") > 0)
            End If
            Set nt = nt.GetNext
        Loop
        If pass = 2 Then swDraw.EditSheet
        On Error GoTo 0
        If gotM And gotD And gotT And sheetDone And scaleDone Then Exit For
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

    AddSheetNote "C36000 BRASS", mX + 0.06, mY - 0.17, "Material: C36000 BRASS"
    AddSheetNote Format(Date, "mm-dd-yyyy"), dX + 0.4, dY - 0.18, "Date: " & Format(Date, "mm-dd-yyyy")
    AddSheetNote "IME 1140 ENGINEERING DRAWINGS I", tX + 0.44, tY, "Title line 1"
    AddSheetNote "IN-LAB ASSIGNMENT #7D", tX + 0.44, tY - 0.17, "Title line 2"
    Check scaleDone, "SCALE: 1:1" & IIf(scaleDone, "", " - edit the scale note by hand")
    Check sheetDone, "SHEET 4 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
End Sub

Sub FinishSheet()
    ' leave the drawing active, in normal sheet mode, with all four views shown and redrawn
    Dim vs As Variant, n As Long, i As Long, v As Variant
    On Error Resume Next
    ActivateDrawing
    swDraw.EditSheet
    swModel.ClearSelection2 True
    For Each v In Array(viewMain, viewDet, viewIso)
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
    lastName = Trim(InputBox("Last name for the file names  HW#7D<LastName>.SLDDRW  and  .PDF" & vbCrLf & _
                             "(leave empty to skip saving)", "Save HW#7D", "Perhach"))
    If Len(lastName) = 0 Then WriteReport SaveFolder() & "\": Exit Sub
    folder = Left(partPath, InStrRev(partPath, "\"))
    drw = folder & "HW#7D" & lastName & ".SLDDRW"
    ok = TrySave(drw)
    If Not ok Then
        folder = SaveFolder() & "\"
        drw = folder & "HW#7D" & lastName & ".SLDDRW"
        ok = TrySave(drw)
    End If
    If Not ok Then
        MsgBox "Could not save the drawing automatically. Use File > Save As and name it  HW#7D" & lastName & _
               ", then File > Save As > PDF with the same name.", vbExclamation, "HW#7D"
        Exit Sub
    End If
    pdf = folder & "HW#7D" & lastName & ".PDF"
    FinishSheet                                         ' views visible + drawing active before the PDF export
    okPdf = TrySavePdf(pdf)
    WriteReport folder
    MsgBox "Saved:" & vbCrLf & drw & vbCrLf & IIf(okPdf, pdf, "PDF NOT saved - use File > Save As > PDF"), _
           IIf(okPdf, vbInformation, vbExclamation), "HW#7D"
End Sub

Sub WriteReport(ByVal folder As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open folder & "HW7D_checks.txt" For Output As #f
    Print #f, "HW#7D drawing checks  (" & Now & ")"
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
