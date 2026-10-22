' =====================================================================================
'  HW#7B  -  IME 1140 In-Lab Assignment #7B  -  SolidWorks DRAWING macro   (inches)
'
'  HOW TO RUN
'   1. Open "IME 1140 Engineering Drawings I In-Lab Assignment Part #2.SLDPRT" in SolidWorks.
'   2. Tools > Macro > New...  save it as  HW7B_Drawing.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main  (Tools > Macro > Run, select "main", Run).
'   5. It builds the drawing on your template, checks everything, asks for your last name and
'      saves  HW#7B<LastName>.SLDDRW  and  HW#7B<LastName>.PDF  next to the part.
'
'  WHAT IT BUILDS (matches the reference sheet)
'   - Plan view (looking down on the bosses, slot on the left, R1.00 end on the right),
'     hidden lines visible.  Right view projected from it (bosses pointing left).
'   - SECTION B-B below the plan view, cut along the centre line, arrows pointing up.
'   - Plan:    R.25 TYP, 4 X R.125, R1.00, 1.00, .125 X 45 deg, .375, .875, 3.25, dia .375
'   - Section: 1.500, 4 X .125, .375, 1.25, dia .75, 1.075, .30, 1.625, dia .290, .75, 1.025
'   - Right:   2.000, 1.250, (.375), .60, .025 X 45 deg TYP
'   - Title block: 6061-T6 ALUMINUM, today's date, title, SHEET 2 OF 4.
'
'  SAFETY
'   - The part's orientation is read from its geometry (bosses, R1.00 end, cross hole).
'   - Every view's orientation is verified; every dimension value is read back and checked.
'   - Edges are picked from their exact positions on the part, so the right edge is always used.
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
Const DRAW_NAME = "HW#7B"        ' file name prefix

' ---- sheet layout, view centres in inches from the sheet's lower-left corner ----
Const PAR_X = 3.7                ' plan view
Const PAR_Y = 6.01
Const SEC_Y = 2.52               ' section B-B (aligned under the plan view)
Const RIGHT_X = 8.95             ' right view (aligned with the plan view)

' ---- display modes (swDisplayMode_e) ----
Const DM_HIDDEN_VISIBLE = 1
Const DM_HIDDEN_REMOVED = 2

' ---- the part's own directions (model coordinates, unit vectors) ----
Dim vR As Variant                ' along the length: slot end -> R1.00 end   (right in the plan view)
Dim vU As Variant                ' across the width                         (up in the plan view)
Dim vF As Variant                ' bottom face -> top of the bosses          (toward you in the plan view)
Dim pOrg As Variant              ' model point (metres): left end, centre line, bottom face

Dim viewPar As Object, viewSec As Object, viewRight As Object

' ---- edge table of the view being dimensioned (sheet inches) ----
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

    curStep = "Step 1 - find the open part":          FindPart
    curStep = "Step 2 - read the part orientation":   FindOrientation
    curStep = "Step 3 - find your template":          tmpl = FindTemplate()
    curStep = "Step 4 - new drawing":                 NewDrawing tmpl
    curStep = "Step 5 - views and section B-B":       MakeViews
    curStep = "Step 6 - centre marks and centrelines": CentreMarksAndLines
    curStep = "Step 7 - plan view dimensions":        DimPlan
    curStep = "Step 8 - section B-B dimensions":      DimSection
    curStep = "Step 9 - right view dimensions":       DimRight
    curStep = "Step 10 - title block":                TitleBlock
    curStep = "Step 11 - finish":                     FinishSheet

    MsgBox DRAW_NAME & " drawing built." & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & vbCrLf & vbCrLf & _
           IIf(nBad = 0, "Every view, dimension and title block item checked out.", "Problems:" & vbCrLf & Left(problems, 700)) & _
           vbCrLf & vbCrLf & "The full check list is saved as HW7B_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), DRAW_NAME

    curStep = "Step 12 - save drawing and PDF":       SaveAll
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
        "No part is open. Open the In-Lab Assignment Part #2 file first, then run the macro again."
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then Err.Raise vbObjectError + 2, , _
        "The part has never been saved, so a drawing cannot reference it. Save the part first."
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

Sub FindOrientation()
    ' Ø.75 boss  (r .375): its axis is "toward you" in the plan view, pointing up out of the plate
    ' R1.00 end  (r 1.00): tells which end is on the right
    ' Ø.29 cross hole (r .145): its axis runs across the width
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, k As Long
    Dim s As Object, cp As Variant, bx As Variant, ax As Variant, r As Double
    Dim bossAx As Variant, bossC(2) As Double, bossN As Long
    Dim endC(2) As Double, endN As Long
    Dim xAx As Variant, xC(2) As Double, xN As Long
    Dim bb As Variant, c As Variant, ix As Integer, iy As Integer, iz As Integer
    Dim rMin As Double, rMax As Double, uMin As Double, uMax As Double, fMin As Double, fMax As Double, t As Double

    bodies = swPart.GetBodies2(0, True)
    If Not IsArray(bodies) Then Err.Raise vbObjectError + 3, , "The part has no solid body."
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder Then
                cp = s.CylinderParams
                r = cp(6) / IN_M
                ax = Canon(Unit(Arr3(cp(3), cp(4), cp(5))))
                bx = faces(i).GetBox
                If Abs(r - 0.375) < 0.004 Then
                    bossAx = ax: bossN = bossN + 1
                    For k = 0 To 2: bossC(k) = bossC(k) + (bx(k) + bx(k + 3)) / 2: Next
                ElseIf Abs(r - 1#) < 0.005 Then
                    endN = endN + 1
                    For k = 0 To 2: endC(k) = endC(k) + (bx(k) + bx(k + 3)) / 2: Next
                ElseIf Abs(r - 0.145) < 0.004 Then
                    xAx = ax: xN = xN + 1
                    For k = 0 To 2: xC(k) = xC(k) + (bx(k) + bx(k + 3)) / 2: Next
                End If
            End If
        Next
    Next
    If bossN = 0 Then Err.Raise vbObjectError + 4, , "Could not find the dia .75 boss in the part."
    If endN = 0 Then Err.Raise vbObjectError + 4, , "Could not find the R1.00 end in the part."
    If xN = 0 Then Err.Raise vbObjectError + 4, , "Could not find the dia .29 cross hole in the part."
    For k = 0 To 2
        bossC(k) = bossC(k) / bossN: endC(k) = endC(k) / endN: xC(k) = xC(k) / xN
    Next

    vF = bossAx                                                  ' boss is above the cross hole
    If (bossC(0) - xC(0)) * vF(0) + (bossC(1) - xC(1)) * vF(1) + (bossC(2) - xC(2)) * vF(2) < 0 Then vF = Scl(vF, -1)
    vR = Unit(Cross(xAx, vF))                                    ' R1.00 end is right of the cross hole
    If (endC(0) - xC(0)) * vR(0) + (endC(1) - xC(1)) * vR(1) + (endC(2) - xC(2)) * vR(2) < 0 Then vR = Scl(vR, -1)
    vU = Cross(vF, vR)

    ' part-frame origin: left end, centre line, bottom face (from the body's bounding box)
    bb = bodies(LBound(bodies)).GetBodyBox
    rMin = 1E+30: rMax = -1E+30: uMin = 1E+30: uMax = -1E+30: fMin = 1E+30: fMax = -1E+30
    For ix = 0 To 1
        For iy = 0 To 1
            For iz = 0 To 1
                c = Arr3(bb(3 * ix), bb(1 + 3 * iy), bb(2 + 3 * iz))
                t = Dot(c, vR): If t < rMin Then rMin = t
                If t > rMax Then rMax = t
                t = Dot(c, vU): If t < uMin Then uMin = t
                If t > uMax Then uMax = t
                t = Dot(c, vF): If t < fMin Then fMin = t
                If t > fMax Then fMax = t
            Next
        Next
    Next
    pOrg = Add3(Add3(Scl(vR, rMin), Scl(vU, (uMin + uMax) / 2)), Scl(vF, fMin))
    Check (Abs((rMax - rMin) / IN_M - 4.25) < 0.01 And Abs((uMax - uMin) / IN_M - 2#) < 0.01 And Abs((fMax - fMin) / IN_M - 1.075) < 0.01), _
          "Part size 4.25 x 2.00 x 1.075 (measured " & Format((rMax - rMin) / IN_M, "0.000") & " x " & _
          Format((uMax - uMin) / IN_M, "0.000") & " x " & Format((fMax - fMin) / IN_M, "0.000") & ")"
End Sub

' part coordinates (inches: r along the length from the left end, u across from the centre line,
' f up from the bottom face) -> model point (metres)
Function PartPt(ByVal r As Double, ByVal u As Double, ByVal f As Double) As Variant
    PartPt = Add3(pOrg, Add3(Add3(Scl(vR, r * IN_M), Scl(vU, u * IN_M)), Scl(vF, f * IN_M)))
End Function

' part coordinates -> sheet point (inches) in view v
Function SheetPt(v As Object, ByVal r As Double, ByVal u As Double, ByVal f As Double) As Variant
    Dim p As Variant
    p = PartPt(r, u, f)
    SheetPt = XfPt(v.ModelToViewTransform, p(0), p(1), p(2))
End Function

' outline of the view on the sheet, from the part's 4.25 x 2.00 x 1.075 box
Sub ViewExtents(v As Object)
    Dim r As Variant, u As Variant, f As Variant, s As Variant
    xMin = 1E+30: xMax = -1E+30: yMin = 1E+30: yMax = -1E+30
    For Each r In Array(0#, 4.25)
        For Each u In Array(-1#, 1#)
            For Each f In Array(0#, 1.075)
                s = SheetPt(v, CDbl(r), CDbl(u), CDbl(f))
                Extend s
            Next
        Next
    Next
End Sub

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
                 "then paste it here:", "HW#7B - drawing template")
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
'  Views and section B-B
' =====================================================================================
Sub MakeViews()
    Set viewPar = OrientedView("HW7B_Plan", vR, vU, vF, PAR_X, PAR_Y, StdFrontName(), False)
    If viewPar Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented plan view."
    Check True, "Plan view: bosses toward you, slot on the left, R1.00 end on the right"
    SetMode viewPar, DM_HIDDEN_VISIBLE, "Plan"

    Set viewRight = Projected(viewPar, RIGHT_X, PAR_Y)
    If viewRight Is Nothing Then Err.Raise vbObjectError + 10, , "Could not create the projected right view."
    Check FrameOK(viewRight, vF, -1, 0, vU, 0, 1), "Right view: bosses pointing left (third angle)"
    SetMode viewRight, DM_HIDDEN_VISIBLE, "Right"

    Set viewSec = MakeSection(viewPar, PAR_X, SEC_Y)
    If viewSec Is Nothing Then Err.Raise vbObjectError + 11, , "Could not create SECTION B-B."
    SetMode viewSec, DM_HIDDEN_REMOVED, "Section B-B"
End Sub

' cutting line along the centre line of the plan view, then the section view under it
Function MakeSection(parent As Object, ByVal px As Double, ByVal py As Double) As Object
    Dim a As Variant, b As Variant, tp As Variant, cand As Integer, seg As Object, ok As Boolean
    Dim opt As Integer, sv As Object, sec As Object

    a = SheetPt(parent, -0.15, 0, 0.6)            ' starts just left of the part ...
    b = SheetPt(parent, 4.49, 0, 0.6)             ' ... ends just right of the R1.00 end
    tp = SheetPt(parent, -0.08, 0, 0.6)           ' a point on the line outside the part (nothing else there)

    For cand = 1 To 3                             ' the view sketch's coordinate system is tried, then verified
        Set seg = DrawViewLine(parent, a, b, cand)
        If SketchLineAt(tp) Then ok = True: Exit For
        DeleteSeg seg
    Next
    Check ok, "Cutting line B-B along the centre line"
    If Not ok Then Exit Function

    For opt = 0 To 1
        SketchLineAt tp                           ' select the cutting line
        Set sv = Nothing
        On Error Resume Next
        Set sv = swDraw.CreateSectionViewAt5(px * IN_M, py * IN_M, 0#, "B", opt, Nothing, 0#)
        If sv Is Nothing Then Err.Clear: Set sv = swDraw.CreateSectionViewAt4(px * IN_M, py * IN_M, 0#, "B", opt, Nothing)
        On Error GoTo 0
        If Not sv Is Nothing Then
            If FrameOK(sv, vR, 1, 0, vF, 0, 1) Then
                Check True, "SECTION B-B: looking toward the back, bosses up (arrows point up)"
                Set MakeSection = sv: Exit Function
            End If
            ' looking the wrong way: try flipping the cut direction in place
            On Error Resume Next
            Set sec = sv.GetSection
            sec.SetReversedCutDirection Not sec.GetReversedCutDirection
            swModel.ForceRebuild3 False
            On Error GoTo 0
            If FrameOK(sv, vR, 1, 0, vF, 0, 1) Then
                Check True, "SECTION B-B: looking toward the back, bosses up (direction flipped)"
                Set MakeSection = sv: Exit Function
            End If
            If opt = 1 Then
                Check False, "SECTION B-B looks the wrong way - click the section line and use Flip Direction"
                Set MakeSection = sv: Exit Function
            End If
            DeleteView sv                         ' start again with the other direction
            Set seg = DrawViewLine(parent, a, b, cand)
        End If
    Next
    Set MakeSection = Nothing
End Function

' draw a sketch line in a drawing view; a, b in sheet inches; cand picks the coordinate convention
Function DrawViewLine(v As Object, a As Variant, b As Variant, ByVal cand As Integer) As Object
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, o As Variant, sc As Double, pos As Variant
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    sc = 1#
    On Error Resume Next
    sc = v.ScaleDecimal
    If sc <= 0 Then sc = 1#
    Select Case cand
        Case 1                                    ' sheet coordinates
            x1 = a(0) * IN_M: y1 = a(1) * IN_M: x2 = b(0) * IN_M: y2 = b(1) * IN_M
        Case 2                                    ' relative to the model origin, model scale
            o = XfPt(v.ModelToViewTransform, 0#, 0#, 0#)
            x1 = (a(0) - o(0)) * IN_M / sc: y1 = (a(1) - o(1)) * IN_M / sc
            x2 = (b(0) - o(0)) * IN_M / sc: y2 = (b(1) - o(1)) * IN_M / sc
        Case 3                                    ' relative to the view centre
            pos = v.Position
            x1 = (a(0) * IN_M - pos(0)) / sc: y1 = (a(1) * IN_M - pos(1)) / sc
            x2 = (b(0) * IN_M - pos(0)) / sc: y2 = (b(1) * IN_M - pos(1)) / sc
    End Select
    swModel.SketchManager.AddToDB = True
    Set DrawViewLine = swModel.SketchManager.CreateLine(x1, y1, 0#, x2, y2, 0#)
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Function

Function SketchLineAt(p As Variant) As Boolean
    swModel.ClearSelection2 True
    On Error Resume Next
    SketchLineAt = swModel.Extension.SelectByID2("", "SKETCHSEGMENT", p(0) * IN_M, p(1) * IN_M, 0#, False, 0, Nothing, 0)
End Function

Sub DeleteSeg(seg As Object)
    On Error Resume Next
    If seg Is Nothing Then Exit Sub
    swModel.ClearSelection2 True
    seg.Select4 False, Nothing
    swModel.EditDelete
    swModel.ClearSelection2 True
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
'  Centre marks, centrelines, dimensions
' =====================================================================================
Sub CentreMarksAndLines()
    CollectEdges viewPar
    CentreMarkAt viewPar, Array(3.25 + 0.1326, -0.1326, 1.025), "Plan view: centre mark on the " & ChrW(216) & ".375 hole"
    CentreMarkAt viewPar, Array(0.625, 0#, 0.6), "Plan view: centre mark on the slot"
    CollectEdges viewSec
    CentreMarkAt viewSec, Array(1.625 + 0.145, 0#, 0.3), "Section: centre mark on the " & ChrW(216) & ".290 hole"
    CentreLines viewPar
    CentreLines viewSec
    CentreLines viewRight
End Sub

Sub CentreMarkAt(v As Object, p As Variant, ByVal label As String)
    Dim cm As Object
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    If Not SelPt(v, p, False) Then Check False, label & " - circle not found": Exit Sub
    Set cm = Nothing
    On Error Resume Next
    Set cm = swDraw.InsertCenterMark3(0, False, True)
    On Error GoTo 0
    swModel.ClearSelection2 True
    Check Not cm Is Nothing, label
End Sub

Sub CentreLines(v As Object)
    On Error Resume Next
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 v.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    swDraw.InsertCenterLine2
    swModel.ClearSelection2 True
End Sub

Sub DimPlan()
    Dim hx As Double
    CollectEdges viewPar
    ViewExtents viewPar
    hx = SheetPt(viewPar, 3.25, 0, 0)(0)
    MakeDimP viewPar, "R", Array(0.9482, 0.5518, 1.075), Empty, xMin - 0.27, yMax + 0.32, 0.25, "Plan     R.25 TYP", "", 2, " TYP"
    MakeDimP viewPar, "R", Array(2.2134, 0.4634, 1.075), Empty, xMin + 3.33, yMax + 0.32, 0.125, "Plan     4 X R.125", "4 X ", 3
    MakeDimP viewPar, "R", Array(3.9571, 0.7071, 0.6), Empty, xMax + 0.54, yMax + 0.32, 1#, "Plan     R1.00", "", 2
    MakeDimP viewPar, "V", Array(1.5, 1#, 0.6), Array(0.625, 0#, 0.6), xMin - 0.65, yMax - 0.5, 1#, "Plan     1.00", "", 2
    LeaderNoteP viewPar, Array(0.0625, -0.9375, 0.6), ".125 X 45<MOD-DEG>", xMin - 1.16, yMin - 0.1, "Plan     .125 X 45" & ChrW(176) & " chamfer note"
    MakeDimP viewPar, "H", Array(0#, 0.6, 0.6), Array(0.625, 0#, 0.6), xMin - 0.39, yMin - 0.4, 0.375, "Plan     .375", "", 3
    MakeDimP viewPar, "H", Array(0#, 0.6, 0.6), Array(0.875, 0.1, 1.075), xMin - 0.39, yMin - 0.68, 0.875, "Plan     .875", "", 3
    MakeDimP viewPar, "H", Array(0#, 0.6, 0.6), Array(3.25 + 0.1326, -0.1326, 1.025), (xMin + hx) / 2, yMin - 0.94, 3.25, "Plan     3.25", "", 2
    MakeDimP viewPar, "D", Array(3.25 + 0.1326, -0.1326, 1.025), Empty, xMax + 0.21, yMin - 0.26, 0.375, "Plan     " & ChrW(216) & ".375", "", 3
End Sub

Sub DimSection()
    CollectEdges viewSec
    ViewExtents viewSec
    MakeDimP viewSec, "H", Array(0.875, 0#, 0.84), Array(2.375, 0#, 0.84), SX(viewSec, 1.625), yMax + 0.71, 1.5, "Section  1.500", "", 3
    MakeDimP viewSec, "H", Array(0.875, 0#, 0.84), Array(1#, 0#, 0.89), xMin + 0.24, yMax + 0.51, 0.125, "Section  4 X .125", "4 X ", 3
    MakeDimP viewSec, "V", Array(0.9375, 0#, 1.075), Array(1.4, 0#, 0.7), xMin + 0.71, yMax + 0.26, 0.375, "Section  .375  (pocket depth)", "", 3
    MakeDimP viewSec, "H", Array(1#, 0#, 0.89), Array(2.25, 0#, 0.89), SX(viewSec, 1.625), yMax + 0.26, 1.25, "Section  1.25", "", 2
    MakeDimP viewSec, "H", Array(2.875, 0#, 0.8125), Array(3.625, 0#, 0.8125), SX(viewSec, 3.25), yMax + 0.36, 0.75, "Section  " & ChrW(216) & ".75", "<MOD-DIAM>", 2
    MakeDimP viewSec, "V", Array(2.6, 0#, 0#), Array(0.9375, 0#, 1.075), xMin - 0.65, (yMin + yMax) / 2, 1.075, "Section  1.075", "", 3
    MakeDimP viewSec, "V", Array(1.77, 0#, 0.3), Array(2.6, 0#, 0#), xMin - 0.27, yMin + 0.16, 0.3, "Section  .30", "", 2
    MakeDimP viewSec, "H", Array(0#, 0#, 0.3), Array(1.77, 0#, 0.3), SX(viewSec, 0.8125), yMin - 0.23, 1.625, "Section  1.625", "", 3
    MakeDimP viewSec, "D", Array(1.77, 0#, 0.3), Empty, xMin + 2.74, yMin - 0.24, 0.29, "Section  " & ChrW(216) & ".290", "", 3
    MakeDimP viewSec, "V", Array(2.97, 0#, 1.025), Array(3.15, 0#, 0.275), xMax + 0.42, SY(viewSec, 0.65), 0.75, "Section  .75  (hole depth)", "", 2
    MakeDimP viewSec, "V", Array(2.6, 0#, 0#), Array(2.97, 0#, 1.025), xMax + 0.89, SY(viewSec, 0.5125), 1.025, "Section  1.025", "", 3
End Sub

Function SX(v As Object, ByVal r As Double) As Double
    SX = SheetPt(v, r, 0, 0)(0)
End Function

Function SY(v As Object, ByVal f As Double) As Double
    SY = SheetPt(v, 0, 0, f)(1)
End Function

Sub DimRight()
    CollectEdges viewRight
    ViewExtents viewRight
    MakeDimP viewRight, "V", Array(2#, 1#, 0.3), Array(2#, -1#, 0.3), xMin - 1.35, (yMin + yMax) / 2, 2#, "Right    2.000", "", 3
    MakeDimP viewRight, "V", Array(1.6, 0.625, 0.84), Array(1.6, -0.625, 0.84), xMin - 0.42, (yMin + yMax) / 2, 1.25, "Right    1.250", "", 3
    MakeDimP viewRight, "V", Array(2#, 1#, 0.3), Array(1.6, 0.625, 0.84), xMin - 0.42, yMax + 0.29, 0.375, "Right    (.375)", "", 3, "", True
    MakeDimP viewRight, "H", Array(2#, 0.8, 0.6), Array(2#, 0.8, 0#), xMax + 0.32, yMin - 0.33, 0.6, "Right    .60", "", 2
    LeaderNoteP viewRight, Array(2#, -0.9875, 0.5875), ".025 X 45<MOD-DEG> TYP", xMin - 1.3, yMin - 0.42, "Right    .025 X 45" & ChrW(176) & " TYP note"
End Sub

' kind: H horizontal, V vertical, D diameter, R radius.  pa / pb: points on the edges in part coordinates
Sub MakeDimP(v As Object, ByVal kind As String, pa As Variant, pb As Variant, ByVal tx As Double, ByVal ty As Double, _
             ByVal expect As Double, ByVal label As String, Optional ByVal prefix As String = "", Optional ByVal prec As Long = -1, _
             Optional ByVal suffix As String = "", Optional ByVal paren As Boolean = False)
    Dim dd As Object, got As Double, nSel As Long, need As Long, txt As String, okA As Boolean, okB As Boolean

    If kind = "H" Or kind = "V" Then need = 2 Else need = 1
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    okA = SelPt(v, pa, False)
    If need = 2 Then okB = SelPt(v, pb, True)
    nSel = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    If nSel <> need Then
        Check False, label & "  - MISSING (found " & nSel & " of " & need & " edge(s))"
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
    dd.ShowParenthesis = paren
    If prec >= 0 Then
        dd.SetPrecision3 prec, -1, -1, -1
        If Err.Number <> 0 Then Err.Clear: dd.SetPrecision2 prec, -1, -1, -1
    End If
    If Len(prefix) > 0 Then
        txt = ""
        txt = dd.GetText(1)                                     ' 1 = prefix (holds the O / R symbol)
        If Left(txt, Len(prefix)) <> prefix Then prefix = prefix & txt Else prefix = txt
        dd.SetText 1, prefix
        txt = dd.GetText(1)
    End If
    If Len(suffix) > 0 Then dd.SetText 2, suffix                ' 2 = suffix
    On Error GoTo 0

    If Abs(got - expect) > 0.003 Then
        Check False, label & "  - WRONG (measured " & Format(got, "0.000") & ")"
    ElseIf Len(prefix) > 0 And Trim(txt) <> Trim(prefix) Then
        Check False, label & "  - value OK, but the prefix did not stick; add it by hand"
    Else
        Check True, label
    End If
End Sub

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

Sub LeaderNoteP(v As Object, p As Variant, ByVal text As String, ByVal tx As Double, ByVal ty As Double, ByVal label As String)
    Dim nt As Object, hasLeader As Boolean
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    hasLeader = SelPt(v, p, False)
    Set nt = Nothing
    On Error Resume Next
    Set nt = swModel.InsertNote(text)
    If Not nt Is Nothing Then nt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
    On Error GoTo 0
    swModel.ClearSelection2 True
    If nt Is Nothing Then
        Check False, label & "  - MISSING"
    ElseIf Not hasLeader Then
        Check False, label & "  - placed, but the chamfer edge was not found, so it has no leader; drag one to the corner"
    Else
        Check True, label
    End If
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
                nt.SetText "SHEET 2 OF 4"
                If UCase(Trim(nt.GetText)) <> "SHEET 2 OF 4" Then nt.PropertyLinkedText = "SHEET 2 OF 4"
                sheetDone = (UCase(Trim(nt.GetText)) = "SHEET 2 OF 4")
            End If
            Set nt = nt.GetNext
        Loop
        If pass = 2 Then swDraw.EditSheet
        On Error GoTo 0
        If gotM And gotD And gotT And sheetDone Then Exit For
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
    AddSheetNote Format(Date, "mm-dd-yy"), dX + 0.4, dY - 0.18, "Date: " & Format(Date, "mm-dd-yy")
    AddSheetNote "IME 1140 ENGINEERING DRAWINGS I", tX + 0.44, tY, "Title line 1"
    AddSheetNote "IN-LAB ASSIGNMENT #7B", tX + 0.44, tY - 0.17, "Title line 2"
    Check sheetDone, "SHEET 2 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
End Sub

Sub FinishSheet()
    ' leave the drawing active, in normal sheet mode, with all four views shown and redrawn
    Dim vs As Variant, n As Long, i As Long, v As Variant
    On Error Resume Next
    ActivateDrawing
    swDraw.EditSheet
    swModel.ClearSelection2 True
    For Each v In Array(viewPar, viewSec, viewRight)
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
    lastName = Trim(InputBox("Last name for the file names  HW#7B<LastName>.SLDDRW  and  .PDF" & vbCrLf & _
                             "(leave empty to skip saving)", "Save HW#7B", "Perhach"))
    If Len(lastName) = 0 Then WriteReport SaveFolder() & "\": Exit Sub
    folder = Left(partPath, InStrRev(partPath, "\"))
    drw = folder & "HW#7B" & lastName & ".SLDDRW"
    ok = TrySave(drw)
    If Not ok Then
        folder = SaveFolder() & "\"
        drw = folder & "HW#7B" & lastName & ".SLDDRW"
        ok = TrySave(drw)
    End If
    If Not ok Then
        MsgBox "Could not save the drawing automatically. Use File > Save As and name it  HW#7B" & lastName & _
               ", then File > Save As > PDF with the same name.", vbExclamation, "HW#7B"
        Exit Sub
    End If
    pdf = folder & "HW#7B" & lastName & ".PDF"
    FinishSheet                                         ' views visible + drawing active before the PDF export
    okPdf = TrySavePdf(pdf)
    WriteReport folder
    MsgBox "Saved:" & vbCrLf & drw & vbCrLf & IIf(okPdf, pdf, "PDF NOT saved - use File > Save As > PDF"), _
           IIf(okPdf, vbInformation, vbExclamation), "HW#7B"
End Sub

Sub WriteReport(ByVal folder As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open folder & "HW7B_checks.txt" For Output As #f
    Print #f, "HW#7B drawing checks  (" & Now & ")"
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
