' =====================================================================================
'  HW#8B  -  IME 1140 Engineering Drawings II, In-Lab Assignment #8B  -  DRAWING macro
'            (inches, sheet 1:1, third angle)
'
'  HOW TO RUN
'   1. Open "IME 1140 Engineering Drawings II In-Lab Assignment Part 2.SLDPRT" in SolidWorks.
'   2. Tools > Macro > New...  save it as  HW8B_Drawing.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main  (Tools > Macro > Run, select "main", Run).
'   5. At the end enter your last name; it saves  HW#8B<LastName>.SLDDRW  and  .PDF  next to
'      the part, plus HW8B_checks.txt with every check.
'   If the final message lists "TO DO BY HAND" items (things SolidWorks does not let a macro
'   pick, such as outline curves), do them, then press Ctrl+S and File > Save As > PDF.
'   If DETAIL A could not be made, make it by hand as the message says and run main again:
'   the macro then finishes the open drawing instead of starting a new one.
'
'  WHAT IT BUILDS (matches the reference sheet)
'   - Main view (dia .925 body on the left, 1/2-13 thread on the right), hidden lines visible
'   - Left end view (tapped hole), DETAIL A at 3:1 round the thread relief, isometric view 1:1
'   - Main:   4.000, 2.000, 1.750, .875, .094, dia .925 (datum A), dia .775, dia .625, (dia .493),
'             2X .05 X 45 deg, 1/2-13 UNC 2A, circularity .003, circular runout .003 A
'   - Left:   5/16-18 UNC deep .750, centre mark
'   - Detail: .094, 30 deg, dia .400, 3 X R.047
'   - NOTES 1-4 top left, as on the reference
'   - Title block: 6061-T6 ALUMINUM, today's date, IME 1140 ENGINEERING DRAWINGS II /
'     IN-LAB ASSIGNMENT #8B, SHEET 2 OF 4, SCALE 1:1; tolerance block checked (+-.01 / +-.005)
'
'  BUILT TO AVOID EVERY HW#8A ERROR (playbook 4.8)
'   - OK / PROBLEM is judged on values read back from SolidWorks, never on an API return value
'   - no temporary dimensions are made; the detail circle is proved by what DETAIL A shows
'   - the sheet is set back to 1:1 after the first view (auto-scale changed it) and every view's
'     scale is read back; tolerance block read back; symbol codes from the playbook
'   - things a macro cannot pick are listed as TO DO BY HAND with exact steps
'
'  v1: first version
'
'  SAFETY
'   - Orientation read from the part: dia .925 body and dia .625 shank give the axis and its ends.
'   - Edges picked from exact positions; diameters from side-on circle edges of that length.
'   - The part file is never saved or changed by this macro.
' =====================================================================================
Option Explicit

Dim swApp As Object, swMath As Object, swPart As Object, swModel As Object, swDraw As Object
Dim partPath As String, curStep As String, report As String, problems As String, todoList As String
Dim nOK As Long, nBad As Long, nTodo As Long
Dim lastDim As Object                      ' the dimension TryDimSel made last
Dim oldInputDim As Boolean, savedInputDim As Boolean

Const IN_M = 0.0254
Const DRAW_NAME = "HW#8B"
Const FILE_TAG = "HW8B"
Const SHEET_DEN = 1#             ' sheet scale 1 : 1
Const DET_SCALE = 3#             ' DETAIL A 3 : 1
Const EMAX = 8000                ' edge table size (the modelled threads add many edges)
Const PLACE_TOL = 0.001          ' views placed to within .001 in

' ---- layout (sheet inches, measured from the reference: border 10.5 x 8 in = 841 x 647 px) ----
Const MAIN_X = 4.85              ' left end of the part, main view
Const MAIN_Y = 2.59              ' axis, main view (the left view is level with it)
Const LEFT_X = 1.69              ' axis of the left end view
Const DET_X = 2.15               ' DETAIL A: the shank's end face (x = 3.125)
Const DET_Y = 5.45               ' DETAIL A: axis
Const ISO_X = 8.81               ' isometric: middle of the part (x = 2.000 on the axis)
Const ISO_Y = 6.46
Const NOTE_X = 0.4               ' NOTES block: top-left of "NOTES:"
Const NOTE_Y = 8.19
Const NOTE_DY = 0.168            ' line spacing
Const NOTE_TAB = 0.42            ' note text starts this far right of its number

Const DM_HIDDEN_VISIBLE = 1
Const DM_HIDDEN_REMOVED = 2

' ---- axial positions (inches from the left end) ----
Const X_GRV1 = 1.75              ' groove (full round R.047, bottom dia .775)
Const X_GRV2 = 1.844             ' groove right wall (1.750 + .094)
Const X_COLLAR = 2#              ' collar's right face = start of the dia .625 shank
Const X_SHANK = 3.125            ' shank's end face (4.000 - .875)
Const X_NECK2 = 3.219            ' end of the dia .400 neck (3.125 + .094) = start of the 30 deg chamfer
Const X_DETC = 2.976             ' DETAIL A circle centre (main view), radius .34
Const DET_R = 0.34

Dim vR As Variant                ' along the axis, body -> thread end   (right in the main view)
Dim vU As Variant                ' a radial direction                     (up in the main view)
Dim vF As Variant                ' toward you in the main view
Dim pOrg As Variant              ' model point on the axis at the left end

Dim viewMain As Object, viewLeft As Object, viewDet As Object, viewIso As Object
Dim detailPending As Boolean, resumed As Boolean

Dim eN As Long
Dim eObj() As Object, eKind() As String
Dim eX1() As Double, eY1() As Double, eX2() As Double, eY2() As Double
Dim eCX() As Double, eCY() As Double, eR() As Double, eFull() As Boolean
Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double

Sub main()
    On Error GoTo Fail
    Dim tmpl As String, msg As String
    Set swApp = Application.SldWorks
    Set swMath = swApp.GetMathUtility
    report = "": problems = "": todoList = "": nOK = 0: nBad = 0: nTodo = 0
    detailPending = False: resumed = False
    Set viewMain = Nothing: Set viewLeft = Nothing: Set viewDet = Nothing: Set viewIso = Nothing   ' VBA keeps these between runs
    Set lastDim = Nothing
    NoDimDialog True

    curStep = "Step 1 - find the open part":           FindPart
    curStep = "Step 2 - read the part orientation":    FindOrientation
    curStep = "Step 3 - continue an open drawing?":    resumed = ResumeDrawing()
    If Not resumed Then
        curStep = "Step 3 - find your template":       tmpl = FindTemplate()
        curStep = "Step 4 - new drawing":              NewDrawing tmpl
        curStep = "Step 5 - main, left and iso views": MakeViews
    End If
    ' the main view is dimensioned BEFORE DETAIL A exists (a detail makes SolidWorks double-select
    ' main-view edges and refuse the dimensions)
    curStep = "Step 6 - main and left view dimensions"
    ClearCommand
    If HasDims(viewMain) Then
        Check True, "Main view dimensions already there - left as they are"
    ElseIf Not viewDet Is Nothing Then
        Check False, "Main view dimensions: this drawing already has DETAIL A, which blocks them. Close it without saving and run main on a fresh drawing"
    Else
        DimMain
        ClearCommand
        DimLeft
        ClearCommand
        GtolMain
    End If
    If Not resumed Then
        curStep = "Step 7 - DETAIL A"
        ClearCommand
        Set viewDet = MakeDetail()
        If viewDet Is Nothing Then detailPending = True
    End If
    If Not viewDet Is Nothing Then
        curStep = "Step 8 - DETAIL A dimensions"
        FinishDetail
        ClearCommand
        If HasDims(viewDet) Then Check True, "DETAIL A dimensions already there - left as they are" Else DimDetail
    End If
    curStep = "Step 9 - notes"
    ClearCommand
    If SheetHasText("FOR 5/16-18 TAP") Then
        Check True, "NOTES already there - left as they are"
    Else
        AddNotes
    End If
    curStep = "Step 10 - title block"
    ClearCommand
    If SheetHasText("IN-LAB ASSIGNMENT #8B") Then
        Check True, "Title block already filled in - left as it is"
    Else
        TitleBlock
    End If
    curStep = "Step 11 - centre marks and centrelines" ' last: the centreline command can swallow later clicks
    LeftCentreMark
    CentreLines viewMain
    If Not viewDet Is Nothing Then CentreLines viewDet
    ClearCommand
    curStep = "Step 12 - scales and alignment":       CheckScalesAndAlignment
    curStep = "Step 13 - finish":                     FinishSheet
    NoDimDialog False

    If detailPending Then
        msg = vbCrLf & "DETAIL A (the macro could not make it):" & vbCrLf & _
              "  1. Drawing tab > Detail View. Draw a circle round the neck at the shank end (dia .34 in" & vbCrLf & _
              "     round the point 2.976 from the left end, on the axis)." & vbCrLf & _
              "  2. Click above-left of the main view. Label A, Scale > custom 3:1. Green check." & vbCrLf & _
              "  3. Run main again with this drawing open: it adds the DETAIL A dimensions." & vbCrLf
    End If
    MsgBox DRAW_NAME & IIf(resumed, " drawing finished.", " drawing built.") & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & "      to do by hand: " & nTodo & vbCrLf & _
           IIf(nBad = 0, "", vbCrLf & "PROBLEMS:" & vbCrLf & Left(problems, 450)) & _
           IIf(nTodo = 0, "", vbCrLf & "TO DO BY HAND:" & vbCrLf & Left(todoList, 600)) & msg & vbCrLf & _
           "Full list: " & FILE_TAG & "_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), DRAW_NAME

    curStep = "Step 14 - save drawing and PDF":       SaveAll
    Exit Sub
Fail:
    NoDimDialog False
    MsgBox "The macro stopped at: " & curStep & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Checks passed so far: " & nOK & vbCrLf & _
           IIf(nBad = 0, "", "Problems so far:" & vbCrLf & Left(problems, 500)), vbCritical, DRAW_NAME & " macro"
End Sub

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

Function PartMatches(ByVal t As String) As Boolean
    t = UCase(Replace(Replace(t, "_", " "), "#", ""))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    PartMatches = (InStr(t, "DRAWINGS II") > 0 And InStr(t, "PART 2") > 0)
End Function

Sub FindPart()
    ' Use the open part named "...Drawings II In-Lab Assignment Part 2". Otherwise the active part,
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
        "Could not tell which open part is In-Lab Assignment Part 2. Close the other parts, click into its window, and run again."
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then Err.Raise vbObjectError + 2, , "The part has never been saved. Save it first."
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

Sub FindOrientation()
    ' dia .925 body (r .4625): its axis is the part axis.  dia .625 shank (r .3125) tells which end is right.
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, k As Long
    Dim s As Object, cp As Variant, bx As Variant, r As Double
    Dim bAx As Variant, bC(2) As Double, bN As Long, a0 As Variant, pC(2) As Double, pN As Long
    Dim bb As Variant, c As Variant, ix As Integer, iy As Integer, iz As Integer, t As Double
    Dim rMin As Double, rMax As Double, cand As Variant, wMax As Double

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
    If bN = 0 Then Err.Raise vbObjectError + 4, , "This part has no dia .925 body, so it is not In-Lab Assignment Part 2. Click into its window and run again."
    If pN = 0 Then Err.Raise vbObjectError + 4, , "Could not find the dia .625 shank in the part."
    For k = 0 To 2: bC(k) = bC(k) / bN: pC(k) = pC(k) / pN: Next

    vR = bAx                                                      ' shank is to the right of the body
    If (pC(0) - bC(0)) * vR(0) + (pC(1) - bC(1)) * vR(1) + (pC(2) - bC(2)) * vR(2) < 0 Then vR = Scl(vR, -1)
    ' "up": a model axis square to the part axis (prefers +Y, then +Z, then +X) so standard views can be used
    vU = Empty
    For Each cand In Array(Arr3(0, 1, 0), Arr3(0, 0, 1), Arr3(1, 0, 0))
        If IsEmpty(vU) Then
            If Abs(Dot(cand, vR)) < 0.01 Then vU = cand
        End If
    Next
    If IsEmpty(vU) Then vU = Unit(Cross(vR, Arr3(0, 0, 1)))
    vF = Cross(vR, vU)

    bb = bodies(LBound(bodies)).GetBodyBox
    rMin = 1E+30: rMax = -1E+30: wMax = 0
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

' =====================================================================================
'  Template, new drawing  (from HW#7D)
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
'  Views, scales and DETAIL A
' =====================================================================================
Sub MakeViews()
    Dim dv As Variant, upI As Variant, rtI As Variant, ok As Boolean

    Set viewMain = OrientedView("HW8B_Main", vR, vU, vF, MAIN_X + 2#, MAIN_Y, StdFrontName(), False)
    If viewMain Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented main view."
    Check True, "Main view: body on the left, thread on the right"
    ' HW#8A: "auto-scale new drawing views" changed the SHEET to 1:2 when the first view went in
    EnsureSheetScale
    ok = SetViewScale(viewMain, 1#, 1#)
    Check ok, "Main view scale 1:1 (reads " & ScaleTxt(viewMain) & ")"
    SetMode viewMain, DM_HIDDEN_VISIBLE, "Main"
    PlaceView viewMain, 0#, 0#, 0#, MAIN_X, MAIN_Y, True, True, "Main view placed (left end at " & Format(MAIN_X, "0.00") & ", axis at " & Format(MAIN_Y, "0.00") & ")"

    ' left end view, projected to the left (third angle), level with the main view
    Set viewLeft = Projected(viewMain, LEFT_X, MAIN_Y)
    If viewLeft Is Nothing Then
        Check False, "Left end view - SolidWorks would not project it; add it with Projected View, to the left of the main view"
    Else
        ok = FrameOK(viewLeft, vF, 1, 0, vU, 0, 1)
        Check ok, "Left end view looks at the left end (tapped hole)" & IIf(ok, "", " - it shows the wrong end; delete it and project it to the LEFT of the main view")
        ok = SetViewScale(viewLeft, 1#, 1#)
        Check ok, "Left end view scale 1:1 (reads " & ScaleTxt(viewLeft) & ")"
        SetMode viewLeft, DM_HIDDEN_REMOVED, "Left end"
        PlaceView viewLeft, 0#, 0#, 0#, LEFT_X, MAIN_Y, True, False, "Left end view placed (axis at x = " & Format(LEFT_X, "0.00") & ")"
    End If

    dv = Unit(Add3(Add3(vR, vU), vF))
    upI = Unit(Sub3(vU, Scl(dv, Dot(vU, dv))))
    rtI = Cross(upI, dv)
    Set viewIso = OrientedView("HW8B_Iso", rtI, upI, dv, ISO_X, ISO_Y, StdIsoName(), True)
    If viewIso Is Nothing Then
        Check False, "Isometric view - could not create it; add it by hand (View Palette > Isometric)"
    Else
        Check True, "Isometric view"
        ok = SetViewScale(viewIso, 1#, 1#)
        Check ok, "Isometric view scale 1:1 (reads " & ScaleTxt(viewIso) & ")"
        SetMode viewIso, DM_HIDDEN_REMOVED, "Isometric"
        PlaceView viewIso, X_COLLAR, 0#, 0#, ISO_X, ISO_Y, True, True, "Isometric view placed (middle of the part)"
    End If
End Sub

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
            swModel.ForceRebuild3 False
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
        swModel.ForceRebuild3 False
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

Function Projected(parent As Object, ByVal px As Double, ByVal py As Double) As Object
    Dim before As String
    before = ViewNames()
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 parent.GetName2, "DRAWINGVIEW", 0#, 0#, 0#, False, 0, Nothing, 0
    On Error Resume Next
    Set Projected = swDraw.CreateUnfoldedViewAt3(px * IN_M, py * IN_M, 0#, False)
    On Error GoTo 0
    If Projected Is Nothing Then Set Projected = NewViewSince(before)
    swModel.ClearSelection2 True
End Function

' DETAIL A: sketch a circle round the neck in the main view, then make the 3:1 detail from it.
' Each way of placing the circle is tried; a detail is kept only if it really shows the dia .400 neck.
Function MakeDetail() As Object
    Dim c As Variant, rad As Double, cand As Variant, seg As Object, dv As Object, before As String, sc As Double
    Dim combo As Integer, diag As String, used As String, st As Integer, sh As Integer, api As Integer

    c = SheetPt(viewMain, X_DETC, 0#, 0#)
    sc = ViewSc(viewMain)
    rad = DET_R * sc
    For Each cand In Array(0, 3, 2, 1)
        Set seg = DrawViewCircle(viewMain, c, rad, CInt(cand))
        If seg Is Nothing Then
            diag = diag & "m" & cand & ": no circle. "
        Else
            Set dv = Nothing
            For combo = -1 To 29
                If combo >= 0 Then
                    If SelectSeg(seg, c(0) + rad, c(1)) = 0 Then diag = diag & "m" & cand & ": not selectable. ": Exit For
                End If
                before = ViewNames()
                On Error Resume Next
                If combo = -1 Then                        ' as SolidWorks' own example: circle still selected
                    Set dv = swDraw.CreateDetailViewAt4((DET_X - 1#) * IN_M, DET_Y * IN_M, 0#, 0, DET_SCALE, 1#, "A", 1, True, False, False, 5)
                Else
                    api = combo \ 15: st = (combo \ 3) Mod 5: sh = combo Mod 3
                    If api = 0 Then
                        Set dv = swDraw.CreateDetailViewAt4((DET_X - 1#) * IN_M, DET_Y * IN_M, 0#, st, DET_SCALE, 1#, "A", sh, False, False, False, 5)
                    Else
                        Set dv = swDraw.CreateDetailViewAt3((DET_X - 1#) * IN_M, DET_Y * IN_M, 0#, st, DET_SCALE, 1#, "A", sh, False)
                    End If
                End If
                Err.Clear
                On Error GoTo 0
                If dv Is Nothing Then Set dv = NewViewSince(before)
                If Not dv Is Nothing Then Exit For
            Next
            If Not dv Is Nothing Then
                SetViewScale dv, DET_SCALE, 1#
                If DetailHasNeck(dv) Then used = "m" & cand: Exit For
                diag = diag & "m" & cand & ": wrong place. "
                DeleteView dv
                Set dv = Nothing
            ElseIf combo > 29 Then
                diag = diag & "m" & cand & ": refused. "
            End If
            DeleteSeg seg
        End If
    Next
    swModel.ClearSelection2 True
    If dv Is Nothing Then
        ToDo "DETAIL A: SolidWorks would not let the macro make it (" & Trim(diag) & ") - make it by hand (steps at the end), then run main again"
        Exit Function
    End If
    Check True, "DETAIL A created (" & used & ") and shows the " & ChrW(216) & ".400 neck"
    Set MakeDetail = dv
End Function

' scale, display and placement of DETAIL A (also for one made by hand)
Sub FinishDetail()
    Dim ok As Boolean
    ok = SetViewScale(viewDet, DET_SCALE, 1#)
    Check ok, "DETAIL A scale 3:1 (reads " & ScaleTxt(viewDet) & ")" & IIf(ok, "", " - set Scale > Use custom scale 3:1")
    PlaceView viewDet, X_SHANK, 0#, 0#, DET_X, DET_Y, True, True, "DETAIL A placed (shank end face, axis)"
    If Not DetailHasNeck(viewDet) Then
        Check False, "DETAIL A does not show the " & ChrW(216) & ".400 neck - its circle is in the wrong place. Delete DETAIL A and its circle, make it again round the neck"
    End If
End Sub

' a real DETAIL A shows the neck: a side-on circle edge .400 long near x = 3.17 .. 3.22
Function DetailHasNeck(dv As Object) As Boolean
    DetailHasNeck = (DiaLine(dv, X_SHANK, X_NECK2 + 0.01, 0.4, 0.01) >= 0)
End Function

' the side-on circle edge (a vertical line on the sheet) of diameter dia between axial positions x1 and x2
Function DiaLine(v As Object, ByVal x1 As Double, ByVal x2 As Double, ByVal dia As Double, ByVal tol As Double) As Long
    Dim i As Long, sc As Double, a As Double, b As Double, mx As Double
    DiaLine = -1
    CollectEdges v
    sc = ViewSc(v)
    a = SX(v, x1): b = SX(v, x2)
    If a > b Then mx = a: a = b: b = mx
    For i = 0 To eN - 1
        If eKind(i) = "L" Then
            If Abs(eX2(i) - eX1(i)) < 0.003 Then
                mx = (eX1(i) + eX2(i)) / 2
                If mx > a - 0.005 And mx < b + 0.005 Then
                    If Abs(Abs(eY2(i) - eY1(i)) / sc - dia) < tol Then DiaLine = i: Exit Function
                End If
            End If
        End If
    Next
End Function

' continue an open drawing of this part (e.g. after DETAIL A was made by hand)
Function ResumeDrawing() As Boolean
    Dim d As Object, v As Object, sc As Double, a As Variant
    Set d = swApp.ActiveDoc
    If d Is Nothing Then Exit Function
    If d.GetType <> 3 Then Exit Function
    Set swModel = d: Set swDraw = d
    Set viewMain = Nothing: Set viewLeft = Nothing: Set viewDet = Nothing: Set viewIso = Nothing
    On Error Resume Next
    Set v = swDraw.GetFirstView.GetNextView                           ' first "view" is the sheet itself
    Do While Not v Is Nothing
        sc = 0: sc = v.ScaleDecimal
        If Abs(sc - DET_SCALE) < 0.05 Then
            Set viewDet = v
        ElseIf ViewMatches(v, vR, vU) Then
            Set viewMain = v
        ElseIf ViewMatches(v, vF, vU) Then
            Set viewLeft = v
        Else
            a = ViewDir(v, vR)
            If a(0) > 0.5 And a(1) < -0.2 Then Set viewIso = v
        End If
        Set v = v.GetNextView
    Loop
    On Error GoTo 0
    If viewMain Is Nothing Then
        Set swModel = Nothing: Set swDraw = Nothing
        Exit Function
    End If
    On Error Resume Next
    swDraw.EditSheet                                  ' normal sheet mode (not sheet format, not a sketch)
    swModel.ClearSelection2 True
    swModel.ForceRebuild3 False
    swModel.GraphicsRedraw2
    On Error GoTo 0
    If viewDet Is Nothing Then detailPending = True
    ResumeDrawing = True
    Check True, "Continuing the open drawing: main view found" & IIf(viewDet Is Nothing, ", DETAIL A not made yet", ", DETAIL A found")
End Function

' every view at its scale and where it was measured; the left view level with the main view
Sub CheckScalesAndAlignment()
    Dim a As Variant, b As Variant, ok As Boolean
    EnsureSheetScale
    swModel.ForceRebuild3 False
    a = SheetPt(viewMain, 0, 0, 0)
    ok = (Abs(a(0) - MAIN_X) < 0.002 And Abs(a(1) - MAIN_Y) < 0.002)
    Check ok, "Alignment: main view left end at " & Format(a(0), "0.000") & ", " & Format(a(1), "0.000")
    If Not viewLeft Is Nothing Then
        b = SheetPt(viewLeft, 0, 0, 0)
        ok = (Abs(b(1) - a(1)) < 0.001)
        Check ok, "Alignment: left end view level with the main view (axis y " & Format(a(1), "0.000") & " / " & Format(b(1), "0.000") & ")"
    End If
    If Not viewDet Is Nothing Then
        Check (Abs(ViewSc(viewDet) - DET_SCALE) < 0.0005), "DETAIL A scale reads 3:1 at the end"
        Info "DETAIL A label: if it shows a SCALE 3:1 line the reference does not have, that is SolidWorks' standard label (leave it, or double-click the label to edit)"
    End If
    Check (Abs(ViewSc(viewMain) - 1#) < 0.0005), "Main view scale reads 1:1 at the end"
End Sub

' =====================================================================================
'  View helpers  (from HW#7D)
' =====================================================================================
Function StdIsoName() As String
    If Near(vR, 1, 0, 0) And Near(vU, 0, 1, 0) And Near(vF, 0, 0, 1) Then StdIsoName = "*Isometric"
End Function

' select a sketch segment: the object itself, else by clicking on it; returns how many are selected
Function SelectSeg(seg As Object, ByVal px As Double, ByVal py As Double) As Long
    swModel.ClearSelection2 True
    On Error Resume Next
    seg.Select4 False, Nothing
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) = 0 Then
        swModel.Extension.SelectByID2 "", "SKETCHSEGMENT", px * IN_M, py * IN_M, 0#, False, 0, Nothing, 0
    End If
    SelectSeg = swModel.SelectionManager.GetSelectedObjectCount2(-1)
End Function

' sketch a circle in a drawing view; c = centre (sheet inches), rad = radius (sheet inches).
' method 0: the view sketch's own sheet->sketch transform; 1: sheet; 2: from the model origin; 3: from the view centre.
' The new circle is left selected (as in SolidWorks' example).
Function DrawViewCircle(v As Object, c As Variant, ByVal rad As Double, ByVal cand As Integer) As Object
    Dim x As Double, y As Double, rr As Double, o As Variant, sc As Double, pos As Variant
    Dim sk As Object, xf As Object, a(2) As Double, va As Variant, p As Object, r As Variant
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    sc = 1#
    On Error Resume Next
    sc = v.ScaleDecimal
    If sc <= 0 Then sc = 1#
    Select Case cand
        Case 0
            Set sk = v.GetSketch
            Set xf = sk.ModelToSketchTransform
            If xf Is Nothing Then Exit Function
            a(0) = c(0) * IN_M: a(1) = c(1) * IN_M: a(2) = 0#: va = a
            Set p = swMath.CreatePoint((va))
            Set p = p.MultiplyTransform(xf)
            r = p.ArrayData
            If Not IsArray(r) Then Exit Function
            x = r(0): y = r(1): rr = rad * IN_M / sc
        Case 1
            x = c(0) * IN_M: y = c(1) * IN_M: rr = rad * IN_M
        Case 2
            o = XfPt(v.ModelToViewTransform, 0#, 0#, 0#)
            x = (c(0) - o(0)) * IN_M / sc: y = (c(1) - o(1)) * IN_M / sc: rr = rad * IN_M / sc
        Case 3
            pos = v.Position
            x = (c(0) * IN_M - pos(0)) / sc: y = (c(1) * IN_M - pos(1)) / sc: rr = rad * IN_M / sc
    End Select
    swModel.SketchManager.AddToDB = False
    Set DrawViewCircle = swModel.SketchManager.CreateCircle(x, y, 0#, x + rr, y, 0#)
    swModel.GraphicsRedraw2
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

' =====================================================================================
'  Dimensions, notes, GD&T  (text positions measured on the reference, sheet inches)
' =====================================================================================
Sub DimMain()
    Dim xL As Double, xR As Double, yA As Double, why As String, D As String
    D = ChrW(216)
    CollectEdges viewMain
    xL = SX(viewMain, 0): xR = SX(viewMain, 4#): yA = SY(viewMain, 0)

    ' lengths, stacked above the part
    DimToEnd viewMain, Array(0#, 0.3, 0#), (xL + xR) / 2, yA + 1.22, 4#, "Main     4.000"
    MakeDimP viewMain, "H", Array(0#, 0.3, 0#), Array(X_COLLAR, 0.4, 0#), (xL + SX(viewMain, X_COLLAR)) / 2, yA + 0.91, 2#, "Main     2.000", "", 3
    MakeDimP viewMain, "H", Array(0#, 0.3, 0#), Array(X_GRV1, 0.45, 0#), (xL + SX(viewMain, X_GRV1)) / 2, yA + 0.65, 1.75, "Main     1.750", "", 3
    DimToEnd viewMain, Array(X_SHANK, 0.28, 0#), xL + 2.78, yA + 0.87, 0.875, "Main     .875"
    MakeDimP viewMain, "H", Array(X_GRV1, 0.45, 0#), Array(X_GRV2, 0.45, 0#), xL + 1.29, yA - 0.77, 0.094, "Main     .094 (groove)", "", 3

    ' diameters, from the side-on circle edges of that size
    DiaByLen viewMain, 0#, X_COLLAR, 0.925, xL - 1.42, yA, D & ".925", 3, False
    DatumOnLastDim xL - 1.43, yA - 1.39
    why = TryDim(viewMain, "V", Array((X_GRV1 + X_GRV2) / 2, 0.3875, 0#), Array((X_GRV1 + X_GRV2) / 2, -0.3875, 0#), xL - 0.7, yA, 0.775, "<MOD-DIAM>", 3, "", False, True)
    If Len(why) = 0 Then
        Check True, "Main     " & D & ".775 (groove bottom)"
    Else
        ToDo "Main " & D & ".775: the groove bottom is an outline curve a macro cannot pick. Smart Dimension: click the groove's top bottom-curve, then its lower one, place it left of the part next to " & D & ".925 (" & D & " is added automatically for a diameter; if not, add it)"
    End If
    DiaByLen viewMain, X_COLLAR - 0.01, X_SHANK + 0.01, 0.625, xR + 0.67, yA + 0.47, D & ".625", 3, False
    DiaByLen viewMain, X_NECK2, 4.01, 0.493, xR + 0.32, yA + 0.8, "(" & D & ".493)", 3, True, _
             "the modelled thread leaves no whole " & D & ".493 circle edge. Smart Dimension the thread's top and bottom crest lines, place it right of the part above " & D & ".625, then in the PropertyManager add " & D & " and tick Parentheses (3 decimals)"

    ' notes with leaders
    LeaderNote2 viewMain, Array(0.05, 0.44, 0#), Array(0#, 0.39, 0#), "2X .05 X 45<MOD-DEG>", xL - 1.45, yA + 0.81, "Main     2X .05 X 45" & ChrW(176)
    LeaderNote2 viewMain, Array(4#, -0.15, 0#), Array(4#, 0.15, 0#), "1/2-13 UNC 2A", xR + 0.57, yA - 0.6, "Main     1/2-13 UNC 2A"
End Sub

' horizontal dimension from pa to the right end (x = 4.000); the right end face is tried at several heights
' (the modelled thread can break its edge into pieces)
Sub DimToEnd(v As Object, pa As Variant, ByVal tx As Double, ByVal ty As Double, ByVal expect As Double, ByVal label As String)
    Dim pb As Variant, why As String
    For Each pb In Array(Array(4#, 0.08, 0#), Array(4#, -0.08, 0#), Array(4#, 0.15, 0#), Array(4#, -0.15, 0#), Array(4#, 0.03, 0#))
        why = TryDim(v, "H", pa, pb, tx, ty, expect, "", 3, "", False, True)
        If Len(why) = 0 Then Exit For
    Next
    Check (Len(why) = 0), label & IIf(Len(why) = 0, "", "  - " & why)
End Sub

' diameter from a side-on circle edge of exactly that size between axial positions x1 and x2
Sub DiaByLen(v As Object, ByVal x1 As Double, ByVal x2 As Double, ByVal dia As Double, ByVal tx As Double, ByVal ty As Double, _
             ByVal label As String, ByVal prec As Long, ByVal paren As Boolean, Optional ByVal handHint As String = "")
    Dim i As Long, why As String, nm As String
    Set lastDim = Nothing                              ' so datum A can never land on an older dimension
    nm = IIf(v Is viewDet, "Detail A ", "Main     ") & label
    i = DiaLine(v, x1, x2, dia, 0.006)
    If i < 0 Then
        If Len(handHint) > 0 Then
            ToDo nm & ": " & handHint
        Else
            Check False, nm & "  - MISSING (no " & Format(dia, ".000") & " circle edge found)"
        End If
        Exit Sub
    End If
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    SelEdge v, i, False
    why = TryDimSel("VL", 1, tx, ty, dia, "<MOD-DIAM>", prec, "", paren, True)
    If Len(why) > 0 Then
        swModel.ClearSelection2 True
        SelEdge v, i, False
        why = TryDimSel("S1", 1, tx, ty, dia, "<MOD-DIAM>", prec, "", paren, True)
    End If
    If Len(why) > 0 Then Set lastDim = Nothing
    Check (Len(why) = 0), nm & IIf(Len(why) = 0, "", "  - " & why)
End Sub

' datum feature A on the dimension made last (the dia .925)
Sub DatumOnLastDim(ByVal tx As Double, ByVal ty As Double)
    Dim dt As Object, n0 As Long
    If lastDim Is Nothing Then Check False, "Main     datum A - MISSING (the " & ChrW(216) & ".925 dimension was not made)": Exit Sub
    swModel.ClearSelection2 True
    On Error Resume Next
    lastDim.GetAnnotation.Select3 False, Nothing
    n0 = swModel.SelectionManager.GetSelectedObjectCount2(-1)
    Set dt = Nothing
    If n0 = 1 Then Set dt = swModel.InsertDatumTag2
    If Not dt Is Nothing Then
        dt.SetLabel "A"
        dt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
    End If
    On Error GoTo 0
    swModel.ClearSelection2 True
    If dt Is Nothing Then
        Check False, "Main     datum A - MISSING; select the " & ChrW(216) & ".925 dimension, Annotations > Datum Feature, label A"
    Else
        Check (UCase(Trim(DatumLabel(dt))) = "A"), "Main     datum A on the " & ChrW(216) & ".925 dimension (label reads " & DatumLabel(dt) & ")"
    End If
End Sub

Function DatumLabel(dt As Object) As String
    On Error Resume Next
    DatumLabel = dt.GetLabel
End Function

' left end view: the tapped-hole callout
Sub DimLeft()
    Dim i As Long, best As Long, bestR As Double, c As Variant, xL As Double, yA As Double
    If viewLeft Is Nothing Then Check False, "Left     5/16-18 UNC callout - MISSING (no left end view)": Exit Sub
    CollectEdges viewLeft
    c = SheetPt(viewLeft, 0, 0, 0)
    best = -1: bestR = 1E+30
    For i = 0 To eN - 1
        If eKind(i) = "C" Then
            If Sqr((eCX(i) - c(0)) ^ 2 + (eCY(i) - c(1)) ^ 2) < 0.01 And eR(i) < bestR Then bestR = eR(i): best = i
        End If
    Next
    xL = SX(viewMain, 0): yA = SY(viewMain, 0)
    LeaderNoteI viewLeft, best, "5/16-18 UNC <HOLE-DEPTH> .750", xL - 2.24, yA + 1.09, "Left     5/16-18 UNC " & ChrW(&H21A7) & " .750 on the tapped hole" & _
                IIf(best >= 0, " (hole edge " & ChrW(216) & Format(2 * bestR / ViewSc(viewLeft), ".000") & ")", "")
End Sub

' centre mark (with extended lines) on the left end view's outer circle
Sub LeftCentreMark()
    Dim i As Long, best As Long, bestR As Double, c As Variant, cm As Object
    If viewLeft Is Nothing Then Exit Sub
    CollectEdges viewLeft
    c = SheetPt(viewLeft, 0, 0, 0)
    best = -1: bestR = 0
    For i = 0 To eN - 1
        If eKind(i) = "C" Then
            If Sqr((eCX(i) - c(0)) ^ 2 + (eCY(i) - c(1)) ^ 2) < 0.01 And eR(i) > bestR Then bestR = eR(i): best = i
        End If
    Next
    If best < 0 Then Check False, "Left     centre mark - outer circle not found": Exit Sub
    swModel.ClearSelection2 True
    swDraw.ActivateView viewLeft.GetName2
    SelEdge viewLeft, best, False
    Set cm = Nothing
    On Error Resume Next
    Set cm = swDraw.InsertCenterMark3(0, False, True)
    If Not cm Is Nothing Then cm.ShowLines = True
    On Error GoTo 0
    swModel.ClearSelection2 True
    Check Not cm Is Nothing, "Left     centre mark on the " & ChrW(216) & Format(2 * bestR / ViewSc(viewLeft), ".000") & " outer circle"
End Sub

' circularity .003 on the dia .925 body; circular runout .003 A on the dia .625 shank
Sub GtolMain()
    Dim xL As Double, yA As Double
    xL = SX(viewMain, 0): yA = SY(viewMain, 0)
    MakeGtol viewMain, Array(0.385, -0.4625, 0#), "<IGTOL-CIRC>", ".003", "", xL - 0.88, yA - 0.855, "Main     circularity .003 frame"
    MakeGtol viewMain, Array(2.43, -0.3125, 0#), "<IGTOL-SRUN>", ".003", "A", xL + 3.2, yA - 0.79, "Main     circular runout .003 A frame"
End Sub

' a feature control frame; its leader goes to the outline at p if SolidWorks lets the macro pick it.
' Judged on what reads back (symbol in your gtol.sym, the values), never on the call's return value.
Sub MakeGtol(v As Object, p As Variant, ByVal sym As String, ByVal tol1 As String, ByVal datum As String, _
             ByVal tx As Double, ByVal ty As Double, ByVal label As String)
    Dim g As Object, s As Variant, hasLeader As Boolean, n0 As Long, rb As String
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    CollectEdges v
    hasLeader = SelPt(v, p, False)                     ' table edge, else EDGE / SILHOUETTE pick at that point
    Set g = Nothing
    On Error Resume Next
    Set g = swModel.InsertGtol
    If Not g Is Nothing Then
        Err.Clear
        g.SetFrameSymbols2 1, sym, False, "", False, "", "", "", ""
        If Err.Number <> 0 Then Err.Clear: g.SetFrameSymbols3 1, sym, False, "", False, "", "", "", ""
        Err.Clear
        g.SetFrameValues2 1, tol1, "", datum, "", ""
        If Err.Number <> 0 Then Err.Clear: g.SetFrameValues 1, tol1, "", datum, "", ""
        g.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
        rb = GtolReadBack(g)
    End If
    On Error GoTo 0
    swModel.ClearSelection2 True
    If g Is Nothing Then
        Check False, label & " - MISSING; add it with Annotations > Geometric Tolerance"
        Exit Sub
    End If
    If Len(rb) = 0 Then
        Info label & " placed; SolidWorks does not report its values - look that it shows " & tol1 & IIf(Len(datum) > 0, " | " & datum, "")
    ElseIf InStr(rb, tol1) > 0 And InStr(rb, datum) > 0 Then
        Check True, label & " (reads " & rb & ")"
    Else
        Check False, label & " - values read back as " & rb & "; double-click it and set " & tol1 & IIf(Len(datum) > 0, ", datum " & datum, "")
    End If
    Info "  symbol " & sym & IIf(SymInLibrary(sym), " found in your gtol.sym", " not checked (gtol.sym not found or symbol not listed) - look that the symbol shows")
    If Not hasLeader Then ToDo label & ": its leader must go to the part outline, which a macro cannot pick. Click the frame, drag the small leader handle (or right-click > Add leader) to the bottom outline as on the reference"
End Sub

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

Sub DimDetail()
    Dim dX As Double, dY As Double, why As String, D As String
    D = ChrW(216)
    CollectEdges viewDet
    dX = SX(viewDet, X_SHANK): dY = SY(viewDet, 0)
    MakeDimP viewDet, "H", Array(X_SHANK, 0.28, 0#), Array(X_NECK2, 0.15, 0#), dX - 0.52, dY + 1.106, 0.094, "Detail A .094 (neck)", "", 3
    why = TryDim(viewDet, "A", Array(X_SHANK, 0.28, 0#), Array(X_NECK2 + 0.0134, 0.2232, 0#), dX + 0.42, dY + 1.48, 30#, "", 0, "", False, True)
    If Len(why) = 0 Then
        Check True, "Detail A 30" & ChrW(176)
    Else
        ToDo "Detail A 30" & ChrW(176) & ": the chamfer is an outline line a macro cannot pick. Smart Dimension: click the shank's end face line, then the slanted chamfer line at the thread start; place it above (30" & ChrW(176) & ")"
    End If
    DiaByLen viewDet, X_SHANK, X_NECK2 + 0.01, 0.4, dX + 1.95, dY, D & ".400", 3, False
    why = TryDim(viewDet, "R", Array(X_SHANK + 0.0138, -0.2138, 0#), Empty, dX + 1.14, dY - 1.12, 0.047, "3 X ", 3, "", False, True)
    If Len(why) = 0 Then
        Check True, "Detail A 3 X R.047"
    Else
        ToDo "Detail A 3 X R.047: the fillet is an outline curve a macro cannot pick. Smart Dimension on the lower fillet curve at the neck, place it lower right, then type ""3 X "" in front of R.047 (3 decimals)"
    End If
End Sub

' NOTES block, top left: each number and each text is its own note so the columns line up exactly
Sub AddNotes()
    Dim lines As Variant, i As Long, y As Double, ok As Boolean
    lines = Array("USE THE IN-LAB ASSIGNMENT #8B PART MODEL ON CANVAS TO RE-CREATE THE ENGINEERING DRAWING", _
                  "MAKE SURE YOUR DRAWING'S TOLERANCES MATCH THE DRAWING TOLERANCES SHOWN BELOW.", _
                  "CHAMFER EDGES .01 X 45<MOD-DEG> UNLESS SPECIFIED", _
                  "DRILL 1.00"" <HOLE-DEPTH> FOR 5/16-18 TAP")
    ok = PlaceNote("NOTES:", NOTE_X, NOTE_Y)
    For i = 0 To 3
        y = NOTE_Y - (i + 1) * NOTE_DY
        ok = PlaceNote((i + 1) & ".", NOTE_X, y) And ok
        ok = PlaceNote(CStr(lines(i)), NOTE_X + NOTE_TAB, y) And ok
    Next
    Check ok, "NOTES 1-4 top left (each read back)"
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

' does this view already carry dimensions (from an earlier run)?
Function HasDims(v As Object) As Boolean
    Dim n As Long
    On Error Resume Next
    n = 0
    n = v.GetDisplayDimensionCount
    HasDims = (n > 0)
End Function

' is this text already on the sheet (filled in by an earlier run)?
Function SheetHasText(ByVal t As String) As Boolean
    Dim nt As Object
    On Error Resume Next
    Set nt = swDraw.GetFirstView.GetFirstNote
    Do While Not nt Is Nothing
        If InStr(UCase(nt.GetText), UCase(t)) > 0 Then SheetHasText = True: Exit Function
        Set nt = nt.GetNext
    Loop
End Function

Sub ClearCommand()
    On Error Resume Next
    swModel.ClearSelection2 True
    swDraw.ActivateSheet swDraw.GetCurrentSheet.GetName
    swModel.ClearSelection2 True
    swModel.GraphicsRedraw2
End Sub

' =====================================================================================
'  Title block
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
                nt.SetText "SHEET 2 OF 4"
                If UCase(Trim(nt.GetText)) <> "SHEET 2 OF 4" Then nt.PropertyLinkedText = "SHEET 2 OF 4"
                sheetDone = (UCase(Trim(nt.GetText)) = "SHEET 2 OF 4")
            End If
            If Left(t, 6) = "SCALE:" And Not scaleDone Then              ' sheet is 1:1
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
    AddSheetNote "IN-LAB ASSIGNMENT #8B", tX + 0.44, tY - 0.17, "Title line 2"
    Check scaleDone, "SCALE: 1:1" & IIf(scaleDone, "", " - edit the scale note by hand")
    Check sheetDone, "SHEET 2 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
    Check (twoOK And threeOK), "Tolerance block reads TWO PLACE " & ChrW(177) & ".01, THREE PLACE " & ChrW(177) & ".005 (as on the reference)" & _
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

' =====================================================================================
'  Edge table: every visible edge of a view, in sheet inches  (from HW#7D)
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
'  Finish and save  (from HW#7D)
' =====================================================================================
Sub FinishSheet()
    ' leave the drawing active, in normal sheet mode, with every view shown and redrawn
    Dim vs As Variant, n As Long, i As Long, v As Variant
    On Error Resume Next
    ActivateDrawing
    swDraw.EditSheet
    swModel.ClearSelection2 True
    For Each v In Array(viewMain, viewLeft, viewDet, viewIso)
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
    If n >= 0 Then Check (n >= 4 Or (n >= 3 And viewDet Is Nothing)), n & " views on the sheet (expected 4)"
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
'  Geometry helpers  (from HW#7D)
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
