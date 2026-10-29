' =====================================================================================
'  HW#8C  -  IME 1140 Engineering Drawings II, In-Lab Assignment #8C  -  DRAWING macro
'            (inches, third angle; views 1:2, DETAIL B 1:1, isometric 1:4)
'
'  HOW TO RUN
'   1. Restart SolidWorks if it has been running a long time. Open
'      "IME 1140 Engineering Drawings II In-Lab Assignment Part 3.SLDPRT".
'   2. Tools > Macro > New...  save it as  HW8C_Drawing.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main.  The current step shows in SolidWorks' status bar.
'   5. At the end enter your last name; it saves  HW#8C<LastName>.SLDDRW  and  .PDF  next to
'      the part, plus HW8C_checks.txt with every check.
'
'  WHAT IT BUILDS (matches the reference sheet)
'   - Main view: square flange on the right (seen edge-on), elbow sweeping up-left to the round
'     flange at 45 deg; hidden lines visible
'   - VIEW A-A: auxiliary view looking square onto the round flange (bolt pattern)
'   - Right view: square flange face-on.  DETAIL B 1:1 round the lower-left hole.  Isometric 1:4.
'   - Main:   4.00+-.02, .375, dia 3.755/3.745, .375, 45 deg BASIC, (R6.000), datum A,
'             angularity .005 A, .050 X 45 deg TYP, 2 X R .050
'   - VIEW A-A: 5 X 72 deg BASIC, dia 3.150 BASIC (bolt circle), 5 X dia .257 with position
'             dia .01 | A | B, dia 2.250+-.005 with datum B
'   - Right:  3.00+-.02, 2.250+-.004, 1.125, 3.250+-.004, 1.625, centre marks
'   - DETAIL B: 4 X dia .375 THRU
'   - NOTES 1-2; title block CAST IRON, today's date (mm-dd-yy), #8C, SHEET 3 OF 4, SCALE: 1:4
'     (as on the reference); tolerance block THREE PLACE +-.008 (note 2)
'
'  BUILT FROM EVERY LESSON SO FAR (playbook 4.8)
'   - orientation measured from the part (both flanges' hole patterns, the bend's torus);
'     the 45 deg and R6.000 are measured, not assumed
'   - real-edge dimensions first in each view, then exact sketch items (bolt circle, 72 deg
'     lines, R6.000 arc) placed by a method proved once per view; values read back
'   - lean: one edge read per view, partial rebuilds, no per-step redraws, batch mode
'   - OK / PROBLEM judged on read-back values; GD&T symbols looked up in your gtol.sym
'
'  v3: compile fix (a variable named eNum reads as the reserved word Enum)
'  v2: fixed VIEW A-A stopping the macro (a variable named sp hid the function SP); the real
'      error is now shown if the macro stops; each big step (VIEW A-A, dimensions, DETAIL B ...)
'      reports its own error and the rest of the drawing is still built
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
Dim skCandMain As Integer
Dim eKey As String

Const IN_M = 0.0254
Const DRAW_NAME = "HW#8C"
Const FILE_TAG = "HW8C"
Const SHEET_DEN = 2#             ' sheet scale 1 : 2 (the views; the title block says 1:4 as on the reference)
Const ISO_DEN = 4#               ' isometric 1 : 4
Const EMAX = 8000
Const PLACE_TOL = 0.001
Const DEG = 0.0174532925199433

' ---- layout (sheet inches, measured from the reference: border 10.5 x 8 in = 841 x 648 px) ----
Const MAIN_X = 6.015             ' square flange's outer face, centre (main view)
Const MAIN_Y = 3.067
Const RIGHT_X = 9.557            ' bore centre, right view (level with the main view)
Const AA_X = 2.285               ' round flange centre, VIEW A-A
Const AA_Y = 5.479
Const DETB_X = 6.818             ' DETAIL B: the lower-left hole's centre
Const DETB_Y = 6.278
Const ISO_X = 9.551              ' isometric: half way between the flanges
Const ISO_Y = 6.947
Const NOTE_X = 0.45
Const NOTE_Y = 8.14
Const NOTE_DY = 0.167
Const NOTE_TAB = 0.41

Const DM_HIDDEN_VISIBLE = 1
Const DM_HIDDEN_REMOVED = 2
Const TOL_BASIC = 1              ' swTolType_e
Const TOL_LIMIT = 3
Const TOL_SYM = 4

' ---- the part (measured in FindOrientation) ----
Dim vR As Variant                ' square flange's outward normal          (right in the main view)
Dim vU As Variant                ' toward the round flange, square to vR   (up in the main view)
Dim vF As Variant                ' toward you in the main view
Dim pOrg As Variant              ' centre of the square flange's outer face (model, metres)
Dim nRf As Variant               ' round flange's outward normal (the VIEW A-A direction)
Dim eRf As Variant               ' along the round flange's face, in the main view's plane (up-right)
Dim pRo As Variant               ' round flange outer face centre, part inches (x, u, z)
Dim rFl As Double                ' round flange outside radius (3.750 / 2)
Dim rPipe As Double              ' pipe outside radius (torus minor radius)
Dim pBend As Variant             ' bend centre, part inches
Dim rBend As Double              ' bend radius (R6.000)
Dim faceAng As Double            ' angle between the two flange faces (45)
Dim rFil As Double               ' small fillet at the flanges (R .050)
Dim hRn As Long, hRx(4) As Double, hRu(4) As Double, hRz(4) As Double   ' round flange holes (part inches)

Dim viewMain As Object, viewRight As Object, viewAA As Object, viewDet As Object, viewIso As Object

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
    Set viewMain = Nothing: Set viewRight = Nothing: Set viewAA = Nothing: Set viewDet = Nothing: Set viewIso = Nothing
    Set lastDim = Nothing: eKey = "": skCandMain = -1: Set swModel = Nothing
    NoDimDialog True
    SpeedMode True

    Stage "Step 1 - find the open part":               FindPart
    Stage "Step 2 - measure the part":                 FindOrientation
    Stage "Step 3 - find your template":               tmpl = FindTemplate()
    Stage "Step 4 - new drawing":                      NewDrawing tmpl
    SpeedMode True
    Stage "Step 5 - main and right views":            MakeMainRight
    Stage "Step 6 - VIEW A-A":                         Set viewAA = MakeAA()
    Stage "Step 7 - isometric view":                   MakeIso
    Stage "Step 8 - main view dimensions and GD&T":    ClearCommand: DimMain
    Stage "Step 9 - right view dimensions":            ClearCommand: DimRight
    Stage "Step 10 - VIEW A-A dimensions and GD&T":    ClearCommand: DimAA
    Stage "Step 11 - DETAIL B":                        ClearCommand: MakeDetailB
    Stage "Step 12 - notes":                           ClearCommand: AddNotes
    Stage "Step 13 - title block":                     ClearCommand: TitleBlock
    Stage "Step 14 - centre marks":                    ClearCommand: CentreMarks
    Stage "Step 15 - scales and alignment":            CheckScalesAndAlignment
    Stage "Step 16 - finish":                          FinishSheet
    NoDimDialog False
    SpeedMode False

    MsgBox DRAW_NAME & " drawing built." & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & "      to do by hand: " & nTodo & vbCrLf & _
           IIf(nBad = 0, "", vbCrLf & "PROBLEMS:" & vbCrLf & Left(problems, 450)) & _
           IIf(nTodo = 0, "", vbCrLf & "TO DO BY HAND:" & vbCrLf & Left(todoList, 500)) & vbCrLf & _
           "Full list: " & FILE_TAG & "_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), DRAW_NAME

    Stage "Step 17 - save drawing and PDF":           SaveAll
    Exit Sub
Fail:
    errNo = Err.Number: errText = Err.Description       ' v2: before the clean-up below clears it
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
    PartMatches = (InStr(t, "DRAWINGS II") > 0 And InStr(t, "PART 3") > 0)
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

' point + a * direction (direction given in model vectors, result in part inches)
Function PAlong(p As Variant, dirM As Variant, ByVal a As Double) As Variant
    PAlong = Array(p(0) + a * Dot(dirM, vR), p(1) + a * Dot(dirM, vU), p(2) + a * Dot(dirM, vF))
End Function

' the square flange (4 x dia .375 holes), the round flange (5 x dia .257 holes), the bend (a torus)
Sub FindOrientation()
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, k As Long, s As Object, cp As Variant, r As Double
    Dim ax As Variant, q As Variant, n1 As Variant, n2 As Variant, p1 As Variant, p2 As Variant
    Dim c1 As Variant, c2 As Variant, k1 As Long, k2 As Long, pts1(15) As Variant, pts2(15) As Variant
    Dim t As Double, t1 As Double, t2 As Double, pp As Variant, nn As Variant, tp As Variant, d As Variant
    Dim uMin As Double, uMax As Double, zMin As Double, zMax As Double, cs As Double

    bodies = swPart.GetBodies2(0, True)
    If Not IsArray(bodies) Then Err.Raise vbObjectError + 3, , "The part has no solid body."
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder Then
                cp = s.CylinderParams
                r = cp(6) / IN_M
                ax = Unit(Arr3(cp(3), cp(4), cp(5)))
                q = Arr3(cp(0), cp(1), cp(2))
                If Abs(r - 0.1875) < 0.004 And k1 < 16 Then
                    If k1 = 0 Then n1 = ax: p1 = q
                    pts1(k1) = Sub3(q, Scl(n1, Dot(Sub3(q, p1), n1))): k1 = k1 + 1
                ElseIf Abs(r - 0.1285) < 0.004 And k2 < 16 Then
                    If k2 = 0 Then n2 = ax: p2 = q
                    pts2(k2) = Sub3(q, Scl(n2, Dot(Sub3(q, p2), n2))): k2 = k2 + 1
                End If
            End If
        Next
    Next
    If k1 = 0 Then Err.Raise vbObjectError + 4, , "No dia .375 holes found: this is not the In-Lab #8C part. Click into its window and run again."
    If k2 = 0 Then Err.Raise vbObjectError + 4, , "No dia .257 holes found on the round flange."
    c1 = Arr3(0, 0, 0): c2 = Arr3(0, 0, 0)
    For k = 0 To k1 - 1: c1 = Add3(c1, pts1(k)): Next
    For k = 0 To k2 - 1: c2 = Add3(c2, pts2(k)): Next
    c1 = Scl(c1, 1# / k1): c2 = Scl(c2, 1# / k2)
    If Dot(Sub3(c2, c1), n1) > 0 Then n1 = Scl(n1, -1)         ' square flange normal points away from the pipe
    If Dot(Sub3(c1, c2), n2) > 0 Then n2 = Scl(n2, -1)         ' round flange normal too
    vR = n1
    d = Sub3(c2, c1)
    vU = Unit(Sub3(d, Scl(n1, Dot(d, n1))))
    vF = Cross(vR, vU)
    nRf = n2
    eRf = Unit(Cross(nRf, vF))
    If Dot(eRf, Add3(vR, vU)) < 0 Then eRf = Scl(eRf, -1)          ' up-right in the main view

    ' outer faces: the flat faces square to each flange normal, furthest out
    t1 = -1E+30: t2 = -1E+30
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsPlane Then
                pp = s.PlaneParams
                nn = Unit(Arr3(pp(0), pp(1), pp(2)))
                If Abs(Dot(nn, n1)) > 0.9999 Then
                    t = Dot(Arr3(pp(3), pp(4), pp(5)), n1): If t > t1 Then t1 = t
                End If
                If Abs(Dot(nn, n2)) > 0.9999 Then
                    t = Dot(Arr3(pp(3), pp(4), pp(5)), n2): If t > t2 Then t2 = t
                End If
            ElseIf s.IsCylinder Then
                cp = s.CylinderParams
                If Abs(Dot(Unit(Arr3(cp(3), cp(4), cp(5))), n2)) > 0.999 And cp(6) / IN_M > 1.7 And cp(6) / IN_M < 2.2 Then
                    If cp(6) / IN_M > rFl Then rFl = cp(6) / IN_M
                End If
            ElseIf s.IsTorus Then
                tp = s.TorusParams                         ' centre(3), axis(3), major radius, minor radius
                If IsArray(tp) Then
                    If tp(6) / IN_M > 3# Then
                        pBend = Arr3(tp(0), tp(1), tp(2)): rBend = tp(6) / IN_M: rPipe = tp(7) / IN_M
                    ElseIf tp(7) / IN_M < 0.1 And tp(7) / IN_M > 0.01 Then
                        rFil = tp(7) / IN_M
                    End If
                End If
            End If
        Next
    Next
    pOrg = Add3(c1, Scl(n1, t1 - Dot(c1, n1)))
    pRo = ToPart(Add3(c2, Scl(n2, t2 - Dot(c2, n2))))
    If IsArray(pBend) Then pBend = ToPart(pBend)
    cs = -Dot(n1, n2): If cs > 1 Then cs = 1
    faceAng = Atn(Sqr(Abs(1 - cs * cs)) / IIf(Abs(cs) < 0.000001, 0.000001, cs)) / DEG
    If faceAng < 0 Then faceAng = faceAng + 180#

    ' hole patterns
    uMin = 1E+30: uMax = -1E+30: zMin = 1E+30: zMax = -1E+30
    For k = 0 To k1 - 1
        q = ToPart(pts1(k))
        If q(1) < uMin Then uMin = q(1)
        If q(1) > uMax Then uMax = q(1)
        If q(2) < zMin Then zMin = q(2)
        If q(2) > zMax Then zMax = q(2)
    Next
    hRn = 0
    For k = 0 To k2 - 1
        q = ToPart(Add3(pts2(k), Scl(n2, t2 - Dot(pts2(k), n2))))
        If HoleNew(q) And hRn < 5 Then hRx(hRn) = q(0): hRu(hRn) = q(1): hRz(hRn) = q(2): hRn = hRn + 1
    Next
    Check (Abs(uMax - uMin - 3.25) < 0.01 And Abs(zMax - zMin - 2.25) < 0.01), _
          "Square flange holes 3.250 x 2.250 apart (measured " & Format(uMax - uMin, "0.000") & " x " & Format(zMax - zMin, "0.000") & ")"
    Check (hRn = 5), hRn & " holes on the round flange (reference 5)"
    Check (Abs(faceAng - 45#) < 0.3), "Round flange face at " & Format(faceAng, "0.00") & " deg to the square flange (reference 45 BASIC)"
    Check (Abs(Dot(nRf, vF)) < 0.002), "Round flange square to the main view's plane (so VIEW A-A and the 45 deg read true)"
    Check (Abs(rFl - 1.875) < 0.01), "Round flange " & ChrW(216) & Format(2 * rFl, "0.000") & " (reference 3.755/3.745)"
    If rBend > 0 Then
        Info "Bend: R" & Format(rBend, "0.000") & ", pipe " & ChrW(216) & Format(2 * rPipe, "0.000") & ", centre at x " & Format(pBend(0), "0.000") & ", u " & Format(pBend(1), "0.000")
    Else
        Info "Bend torus not found - (R6.000) will be left for you"
    End If
    If rFil <= 0 Then rFil = 0.05
    Info "Small fillet R" & Format(rFil, ".000") & "; round flange outer face centre at x " & Format(pRo(0), "0.000") & ", u " & Format(pRo(1), "0.000")
End Sub

Function HoleNew(q As Variant) As Boolean
    Dim k As Long
    HoleNew = True
    For k = 0 To hRn - 1
        If Abs(hRx(k) - q(0)) + Abs(hRu(k) - q(1)) + Abs(hRz(k) - q(2)) < 0.02 Then HoleNew = False: Exit Function
    Next
End Function

' =====================================================================================
'  Run control, part, template  (from HW#8B)
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

' show the step in SolidWorks' status bar (the modelled threads make some steps slow)
Sub Stage(ByVal s As String)
    curStep = s
    On Error Resume Next
    swApp.Frame.SetStatusBarText DRAW_NAME & " macro: " & s & "   (modelled threads make SolidWorks slow - please wait)"
    DoEvents
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

Sub FindPart()
    ' Use the open part named "...Drawings II In-Lab Assignment Part 3". Otherwise the active part,
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
        "Could not tell which open part is In-Lab Assignment Part 3. Close the other parts, click into its window, and run again."
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
        Check (Abs(p(2) - 1) < 0.001 And Abs(p(3) - SHEET_DEN) < 0.001), "Sheet scale 1:2"
        Check (p(4) = 0 Or p(4) = False), "Sheet projection: third angle"
    End If
End Sub

' =====================================================================================
'  Views
' =====================================================================================
Sub MakeMainRight()
    Dim ok As Boolean
    Set viewMain = OrientedView("HW8C_Main", vR, vU, vF, MAIN_X - 1.5, MAIN_Y + 1#, StdFrontName(), False)
    If viewMain Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented main view."
    Check True, "Main view: square flange on the right (edge-on), elbow up-left to the round flange"
    EnsureSheetScale                                   ' auto-scale can change the sheet when the first view goes in
    ok = SetViewScale(viewMain, 1#, SHEET_DEN)
    Check ok, "Main view scale 1:2 (reads " & ScaleTxt(viewMain) & ")"
    SetMode viewMain, DM_HIDDEN_VISIBLE, "Main"
    PlaceView viewMain, 0#, 0#, 0#, MAIN_X, MAIN_Y, True, True, "Main view placed (square flange's outer face)"

    Set viewRight = Projected(viewMain, RIGHT_X, MAIN_Y)
    If viewRight Is Nothing Then
        Check False, "Right view - SolidWorks would not project it; add it with Projected View, right of the main view"
        Exit Sub
    End If
    ok = FrameOK(viewRight, vF, -1, 0, vU, 0, 1)
    Check ok, "Right view looks at the square flange face" & IIf(ok, "", " - it shows the wrong side; delete it and project it to the RIGHT")
    ok = SetViewScale(viewRight, 1#, SHEET_DEN)
    Check ok, "Right view scale 1:2 (reads " & ScaleTxt(viewRight) & ")"
    SetMode viewRight, DM_HIDDEN_REMOVED, "Right"
    PlaceView viewRight, 0#, 0#, 0#, RIGHT_X, MAIN_Y, True, False, "Right view placed (bore centre)"
End Sub

' VIEW A-A: auxiliary view square to the round flange, from its outer face edge in the main view.
' If SolidWorks will not make it, an ordinary view turned to look square at the face (same picture).
Function MakeAA() As Object
    On Error GoTo Oops
    Dim av As Object, before As String, m As Integer, sPick As Variant, how As String, n2 As Variant, ok As Boolean
    Dim ax As Variant, ay As Variant, pk As Variant
    CollectEdges viewMain
    pk = PAlong(pRo, eRf, 1.3)                         ' on the outer face, between the bore and the chamfer
    sPick = SP(viewMain, pk)
    For m = 1 To 2
        before = ViewNames()
        swModel.ClearSelection2 True
        swDraw.ActivateView viewMain.GetName2
        If m = 1 Then
            SelPt viewMain, pk, False
        Else
            On Error Resume Next
            swModel.Extension.SelectByID2 "", "EDGE", sPick(0) * IN_M, sPick(1) * IN_M, 0#, False, 0, Nothing, 0
            On Error GoTo Oops
        End If
        If swModel.SelectionManager.GetSelectedObjectCount2(-1) > 0 Then
            On Error Resume Next
            Set av = Nothing
            Set av = swDraw.CreateAuxiliaryViewAt2(AA_X * IN_M, AA_Y * IN_M, 0#, False)
            If av Is Nothing Then Err.Clear: Set av = swDraw.CreateAuxiliaryViewAt(AA_X * IN_M, AA_Y * IN_M, 0#, False)
            On Error GoTo Oops
            If av Is Nothing Then Set av = NewViewSince(before)
            If Not av Is Nothing Then how = IIf(m = 1, "edge selected directly", "edge picked on the sheet"): Exit For
        End If
    Next
    swModel.ClearSelection2 True
    If Not av Is Nothing Then
        n2 = ViewDir(av, nRf)
        If n2(2) * ZSign() < 0.99 Then                 ' not looking at the outer face: wrong side
            Check False, "VIEW A-A came out looking from the wrong side (" & VecTxt(n2) & ") - made again as a plain view"
            DeleteView av
            Set av = Nothing
        Else
            Check True, "VIEW A-A created as an auxiliary view (" & how & ")"
        End If
    End If
    If av Is Nothing Then
        ax = Unit(Add3(eRf, vF)): ay = Unit(Sub3(eRf, vF))      ' as an aligned auxiliary view shows it
        Set av = OrientedView("HW8C_AA", ax, ay, nRf, AA_X, AA_Y, "", False)
        If av Is Nothing Then
            Check False, "VIEW A-A: SolidWorks would not create it"
            Exit Function
        End If
        ToDo "VIEW A-A was made as a plain view square to the flange (SolidWorks refused the auxiliary command): add its arrows with Drawing > Auxiliary View if wanted, and type the label VIEW A-A under it"
    End If
    n2 = ViewDir(av, nRf)
    Check (n2(2) * ZSign() > 0.99), "VIEW A-A looks square onto the round flange's outer face"
    ok = SetViewScale(av, 1#, SHEET_DEN)
    Check ok, "VIEW A-A scale 1:2 (reads " & ScaleTxt(av) & ")"
    SetMode av, DM_HIDDEN_REMOVED, "VIEW A-A"
    PlaceView av, CDbl(pRo(0)), CDbl(pRo(1)), CDbl(pRo(2)), AA_X, AA_Y, True, True, "VIEW A-A placed (round flange centre)"
    Info "VIEW A-A label: SolidWorks names it by itself; if it does not read VIEW A-A, click the arrows and set the label to A"
    Set MakeAA = av
    Exit Function
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Function

Sub MakeIso()
    On Error GoTo Oops
    Dim dv As Variant, upI As Variant, rtI As Variant, ok As Boolean
    dv = Unit(Add3(Add3(vR, vU), vF))
    upI = Unit(Sub3(vU, Scl(dv, Dot(vU, dv))))
    rtI = Cross(upI, dv)
    Set viewIso = OrientedView("HW8C_Iso", rtI, upI, dv, ISO_X, ISO_Y, StdIsoName(), True)
    If viewIso Is Nothing Then
        Check False, "Isometric view - could not create it; add it by hand (View Palette > Isometric, scale 1:4)"
        Exit Sub
    End If
    Check True, "Isometric view"
    ok = SetViewScale(viewIso, 1#, ISO_DEN)
    Check ok, "Isometric view scale 1:4 (reads " & ScaleTxt(viewIso) & ")"
    SetMode viewIso, DM_HIDDEN_REMOVED, "Isometric"
    PlaceView viewIso, CDbl(pRo(0)) / 2, CDbl(pRo(1)) / 2, 0#, ISO_X, ISO_Y, True, True, "Isometric view placed (half way between the flanges)"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' DETAIL B (1:1) round the right view's lower-left hole, with its 4 X dia .375 THRU callout
Sub MakeDetailB()
    On Error GoTo Oops
    Dim cand As Integer, hp As Variant, c As Variant, rad As Double, seg As Object, dv As Object, before As String
    Dim combo As Integer, st As Integer, sh As Integer, i As Long, best As Long, tx As Double, ty As Double, why As String
    If viewRight Is Nothing Then Check False, "DETAIL B - no right view to take it from": Exit Sub
    hp = Array(0#, -1.625, 1.125)                     ' lower-left hole as seen in the right view (part inches)
    cand = FindCandP(viewRight, Array(0#, -1#, 1#), Array(0#, -1#, 1.5), 0.5, "Right view")
    If cand < 0 Then
        ToDo "DETAIL B: Drawing > Detail View, circle round the right view's lower-left hole, place it above the right view, label B, scale 1:1. Then add 4 X " & ChrW(216) & ".375 THRU on its hole"
        Exit Sub
    End If
    c = SP(viewRight, hp)
    rad = 0.774 * ViewSc(viewRight)
    Set seg = DrawViewCircle(viewRight, c, rad, cand)
    If seg Is Nothing Then
        ToDo "DETAIL B: the macro could not draw its circle - make it with Drawing > Detail View (label B, 1:1)"
        Exit Sub
    End If
    For combo = -1 To 5
        If combo >= 0 Then
            If SelectSeg(seg, c(0) + rad, c(1)) = 0 Then Exit For
        End If
        before = ViewNames()
        On Error Resume Next
        Set dv = Nothing
        If combo = -1 Then
            Set dv = swDraw.CreateDetailViewAt4(DETB_X * IN_M, DETB_Y * IN_M, 0#, 0, 1#, 1#, "B", 1, True, False, False, 5)
        Else
            st = (combo \ 3) Mod 5: sh = combo Mod 3
            Set dv = swDraw.CreateDetailViewAt4(DETB_X * IN_M, DETB_Y * IN_M, 0#, st, 1#, 1#, "B", sh, False, False, False, 5)
        End If
        Err.Clear
        On Error GoTo Oops
        If dv Is Nothing Then Set dv = NewViewSince(before)
        If Not dv Is Nothing Then Exit For
    Next
    swModel.ClearSelection2 True
    If dv Is Nothing Then
        DeleteSeg seg
        ToDo "DETAIL B: SolidWorks refused it - Drawing > Detail View, circle round the right view's lower-left hole, label B, scale 1:1; then 4 X " & ChrW(216) & ".375 THRU on its hole"
        Exit Sub
    End If
    swModel.EditRebuild3
    Set viewDet = dv
    Check SetViewScale(viewDet, 1#, 1#), "DETAIL B scale 1:1 (reads " & ScaleTxt(viewDet) & ")"
    PlaceView viewDet, CDbl(hp(0)), CDbl(hp(1)), CDbl(hp(2)), DETB_X, DETB_Y, True, True, "DETAIL B placed (hole centre)"
    ' 4 X dia .375 THRU on the detail's hole
    eKey = ""
    CollectEdges viewDet
    c = SP(viewDet, hp)
    best = -1
    For i = 0 To eN - 1
        If eKind(i) = "C" Then
            If Sqr((eCX(i) - c(0)) ^ 2 + (eCY(i) - c(1)) ^ 2) < 0.02 And Abs(eR(i) / ViewSc(viewDet) - 0.1875) < 0.006 Then best = i
        End If
    Next
    If best < 0 Then
        Check False, "DETAIL B: its hole was not found - DETAIL B may show the wrong place; check it"
        Exit Sub
    End If
    Check True, "DETAIL B shows the lower-left hole"
    swModel.ClearSelection2 True
    swDraw.ActivateView viewDet.GetName2
    SelEdge viewDet, best, False
    why = TryDimSel("D", 1, c(0) + 1.6, c(1) - 0.68, 0.375, "4 X ", 3, " THRU", False, True)
    ReportW why, "Detail B 4 X " & ChrW(216) & ".375 THRU"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' every view at its scale and where it was measured; the right view level with the main view
Sub CheckScalesAndAlignment()
    On Error GoTo Oops
    Dim a As Variant, b As Variant, ok As Boolean
    EnsureSheetScale
    a = SheetPt(viewMain, 0, 0, 0)
    ok = (Abs(a(0) - MAIN_X) < 0.002 And Abs(a(1) - MAIN_Y) < 0.002)
    Check ok, "Alignment: main view at " & Format(a(0), "0.000") & ", " & Format(a(1), "0.000")
    If Not viewRight Is Nothing Then
        b = SheetPt(viewRight, 0, 0, 0)
        Check (Abs(b(1) - a(1)) < 0.001), "Alignment: right view level with the main view (" & Format(a(1), "0.000") & " / " & Format(b(1), "0.000") & ")"
    End If
    If Not viewAA Is Nothing Then
        b = SP(viewAA, pRo)
        Check (Abs(b(0) - AA_X) < 0.01 And Abs(b(1) - AA_Y) < 0.01), "Alignment: VIEW A-A at " & Format(b(0), "0.000") & ", " & Format(b(1), "0.000") & _
              " (reference " & Format(AA_X, "0.000") & ", " & Format(AA_Y, "0.000") & ")"
    End If
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' centre marks: the right view's four holes and bore, VIEW A-A's five holes and bore
Sub CentreMarks()
    On Error GoTo Oops
    Dim n As Long, i As Long, k As Long, c As Variant, cm As Object, vv As Variant, v As Object, want As Long
    For Each vv In Array(viewRight, viewAA)
        Set v = Nothing
        If IsObject(vv) Then Set v = vv                ' (an Object variable: helpers take views ByRef As Object)
        If Not v Is Nothing Then
            n = 0: want = 0
            eKey = ""
            CollectEdges v
            For i = 0 To eN - 1
                If eKind(i) = "C" And eFull(i) Then
                    If Abs(eR(i) / ViewSc(v) - 0.1875) < 0.006 Or Abs(eR(i) / ViewSc(v) - 0.1285) < 0.006 Or Abs(eR(i) / ViewSc(v) - 1.125) < 0.006 Then
                        want = want + 1
                        swModel.ClearSelection2 True
                        swDraw.ActivateView v.GetName2
                        SelEdge v, i, False
                        Set cm = Nothing
                        On Error Resume Next
                        Set cm = swDraw.InsertCenterMark3(0, False, True)
                        On Error GoTo Oops
                        If Not cm Is Nothing Then n = n + 1
                    End If
                End If
            Next
            swModel.ClearSelection2 True
            Check (n = want And n > 0), IIf(v Is viewRight, "Right", "VIEW A-A") & "  centre marks on " & n & " of " & want & " holes and bore"
        End If
    Next
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' +1 if a view's z points toward you (proved on the main view, where vF faces you), else -1
Function ZSign() As Double
    Dim z As Variant
    z = ViewDir(viewMain, vF)
    If z(2) < 0 Then ZSign = -1# Else ZSign = 1#
End Function

' =====================================================================================
'  View helpers  (from HW#8B)
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
'  Dimensions and GD&T  (text positions measured on the reference, sheet inches)
' =====================================================================================
Sub DimMain()
    On Error GoTo Oops
    Dim o As Variant, ro As Variant, yb As Double, why As String, D As String, i As Long, cand As Integer
    Dim pOut As Variant, pIn As Variant, s As Variant, e As Variant, c As Variant, sS As Variant, sE As Variant, sC As Variant
    Dim arc As Object, dirn As Integer, x As Double
    D = ChrW(216)
    eKey = ""
    CollectEdges viewMain
    o = SP(viewMain, Array(0#, 0#, 0#)): ro = SP(viewMain, pRo)
    yb = SP(viewMain, Array(0#, -2#, 0#))(1)

    ' square flange: height and thickness
    DimT viewMain, "V", Array(-0.1875, 2#, 0#), Array(-0.1875, -2#, 0#), o(0) + 0.832, o(1), 4#, "", 2, "", False, TOL_SYM, -0.02, 0.02, 2, "Main     4.00" & ChrW(177) & ".02"
    DimT viewMain, "H", Array(-0.375, -1.8, 0#), Array(0#, -1.8, 0#), o(0) + 0.58, yb - 0.313, 0.375, "", 3, "", False, 0, 0, 0, 3, "Main     .375 (square flange)"

    ' round flange: dia 3.755/3.745 on its outside edge (seen side-on), thickness .375, 45 deg BASIC, datum A
    pOut = PAlong(pRo, eRf, 1.72)
    pIn = PAlong(PAlong(pRo, nRf, -0.375), eRf, 1.72)
    i = LineByLen(viewMain, 2 * rFl, eRf, pRo)
    If i < 0 Then
        Check False, "Main     " & D & "3.755/3.745 - MISSING (no " & Format(2 * rFl, "0.000") & " long edge along the round flange)"
    Else
        swModel.ClearSelection2 True
        swDraw.ActivateView viewMain.GetName2
        SelEdge viewMain, i, False
        why = TryDimSel("S1", 1, ro(0) - 0.139, ro(1) + 0.494, 2 * rFl, "<MOD-DIAM>", 3, "", False, True)
        If Len(why) = 0 Then why = SetTol(lastDim, TOL_LIMIT, -0.005, 0.005, 3, 3)
        ReportW why, "Main     " & D & "3.755/3.745 (round flange)"
    End If
    DimT viewMain, "S", pOut, pIn, ro(0) - 0.167, ro(1) - 1.04, 0.375, "", 3, "", False, 0, 0, 0, 3, "Main     .375 (round flange)"
    why = DimP2(viewMain, "A", pOut, Array(0#, 1#, 0#), o(0) - 0.846, o(1) + 1.467, faceAng, "", 0, "", False)
    If Len(why) > 0 Then why = DimP2(viewMain, "A", pOut, Array(0#, 1#, 0#), o(0) - 0.6, o(1) + 2.3, faceAng, "", 0, "", False)
    If Len(why) = 0 Then why = SetTol(lastDim, TOL_BASIC, 0, 0, 0, 0)
    ReportW why, "Main     45" & ChrW(176) & " BASIC (boxed)"
    DatumOnEdge viewMain, PAlong(pRo, eRf, -1.6), ro(0) - 1.33, ro(1) - 0.62, "A", "Main     datum A on the round flange face"

    ' GD&T and notes
    MakeFrame viewMain, Array(0#, 1.7, 0#), "<IGTOL-ANGULAR>", False, ".005", "A", "", o(0) + 0.693, o(1) + 1.443, False, "Main     angularity .005 | A"
    NoteAt viewMain, Array(-0.35, -1.975, 0#), Array(-0.375, -1.9, 0#), ".050 X 45<MOD-DEG> TYP", o(0) - 1.748, yb - 0.113, "Main     .050 X 45" & ChrW(176) & " TYP"
    x = -(rPipe + rFil) - 0.02
    NoteAt viewMain, Array(-0.375, x, 0#), Array(-0.375, -1.8, 0#), "2 X R " & Format(rFil, ".000"), o(0) - 1.574, o(1) - 0.867, "Main     2 X R " & Format(rFil, ".000")

    ' (R6.000) on the bend's centre line: an exact sketch arc (placement proved first)
    If rBend <= 0 Then
        ToDo "Main (R6.000): Smart Dimension the pipe's centre-line arc, add parentheses (3 decimals)"
        Exit Sub
    End If
    cand = FindCandP(viewMain, Array(-1#, 1#, 0#), Array(0#, 1#, 0#), 1#, "Main view")
    skCandMain = cand
    If cand < 0 Then
        ToDo "Main (R6.000): Smart Dimension the pipe's centre-line arc, add parentheses (3 decimals)"
        Exit Sub
    End If
    c = pBend
    s = OnCircle(c, Array(-0.375, 0#, 0#), rBend)
    e = OnCircle(c, PAlong(pRo, nRf, -0.375), rBend)
    sC = SP(viewMain, c): sS = SP(viewMain, s): sE = SP(viewMain, e)
    If (sS(0) - sC(0)) * (sE(1) - sC(1)) - (sS(1) - sC(1)) * (sE(0) - sC(0)) > 0 Then dirn = 1 Else dirn = -1
    Set arc = SkArcP(viewMain, c, s, e, dirn, cand)
    why = "MISSING (centre-line arc not made)"
    If Not arc Is Nothing Then
        swModel.ClearSelection2 True
        swDraw.ActivateView viewMain.GetName2
        On Error Resume Next
        arc.Select4 False, Nothing
        On Error GoTo Oops
        why = TryDimSel("R", 1, o(0) - 0.361, o(1) + 2.013, rBend, "", 3, "", True, True)
        CentreStyle arc
    End If
    If Len(why) > 0 Then DeleteSeg arc
    ReportW why, "Main     (R" & Format(rBend, "0.000") & ") bend, on the pipe's centre line"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' point at distance r from c (both part inches, in the main view's plane), toward p
Function OnCircle(c As Variant, p As Variant, ByVal r As Double) As Variant
    Dim dx As Double, du As Double, L As Double
    dx = p(0) - c(0): du = p(1) - c(1): L = Sqr(dx * dx + du * du)
    If L < 0.000001 Then L = 1
    OnCircle = Array(c(0) + r * dx / L, c(1) + r * du / L, 0#)
End Function

Sub DimRight()
    On Error GoTo Oops
    Dim rc As Variant
    If viewRight Is Nothing Then Exit Sub
    eKey = ""
    CollectEdges viewRight
    rc = SP(viewRight, Array(0#, 0#, 0#))
    DimT viewRight, "H", Array(0#, -1#, 1.5), Array(0#, -1#, -1.5), rc(0) + 0.012, rc(1) - 1.79, 3#, "", 2, "", False, TOL_SYM, -0.02, 0.02, 2, "Right    3.00" & ChrW(177) & ".02"
    DimT viewRight, "H", Array(0#, -1.8125, 1.125), Array(0#, -1.8125, -1.125), rc(0) + 0.04, rc(1) - 1.519, 2.25, "", 3, "", False, TOL_SYM, -0.004, 0.004, 3, "Right    2.250" & ChrW(177) & ".004"
    DimT viewRight, "H", Array(0#, -1.125, 0#), Array(0#, -1.8125, -1.125), rc(0) - 0.27, rc(1) - 1.272, 1.125, "", 3, "", False, 0, 0, 0, 3, "Right    1.125"
    DimT viewRight, "V", Array(0#, 1.625, 1.3125), Array(0#, -1.625, 1.3125), rc(0) - 1.86, rc(1) + 0.025, 3.25, "", 3, "", False, TOL_SYM, -0.004, 0.004, 3, "Right    3.250" & ChrW(177) & ".004"
    DimT viewRight, "V", Array(0#, 1.625, 1.3125), Array(0#, 1.125, 0#), rc(0) - 1.2, rc(1) + 0.35, 1.625, "", 3, "", False, 0, 0, 0, 3, "Right    1.625"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

Sub DimAA()
    On Error GoTo Oops
    Dim ac As Variant, why As String, D As String, k As Long, kA As Long, kB As Long, a As Double, bestA As Double, bestB As Double
    Dim hp As Variant, hs As Variant, cand As Integer, seg As Object, l1 As Object, l2 As Object, angs(4) As Double
    Dim dimHole As Object, dimBore As Object, sc As Double, p1 As Variant, p2 As Variant
    D = ChrW(216)
    If viewAA Is Nothing Then Check False, "VIEW A-A dimensions - no VIEW A-A": Exit Sub
    eKey = ""
    CollectEdges viewAA
    ac = SP(viewAA, pRo)
    sc = ViewSc(viewAA)

    ' dia 2.250 +-.005 (the bore) with datum B
    why = DimP2(viewAA, "D", PAlong(pRo, vF, 1.125), Empty, ac(0) - 1.415, ac(1) - 1.039, 2.25, "", 3, "", False)
    If Len(why) = 0 Then
        Set dimBore = lastDim
        why = SetTol(lastDim, TOL_SYM, -0.005, 0.005, 3, 3)
    End If
    ReportW why, "VIEW A-A " & D & "2.250" & ChrW(177) & ".005 (bore)"
    DatumOnDim dimBore, ac(0) - 0.724, ac(1) - 1.459, "B", "VIEW A-A datum B on the " & D & "2.250"

    ' the holes' angles on the sheet; the 5 X dia .257 callout goes on the one at the top (about 82 deg)
    bestA = 1E+30: kA = -1
    For k = 0 To hRn - 1
        hs = SP(viewAA, Array(hRx(k), hRu(k), hRz(k)))
        angs(k) = AngDeg(hs(0) - ac(0), hs(1) - ac(1))
        a = Abs(angs(k) - 82#): If a > 180 Then a = 360 - a
        If a < bestA Then bestA = a: kA = k
    Next
    ' its counter-clockwise neighbour (about 72 deg further round) for 5 X 72 deg
    bestB = 1E+30: kB = -1
    For k = 0 To hRn - 1
        If k <> kA Then
            a = angs(k) - angs(kA): If a < 0 Then a = a + 360
            If Abs(a - 72) < bestB Then bestB = Abs(a - 72): kB = k
        End If
    Next
    If kA < 0 Then Check False, "VIEW A-A holes not found": Exit Sub
    hp = Array(hRx(kA), hRu(kA), hRz(kA))
    why = DimP2(viewAA, "D", PAlong(hp, vF, 0.1285), Empty, ac(0) + 2.18, ac(1) + 1.85, 0.257, "5 X ", 3, "", False)
    If Len(why) = 0 Then Set dimHole = lastDim
    ReportW why, "VIEW A-A 5 X " & D & ".257"
    Set lastDim = dimHole
    MakeFrame viewAA, Empty, "<IGTOL-POSI>", True, ".01", "A", "B", ac(0) + 1.78, ac(1) + 1.73, True, "VIEW A-A position " & D & ".01 | A | B under 5 X " & D & ".257"

    ' dia 3.150 BASIC bolt circle and 5 X 72 deg BASIC: exact sketch items (placement proved first)
    cand = FindCandAA(viewAA, hp)
    If cand < 0 Then
        ToDo "VIEW A-A " & D & "3.150 BASIC and 5 X 72" & ChrW(176) & " BASIC: sketch a centre-line circle through the holes and dimension it (Tolerance: Basic); dimension the angle between two holes' centre lines, add ""5 X "", Tolerance: Basic"
        Exit Sub
    End If
    Set seg = DrawViewCircle(viewAA, ac, 1.575 * sc, cand)
    why = "MISSING (bolt circle not drawn)"
    If Not seg Is Nothing Then
        eKey = ""
        CentreStyle seg
        swModel.ClearSelection2 True
        swDraw.ActivateView viewAA.GetName2
        On Error Resume Next
        seg.Select4 False, Nothing
        On Error GoTo Oops
        why = TryDimSel("D", 1, ac(0) + 2.13, ac(1) + 1.144, 3.15, "", 3, "", False, True)
        If Len(why) = 0 Then why = SetTol(lastDim, TOL_BASIC, 0, 0, 3, 3)
        If Len(why) > 0 And InStr(why, "WRONG") > 0 Then DeleteSeg seg
    End If
    ReportW why, "VIEW A-A " & D & "3.150 BASIC (boxed) bolt circle"
    If kB < 0 Then Check False, "VIEW A-A 5 X 72" & ChrW(176) & " - second hole not found": Exit Sub
    p1 = Array(hRx(kA), hRu(kA), hRz(kA)): p2 = Array(hRx(kB), hRu(kB), hRz(kB))
    Set l1 = SkLineP(viewAA, Toward(pRo, p1, 0.9), Toward(pRo, p1, 3.6), cand)
    Set l2 = SkLineP(viewAA, Toward(pRo, p2, 0.9), Toward(pRo, p2, 3.6), cand)
    why = "MISSING (centre lines not made)"
    If Not l1 Is Nothing And Not l2 Is Nothing Then
        swModel.ClearSelection2 True
        swDraw.ActivateView viewAA.GetName2
        On Error Resume Next
        l1.Select4 False, Nothing
        l2.Select4 True, Nothing
        On Error GoTo Oops
        why = TryDimSel("A", 2, ac(0) - 0.964, ac(1) + 1.572, 72#, "5 X ", 0, "", False, True)
        If Len(why) = 0 Then why = SetTol(lastDim, TOL_BASIC, 0, 0, 0, 0)
    End If
    If Len(why) > 0 Then DeleteSeg l1: DeleteSeg l2
    ReportW why, "VIEW A-A 5 X 72" & ChrW(176) & " BASIC (boxed)"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' point at distance d from c, toward p (part inches)
Function Toward(c As Variant, p As Variant, ByVal d As Double) As Variant
    Dim v(2) As Double, L As Double, k As Integer
    For k = 0 To 2: v(k) = p(k) - c(k): Next
    L = Sqr(v(0) ^ 2 + v(1) ^ 2 + v(2) ^ 2)
    If L < 0.000001 Then L = 1
    Toward = Array(c(0) + d * v(0) / L, c(1) + d * v(1) / L, c(2) + d * v(2) / L)
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

' =====================================================================================
'  Picking, dimension and GD&T helpers
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

' the straight edge of length L (part inches) along model direction dirM nearest part point nearP
Function LineByLen(v As Object, ByVal L As Double, dirM As Variant, nearP As Variant) As Long
    Dim i As Long, a As Variant, b As Variant, dx As Double, dy As Double, dl As Double, ux As Double, uy As Double
    Dim sc As Double, n As Variant, best As Double, d As Double, mx As Double, my As Double
    LineByLen = -1
    CollectEdges v
    sc = ViewSc(v)
    a = SP(v, nearP): b = SP(v, PAlong(nearP, dirM, 1#))
    ux = b(0) - a(0): uy = b(1) - a(1): dl = Sqr(ux * ux + uy * uy)
    If dl < 0.000001 Then Exit Function
    ux = ux / dl: uy = uy / dl
    best = 1E+30
    For i = 0 To eN - 1
        If eKind(i) = "L" Then
            dx = eX2(i) - eX1(i): dy = eY2(i) - eY1(i): dl = Sqr(dx * dx + dy * dy)
            If dl > 0.01 Then
                If Abs(dx / dl * uy - dy / dl * ux) < 0.01 And Abs(dl / sc - L) < 0.01 Then
                    mx = (eX1(i) + eX2(i)) / 2: my = (eY1(i) + eY2(i)) / 2
                    d = (mx - a(0)) ^ 2 + (my - a(1)) ^ 2
                    If d < best Then best = d: LineByLen = i
                End If
            End If
        End If
    Next
End Function

' a feature control frame. onDim: attached to lastDim (sits under its text); else its leader goes to
' the edge at lp (one click there). Judged on what reads back, never on the call's return value.
Sub MakeFrame(v As Object, lp As Variant, ByVal sym As String, ByVal dia As Boolean, ByVal tol1 As String, _
              ByVal d1 As String, ByVal d2 As String, ByVal tx As Double, ByVal ty As Double, ByVal onDim As Boolean, ByVal label As String)
    Dim g As Object, attached As Boolean, rb As String
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
        g.SetFrameValues2 1, tol1, "", d1, d2, ""
        If Err.Number <> 0 Then Err.Clear: g.SetFrameValues 1, tol1, "", d1, d2, ""
        If Not (onDim And attached) Then g.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
        rb = GtolReadBack(g)
    End If
    On Error GoTo 0
    swModel.ClearSelection2 True
    If g Is Nothing Then
        Check False, label & " - MISSING; add it with Annotations > Geometric Tolerance"
        Exit Sub
    End If
    If Len(rb) = 0 Then
        Info label & " placed; SolidWorks does not report its values - look that it reads " & tol1 & " | " & d1 & IIf(Len(d2) > 0, " | " & d2, "")
    ElseIf InStr(rb, tol1) > 0 And InStr(rb, d1) > 0 And InStr(rb, d2) > 0 Then
        Check True, label & " (reads " & rb & ")"
    Else
        Check False, label & " - reads " & rb & "; double-click it and set " & tol1 & ", " & d1 & IIf(Len(d2) > 0, ", " & d2, "")
    End If
    Info "  symbol " & sym & IIf(SymInLibrary(sym), " found in your gtol.sym", " not confirmed in gtol.sym - look that the symbol shows")
    If Not attached Then ToDo label & ": the frame is placed without " & IIf(onDim, "being attached to its dimension - drag it under the dimension", "a leader - click it and drag its leader handle to the face as on the reference")
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

' datum feature symbol on a dimension, its box at (tx, ty)
Sub DatumOnDim(dd As Object, ByVal tx As Double, ByVal ty As Double, ByVal lbl As String, ByVal label As String)
    Dim dt As Object
    Set dt = Nothing
    If dd Is Nothing Then Check False, label & " - MISSING (its dimension was not made)": Exit Sub
    swModel.ClearSelection2 True
    On Error Resume Next
    dd.GetAnnotation.Select3 False, Nothing
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) = 1 Then Set dt = swModel.InsertDatumTag2
    If dt Is Nothing Then
        swModel.ClearSelection2 True
        swModel.Extension.SelectByID2 dd.GetNameForSelection, "DIMENSION", 0#, 0#, 0#, False, 0, Nothing, 0
        If swModel.SelectionManager.GetSelectedObjectCount2(-1) = 1 Then Set dt = swModel.InsertDatumTag2
    End If
    If Not dt Is Nothing Then dt.SetLabel lbl: dt.GetAnnotation.SetPosition2 tx * IN_M, ty * IN_M, 0#
    On Error GoTo 0
    swModel.ClearSelection2 True
    If dt Is Nothing Then
        Check False, label & " - MISSING; select the dimension, Annotations > Datum Feature, label " & lbl
    Else
        Check (UCase(Trim(DatumLabel(dt))) = lbl), label & " (label reads " & DatumLabel(dt) & ")"
    End If
End Sub

' =====================================================================================
'  Sketch items at part points (3-D part inches), and proving a view's sketch placement
' =====================================================================================
Function SkPointP(v As Object, p As Variant, ByVal cand As Integer) As Object
    Dim q As Variant
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    q = ToSk(v, SP(v, p), cand)
    If Not IsArray(q) Then Exit Function
    On Error Resume Next
    swModel.SketchManager.AddToDB = True
    Set SkPointP = swModel.SketchManager.CreatePoint(q(0), q(1), 0#)
    swModel.SketchManager.AddToDB = False
    eKey = ""
    swModel.ClearSelection2 True
End Function

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

Function SkArcP(v As Object, c As Variant, s As Variant, e As Variant, ByVal dirn As Integer, ByVal cand As Integer) As Object
    Dim pc As Variant, ps As Variant, pe As Variant
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    pc = ToSk(v, SP(v, c), cand): ps = ToSk(v, SP(v, s), cand): pe = ToSk(v, SP(v, e), cand)
    If Not IsArray(pc) Or Not IsArray(ps) Or Not IsArray(pe) Then Exit Function
    On Error Resume Next
    swModel.SketchManager.AddToDB = True
    Set SkArcP = swModel.SketchManager.CreateArc(pc(0), pc(1), 0#, ps(0), ps(1), 0#, pe(0), pe(1), 0#, dirn)
    swModel.SketchManager.AddToDB = False
    eKey = ""
    swModel.ClearSelection2 True
End Function

' which sketch method puts items where they belong in view v: a test point at part point p must
' measure dWant (horizontal) from the edge through pEdge
Function FindCandP(v As Object, p As Variant, pEdge As Variant, ByVal dWant As Double, ByVal label As String) As Integer
    Dim cand As Variant, pt As Object, d As Double, tried As String
    FindCandP = -1
    For Each cand In Array(2, 3, 1, 0)
        Set pt = SkPointP(v, p, CInt(cand))
        If pt Is Nothing Then
            tried = tried & "m" & cand & ": no point. "
        Else
            d = TestH(v, pt, pEdge)
            DeleteSeg pt
            If Abs(d - dWant) < 0.003 Then
                FindCandP = CInt(cand)
                Info label & ": sketch method " & cand & " (test point measured " & Format(d, "0.000") & ")"
                Exit Function
            End If
            tried = tried & "m" & cand & ": " & Format(d, "0.000") & ". "
        End If
    Next
    Info label & ": no sketch method matched (" & Trim(tried) & ")"
End Function

' VIEW A-A is turned on the sheet, so its test measures the straight distance from a sketch point at
' the flange centre to a bolt hole's centre (1.575)
Function FindCandAA(v As Object, hp As Variant) As Integer
    Dim cand As Variant, pt As Object, d As Double, tried As String, dd As Object
    FindCandAA = -1
    For Each cand In Array(2, 3, 1, 0)
        Set pt = SkPointP(v, pRo, CInt(cand))
        d = -1
        If Not pt Is Nothing Then
            CollectEdges v
            swModel.ClearSelection2 True
            swDraw.ActivateView v.GetName2
            On Error Resume Next
            pt.Select4 False, Nothing
            On Error GoTo 0
            If SelPt(v, PAlong(hp, vF, 0.1285), True) Then
                If swModel.SelectionManager.GetSelectedObjectCount2(-1) = 2 Then
                    Set dd = Nothing
                    On Error Resume Next
                    Set dd = swModel.AddDimension2((SP(v, pRo)(0) - 1.5) * IN_M, (SP(v, pRo)(1) - 1.5) * IN_M, 0#)
                    If Not dd Is Nothing Then d = dd.GetDimension2(0).SystemValue / IN_M
                    On Error GoTo 0
                    swModel.ClearSelection2 True
                    If Not dd Is Nothing Then
                        If Not DeleteDim(dd) Then Check False, "A test dimension (" & Format(d, "0.000") & ") in VIEW A-A could not be deleted - click it and press Delete"
                    End If
                End If
            End If
            swModel.ClearSelection2 True
            DeleteSeg pt
            If Abs(d - 1.575) < 0.003 Then
                FindCandAA = CInt(cand)
                Info "VIEW A-A: sketch method " & cand & " (test measured " & Format(d, "0.000") & ")"
                Exit Function
            End If
        End If
        tried = tried & "m" & cand & ": " & Format(d, "0.000") & ". "
    Next
    Info "VIEW A-A: no sketch method matched (" & Trim(tried) & ")"
End Function

' =====================================================================================
'  Dimension, selection, GD&T and note helpers  (from HW#8B / HW#8A)
' =====================================================================================
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

Sub CentreStyle(seg As Object)
    Dim ok As Boolean
    On Error Resume Next
    swModel.ClearSelection2 True
    seg.Select4 False, Nothing
    ok = swModel.SetLineStyle("Center")
    If Not ok Then ok = swModel.SetLineStyle("CENTER")
    swModel.ClearSelection2 True
End Sub

' NOTES block, top left: each number and each text is its own note so the columns line up exactly
Sub AddNotes()
    On Error GoTo Oops
    Dim lines As Variant, i As Long, y As Double, ok As Boolean
    lines = Array("USE THE IN-LAB ASSIGNMENT #8C PART MODEL ON CANVAS TO RE-CREATE THE ENGINEERING DRAWING", _
                  "MAKE SURE YOUR DRAWING'S TOLERANCES MATCH THE DRAWING TOLERANCES SHOWN BELOW.")
    ok = PlaceNote("NOTES:", NOTE_X, NOTE_Y)
    For i = 0 To 1
        y = NOTE_Y - (i + 1) * NOTE_DY
        ok = PlaceNote((i + 1) & ".", NOTE_X, y) And ok
        ok = PlaceNote(CStr(lines(i)), NOTE_X + NOTE_TAB, y) And ok
    Next
    Check ok, "NOTES 1-2 top left (each read back)"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.SketchManager.AddToDB = False
    swModel.ClearSelection2 True
End Sub

' =====================================================================================
'  Edge table  (from HW#8B)
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

Sub AddEdge(ByVal ent As Object, xf As Object, ByVal sc As Double)
    ' Extents come only from exact geometry: line end points, and circle/arc centre +- radius.
    ' (For this part every arc's full circle stays inside the outline, so centre +- radius is exact.)
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

Sub Extend(p As Variant)
    If p(0) < xMin Then xMin = p(0)
    If p(0) > xMax Then xMax = p(0)
    If p(1) < yMin Then yMin = p(1)
    If p(1) > yMax Then yMax = p(1)
End Sub

' =====================================================================================
'  Title block, finish, save  (from HW#8B)
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
                nt.SetText "SHEET 3 OF 4"
                If UCase(Trim(nt.GetText)) <> "SHEET 3 OF 4" Then nt.PropertyLinkedText = "SHEET 3 OF 4"
                sheetDone = (UCase(Trim(nt.GetText)) = "SHEET 3 OF 4")
            End If
            If Left(t, 6) = "SCALE:" And Not scaleDone Then              ' as on the reference
                If InStr(t, "1:4") = 0 Then nt.SetText "SCALE: 1:4"
                If InStr(UCase(nt.GetText), "1:4") = 0 Then nt.PropertyLinkedText = "SCALE: 1:4"
                scaleDone = (InStr(UCase(nt.GetText), "1:4") > 0)
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

    AddSheetNote "CAST IRON", mX + 0.06, mY - 0.17, "Material: CAST IRON"
    AddSheetNote Format(Date, "mm-dd-yy"), dX + 0.4, dY - 0.18, "Date: " & Format(Date, "mm-dd-yy")
    AddSheetNote "IME 1140 ENGINEERING DRAWINGS II", tX + 0.44, tY, "Title line 1"
    AddSheetNote "IN-LAB ASSIGNMENT #8C", tX + 0.44, tY - 0.17, "Title line 2"
    Check scaleDone, "SCALE: 1:4 (as on the reference)" & IIf(scaleDone, "", " - edit the scale note by hand")
    Check sheetDone, "SHEET 3 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
    Check (twoOK And threeOK), "Tolerance block reads TWO PLACE " & ChrW(177) & ".01, THREE PLACE " & ChrW(177) & ".008 (note 2, as on the reference)" & _
          IIf(twoOK And threeOK, "", IIf(tolSeen, " - edit it by hand (Edit Sheet Format)", " - tolerance note not found; check it by hand"))
End Sub

' make the TWO / THREE PLACE DECIMAL lines read +-.01 / +-.008 (the reference) and read them back
Sub FixTol(nt As Object, twoOK As Boolean, threeOK As Boolean)
    Dim t As String, u As String
    On Error Resume Next
    t = nt.GetText
    If InStr(UCase(LineOf(UCase(t), "TWO PLACE DECIMAL")), ".01") = 0 Or InStr(UCase(LineOf(UCase(t), "THREE PLACE DECIMAL")), ".008") = 0 Then
        t = NewLine(t, "TWO PLACE DECIMAL", "TWO PLACE DECIMAL <MOD-PM> .01")
        t = NewLine(t, "THREE PLACE DECIMAL", "THREE PLACE DECIMAL  <MOD-PM> .008")
        nt.SetText t
    End If
    u = UCase(nt.GetText)
    If InStr(u, "TWO PLACE DECIMAL") > 0 Then twoOK = (InStr(LineOf(u, "TWO PLACE DECIMAL"), ".01") > 0 And InStr(LineOf(u, "TWO PLACE DECIMAL"), ".02") = 0)
    If InStr(u, "THREE PLACE DECIMAL") > 0 Then threeOK = (InStr(LineOf(u, "THREE PLACE DECIMAL"), ".008") > 0)
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
    For Each v In Array(viewMain, viewRight, viewAA, viewDet, viewIso)
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
    If n >= 0 Then Check (n >= 5 Or (n >= 4 And viewDet Is Nothing)), n & " views on the sheet (expected 5)"
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
'  Geometry helpers  (from HW#8B)
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
