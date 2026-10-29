' =====================================================================================
'  HW#8A  -  ENGRAVE THE LOGO .003 DEEP (note 3)  -  SolidWorks PART macro   (inches)
'
'  HOW TO RUN  (run this BEFORE HW8A_Drawing_Macro.bas)
'   1. Open "IME 1140 Engineering Drawings II In-Lab Assignment Part 1.SLDPRT" in SolidWorks.
'   2. Tools > Macro > New...  save it as  HW8A_Engrave.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main.  It asks before saving the part.
'
'  WHAT IT DOES
'   The Canvas part already has the logo cut into the pocket floor: the MANUFACTURING /
'   ENGINEERING text (Cut-Extrude2) and the CP symbol (Cut-Extrude3), but .005 deep.
'   Note 3 asks for .003.  This macro finds every shallow cut (under .02 deep), sets it to
'   .003, rebuilds, then checks:
'    - every engraved floor face sits .003 below the pocket floor
'    - the part got heavier by exactly (engraved area) x (old depth - .003)
'   It writes HW8A_engrave_checks.txt next to the part.
'
'  v1: first version
' =====================================================================================
Option Explicit

Dim swApp As Object, swPart As Object, partPath As String
Dim curStep As String, report As String, problems As String, nOK As Long, nBad As Long
Dim vF As Variant                ' pocket axis, back face -> pocket (model unit vector)
Dim fBack As Double              ' back face position along vF (metres)

Const IN_M = 0.0254
Const NEW_DEPTH = 0.003          ' note 3
Const F_FLOOR = 0.4              ' pocket floor, from the back face
Const PKT = 1.75                 ' pocket radius (dia 3.50)

Sub main()
    On Error GoTo Fail
    Dim cuts(19) As Object, nCut As Long, i As Long, oldD As Double
    Dim v0 As Double, v1 As Double, want As Double, got As Double
    Dim dMin As Double, dMax As Double, nF As Long, area As Double
    Set swApp = Application.SldWorks
    report = "": problems = "": nOK = 0: nBad = 0

    curStep = "Step 1 - find the open part":           FindPart
    curStep = "Step 2 - read the pocket":              FindPocket
    curStep = "Step 3 - find the logo cuts":           nCut = FindCuts(cuts, oldD)
    If nCut = 0 Then
        Check False, "No shallow cut (the logo) found in this part. Tell Claude: the logo has to be added, not just set to .003."
        Finish False
        Exit Sub
    End If
    EngravedFloors dMin, dMax, nF, area
    Check (nF > 0), "Logo on the pocket floor: " & nF & " engraved faces, now " & Format(dMin, ".0000") & _
          IIf(Abs(dMax - dMin) > 0.00005, " to " & Format(dMax, ".0000"), "") & " deep"
    v0 = Volume()

    curStep = "Step 4 - set the logo depth to .003"
    For i = 0 To nCut - 1
        SetDepth cuts(i)
    Next
    swPart.ForceRebuild3 False

    curStep = "Step 5 - check the result"
    EngravedFloors dMin, dMax, nF, area
    Check (nF > 0 And Abs(dMin - NEW_DEPTH) < 0.0002 And Abs(dMax - NEW_DEPTH) < 0.0002), _
          "Every engraved face is .003 below the pocket floor (" & nF & " faces, " & Format(dMin, ".0000") & " to " & Format(dMax, ".0000") & ")"
    v1 = Volume()
    If Abs(oldD - NEW_DEPTH) < 0.00005 Then
        Check True, "The logo was already .003 deep - nothing changed"
    ElseIf v0 > 0 And v1 > 0 And area > 0 Then
        want = area * (oldD - NEW_DEPTH) * IN_M * 1000000000#          ' mm^3
        got = (v1 - v0) * 1000000000#
        Check (Abs(got - want) <= 0.03 * Abs(want) + 0.01), "Volume change " & Format(got, "0.00") & " mm^3 (expected " & _
              Format(want, "0.00") & " = engraved area x " & Format(oldD - NEW_DEPTH, ".0000") & " in)"
    Else
        Check False, "Volume could not be measured - check the depth in the feature tree"
    End If

    curStep = "Step 6 - save the part"
    Finish True
    Exit Sub
Fail:
    MsgBox "The macro stopped at: " & curStep & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Checks passed so far: " & nOK & vbCrLf & _
           IIf(nBad = 0, "", "Problems so far:" & vbCrLf & Left(problems, 500)), vbCritical, "HW#8A engrave macro"
End Sub

Function PartMatches(ByVal t As String) As Boolean
    t = UCase(Replace(Replace(t, "_", " "), "#", ""))
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    PartMatches = (InStr(t, "DRAWINGS II") > 0 And InStr(t, "PART 1") > 0) Or InStr(t, "IME COASTER") > 0
End Function

Sub FindPart()
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
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

' the dia 3.50 pocket gives the front direction; the back face is the far side of the body box
Sub FindPocket()
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, s As Object, cp As Variant, bx As Variant
    Dim pMid As Variant, bb As Variant, got As Boolean, lo As Double, hi As Double
    bodies = swPart.GetBodies2(0, True)
    If Not IsArray(bodies) Then Err.Raise vbObjectError + 3, , "The part has no solid body."
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsCylinder And Not got Then
                cp = s.CylinderParams
                If Abs(cp(6) / IN_M - PKT) < 0.004 Then
                    vF = Unit(Arr3(cp(3), cp(4), cp(5)))
                    bx = faces(i).GetBox
                    pMid = Arr3((bx(0) + bx(3)) / 2, (bx(1) + bx(4)) / 2, (bx(2) + bx(5)) / 2)
                    got = True
                End If
            End If
        Next
    Next
    If Not got Then Err.Raise vbObjectError + 4, , "This part has no dia 3.50 pocket, so it is not the In-Lab #8A part."
    bb = bodies(LBound(bodies)).GetBodyBox
    If Dot(Sub3(pMid, Arr3((bb(0) + bb(3)) / 2, (bb(1) + bb(4)) / 2, (bb(2) + bb(5)) / 2)), vF) < 0 Then vF = Scl(vF, -1)
    Span bb, lo, hi
    fBack = lo
    Check (Abs((hi - lo) / IN_M - 0.48) < 0.005), "Part is .480 thick (measured " & Format((hi - lo) / IN_M, "0.000") & ")"
End Sub

' smallest / largest of the body box corners along vF
Sub Span(bb As Variant, lo As Double, hi As Double)
    Dim ix As Integer, iy As Integer, iz As Integer, t As Double
    lo = 1E+30: hi = -1E+30
    For ix = 0 To 1
        For iy = 0 To 1
            For iz = 0 To 1
                t = Dot(Arr3(bb(3 * ix), bb(1 + 3 * iy), bb(2 + 3 * iz)), vF)
                If t < lo Then lo = t
                If t > hi Then hi = t
            Next
        Next
    Next
End Sub

' every extruded cut shallower than .02 in (the logo); returns how many, and the old depth
Function FindCuts(cuts() As Object, oldD As Double) As Long
    Dim f As Object, def As Object, d As Double, n As Long
    Set f = swPart.FirstFeature
    Do While Not f Is Nothing
        If f.GetTypeName2 = "Cut" Then
            Set def = Nothing
            d = -1
            On Error Resume Next
            Set def = f.GetDefinition
            d = def.GetDepth(True) / IN_M
            On Error GoTo 0
            If d > 0 And d < 0.02 And n < 20 Then
                Set cuts(n) = f
                n = n + 1
                oldD = d
                Check True, "Logo cut found: " & f.Name & ", " & Format(d, ".0000") & " deep" & IIf(f.IsSuppressed, " (suppressed - will be turned on)", "")
            End If
        End If
        Set f = f.GetNextFeature
    Loop
    FindCuts = n
End Function

Sub SetDepth(f As Object)
    Dim def As Object, ok As Boolean, d As Double, dd As Object, dm As Object, nm As String
    nm = f.Name
    On Error Resume Next
    If f.IsSuppressed Then
        swPart.ClearSelection2 True
        f.Select2 False, -1
        swPart.EditUnsuppress2
        swPart.ClearSelection2 True
    End If
    ' first way: the feature definition
    Set def = f.GetDefinition
    ok = def.AccessSelections(swPart, Nothing)
    def.SetDepth True, NEW_DEPTH * IN_M
    ok = f.ModifyDefinition(def, swPart, Nothing)
    If Not ok Then def.ReleaseSelectionAccess
    Err.Clear
    d = -1
    Set def = f.GetDefinition
    d = def.GetDepth(True) / IN_M
    ' second way: the feature's depth dimension
    If Abs(d - NEW_DEPTH) > 0.00005 Then
        Err.Clear
        Set dd = f.GetFirstDisplayDimension
        Do While Not dd Is Nothing
            Set dm = dd.GetDimension2(0)
            If dm.SystemValue / IN_M < 0.02 Then dm.SystemValue = NEW_DEPTH * IN_M
            Set dd = f.GetNextDisplayDimension(dd)
        Loop
        swPart.EditRebuild3
        d = -1
        Set def = f.GetDefinition
        d = def.GetDepth(True) / IN_M
    End If
    On Error GoTo 0
    Check (Abs(d - NEW_DEPTH) < 0.00005), nm & " depth set to .003" & _
          IIf(Abs(d - NEW_DEPTH) < 0.00005, "", " - it reads " & Format(d, ".0000") & "; double-click " & nm & " and type .003")
End Sub

' the flat faces just below the pocket floor (the bottoms of the engraved letters and symbol)
Sub EngravedFloors(dMin As Double, dMax As Double, nF As Long, area As Double)
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, s As Object, pp As Variant, n As Variant, f As Double
    dMin = 1#: dMax = 0#: nF = 0: area = 0#
    bodies = swPart.GetBodies2(0, True)
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsPlane Then
                pp = s.PlaneParams
                n = Unit(Arr3(pp(0), pp(1), pp(2)))
                If Abs(Dot(n, vF)) > 0.999 Then
                    f = (Dot(Arr3(pp(3), pp(4), pp(5)), vF) - fBack) / IN_M
                    If f > F_FLOOR - 0.02 And f < F_FLOOR - 0.0005 Then
                        nF = nF + 1
                        area = area + faces(i).GetArea
                        If F_FLOOR - f < dMin Then dMin = F_FLOOR - f
                        If F_FLOOR - f > dMax Then dMax = F_FLOOR - f
                    End If
                End If
            End If
        Next
    Next
End Sub

' part volume in m^3
Function Volume() As Double
    Dim mp As Object, v As Variant
    On Error Resume Next
    Set mp = swPart.Extension.CreateMassProperty
    If Not mp Is Nothing Then Volume = mp.Volume
    If Volume <= 0 Then
        v = swPart.GetMassProperties
        If IsArray(v) Then Volume = v(3)
    End If
End Function

Sub Finish(ByVal canSave As Boolean)
    Dim msg As String, ans As Long, ok As Boolean, errs As Long, warns As Long
    msg = "Checks passed: " & nOK & "      problems: " & nBad & vbCrLf & vbCrLf & _
          IIf(nBad = 0, "The logo is engraved .003 deep (note 3).", "Problems:" & vbCrLf & Left(problems, 700))
    If canSave Then
        ans = MsgBox(msg & vbCrLf & vbCrLf & "Save the part now?  (Yes = File > Save)", _
                     vbYesNo + IIf(nBad = 0, vbQuestion, vbExclamation), "HW#8A engrave")
        If ans = vbYes Then
            On Error Resume Next
            ok = swPart.Save3(1, errs, warns)                 ' 1 = silent
            On Error GoTo 0
            Check ok, "Part saved" & IIf(ok, "", " - use File > Save (error " & errs & ")")
        Else
            Check True, "Part not saved (your choice) - save it before you run the drawing macro"
        End If
    Else
        MsgBox msg, vbExclamation, "HW#8A engrave"
    End If
    WriteReport
    If canSave Then MsgBox IIf(nBad = 0, "Done. Now run HW8A_Drawing_Macro.bas.", "Done, with problems:" & vbCrLf & Left(problems, 600)) & _
                           vbCrLf & vbCrLf & "Checks: HW8A_engrave_checks.txt next to the part.", _
                           IIf(nBad = 0, vbInformation, vbExclamation), "HW#8A engrave"
End Sub

Sub WriteReport()
    Dim f As Integer, folder As String
    On Error Resume Next
    folder = Left(partPath, InStrRev(partPath, "\"))
    If Len(folder) = 0 Then folder = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    f = FreeFile
    Open folder & "HW8A_engrave_checks.txt" For Output As #f
    Print #f, "HW#8A engrave checks  (" & Now & ")"
    Print #f, "Passed: " & nOK & "   Problems: " & nBad
    Print #f, ""
    Print #f, report
    Close #f
End Sub

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

Function Scl(a As Variant, ByVal s As Double) As Variant
    Scl = Arr3(a(0) * s, a(1) * s, a(2) * s)
End Function

Function Sub3(a As Variant, b As Variant) As Variant
    Sub3 = Arr3(a(0) - b(0), a(1) - b(1), a(2) - b(2))
End Function
