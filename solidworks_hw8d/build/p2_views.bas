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
