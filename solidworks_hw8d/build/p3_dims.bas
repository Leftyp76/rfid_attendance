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
