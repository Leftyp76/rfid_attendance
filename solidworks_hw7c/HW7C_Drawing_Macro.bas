' =====================================================================================
'  HW#7C  -  IME 1140 In-Lab Assignment #7C  -  SolidWorks DRAWING macro   (inches, sheet 1:2)
'
'  HOW TO RUN
'   1. Open "IME 1140 Engineering Drawings I In-Lab Assignment Part #3.SLDPRT" in SolidWorks.
'   2. Tools > Macro > New...  save it as  HW7C_Drawing.swp
'   3. In the VBA editor select ALL the text in the module, delete it, paste this whole file.
'   4. Run  main.  At the end enter your last name; it saves  HW#7C<LastName>.SLDDRW  and  .PDF
'      next to the part.
'
'  WHAT IT BUILDS (matches the reference sheet, scale 1:2)
'   - Front view (long side, R.50 groove on the left, 60 deg tab on the right), hidden lines visible
'   - Top view projected above it, hidden lines visible
'   - Auxiliary VIEW D-D square to the 60 deg end face of the tab
'   - Isometric view
'   - Front: 1.25, 1.00, R.50, 60 deg, .75, 1.50, .75     Top: 4.75, .75, dia .50, .375, 1.50
'   - VIEW D-D: 1.00, .625 THRU, .75, R.75, dia .875
'   - Title block: 1018 CRS, today's date, title, SHEET 3 OF 4, SCALE 1:2
'
'  v2: VIEW D-D tried several ways (direct edge, sheet pick, recover a view built but not returned,
'      last resort an ordinary view square to the face)
'
'  SAFETY
'   - Part orientation read from the dia .50 hole and the R.50 groove.
'   - Edges picked from their exact positions on the part; VIEW D-D edges found by measurement.
'   - Every view verified; every dimension read back and checked; views placed by real edges.
'   - The part file is never saved or changed by this macro.
' =====================================================================================
Option Explicit

Dim swApp As Object, swMath As Object, swPart As Object, swModel As Object, swDraw As Object
Dim partPath As String, curStep As String, report As String, problems As String
Dim nOK As Long, nBad As Long

Const IN_M = 0.0254
Const DRAW_NAME = "HW#7C"
Const PART_TAG = "PART #3"
Const SHEET_DEN = 2#             ' sheet scale 1 : 2

' ---- sheet layout (sheet inches, from the reference) ----
Const FRONT_X = 3.27             ' left end of the part (front view)
Const FRONT_Y = 2.14             ' bottom of the left block (front view)
Const TOP_Y = 6.34               ' front edge in the top view
Const AUX_X = 6.06               ' VIEW D-D centre
Const AUX_Y = 5.12
Const ISO_X = 9.21
Const ISO_Y = 6.94

Const DM_HIDDEN_VISIBLE = 1
Const DM_HIDDEN_REMOVED = 2

' ---- the part's directions (model coordinates, unit vectors) ----
Dim vR As Variant                ' along the length, groove end -> tab end   (right in the front view)
Dim vU As Variant                ' up                                         (up in the front view)
Dim vF As Variant                ' toward you in the front view (= the front face's outward normal)
Dim vN As Variant                ' outward normal of the 60 deg end face (the VIEW D-D direction)
Dim vA As Variant                ' along the end face, in the front-view plane
Dim pOrg As Variant              ' model point: left end, bottom of the left block, front face
Dim faceDeg As Double            ' measured angle of the tab end face from vertical (reference: 60)
Dim pickR As Double, pickF As Double   ' a point on the end face edge in the front view (part coords)

Dim viewFront As Object, viewTop As Object, viewAux As Object, viewIso As Object

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
    curStep = "Step 4 - new drawing (scale 1:2)":      NewDrawing tmpl
    curStep = "Step 5 - views and VIEW D-D":           MakeViews
    curStep = "Step 6 - centre marks and centrelines": CentreMarksAndLines
    curStep = "Step 7 - front view dimensions":        DimFront
    curStep = "Step 8 - top view dimensions":          DimTop
    curStep = "Step 9 - VIEW D-D dimensions":          DimAux
    curStep = "Step 10 - title block":                 TitleBlock
    curStep = "Step 11 - finish":                      FinishSheet

    MsgBox DRAW_NAME & " drawing built." & vbCrLf & vbCrLf & _
           "Checks passed: " & nOK & "      problems: " & nBad & vbCrLf & vbCrLf & _
           IIf(nBad = 0, "Every view, dimension and title block item checked out.", "Problems:" & vbCrLf & Left(problems, 700)) & _
           vbCrLf & vbCrLf & "The full check list is saved as HW7C_checks.txt next to the drawing." & vbCrLf & _
           "Next: enter your last name to save the drawing and the PDF.", _
           IIf(nBad = 0, vbInformation, vbExclamation), DRAW_NAME

    curStep = "Step 12 - save drawing and PDF":        SaveAll
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
    ' Use the open part whose name says Part #3. Otherwise the active part, otherwise the only part.
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
        "Could not tell which open part is In-Lab Assignment Part #3. Close the other parts, click into the Part #3 window, and run again."
    partPath = swPart.GetPathName
    If Len(partPath) = 0 Then Err.Raise vbObjectError + 2, , "The part has never been saved. Save it first."
    Check True, "Part found: " & Mid(partPath, InStrRev(partPath, "\") + 1)
End Sub

Sub FindOrientation()
    ' dia .50 hole (r .25): its axis is vertical.  R.50 groove (r .50): its axis runs front-to-back,
    ' and it sits above and to the right of the hole.
    Dim bodies As Variant, faces As Variant, b As Long, i As Long, k As Long
    Dim s As Object, cp As Variant, bx As Variant, ax As Variant, r As Double
    Dim hAx As Variant, hC(2) As Double, hN As Long, gAx As Variant, gC(2) As Double, gN As Long
    Dim bb As Variant, c As Variant, ix As Integer, iy As Integer, iz As Integer, t As Double, dd As Double
    Dim rMin As Double, rMax As Double, fMin As Double, aMin As Double, aMax As Double

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
                If Abs(r - 0.25) < 0.004 Then
                    hAx = ax: hN = hN + 1
                    For k = 0 To 2: hC(k) = hC(k) + (bx(k) + bx(k + 3)) / 2: Next
                ElseIf Abs(r - 0.5) < 0.004 Then
                    gAx = ax: gN = gN + 1
                    For k = 0 To 2: gC(k) = gC(k) + (bx(k) + bx(k + 3)) / 2: Next
                End If
            End If
        Next
    Next
    If hN = 0 Then Err.Raise vbObjectError + 4, , "This part has no dia .50 hole, so it is not In-Lab Assignment Part #3. Click into the Part #3 window and run again."
    If gN = 0 Then Err.Raise vbObjectError + 4, , "Could not find the R.50 groove in the part."
    For k = 0 To 2: hC(k) = hC(k) / hN: gC(k) = gC(k) / gN: Next

    vU = hAx                                                     ' groove is above the hole's middle
    If (gC(0) - hC(0)) * vU(0) + (gC(1) - hC(1)) * vU(1) + (gC(2) - hC(2)) * vU(2) < 0 Then vU = Scl(vU, -1)
    vR = Unit(Cross(vU, gAx))                                    ' groove is right of the hole
    If (gC(0) - hC(0)) * vR(0) + (gC(1) - hC(1)) * vR(1) + (gC(2) - hC(2)) * vR(2) < 0 Then vR = Scl(vR, -1)
    vF = Cross(vR, vU)
    vN = Add3(Scl(vR, 0.5), Scl(vU, Sqr(3#) / 2))               ' default 60 deg end face normal (measured below)

    bb = bodies(LBound(bodies)).GetBodyBox
    rMin = 1E+30: rMax = -1E+30: fMin = 1E+30: aMin = 1E+30: aMax = -1E+30
    For ix = 0 To 1
        For iy = 0 To 1
            For iz = 0 To 1
                c = Arr3(bb(3 * ix), bb(1 + 3 * iy), bb(2 + 3 * iz))
                t = Dot(c, vR): If t < rMin Then rMin = t
                If t > rMax Then rMax = t
                t = Dot(c, vU): If t < fMin Then fMin = t
                t = Dot(c, vF): If t < aMin Then aMin = t
                If t > aMax Then aMax = t
            Next
        Next
    Next
    ' origin: left end, bottom of the left block (.75 above the lowest point), front face
    pOrg = Add3(Add3(Scl(vR, rMin), Scl(vU, fMin + 0.75 * IN_M)), Scl(vF, aMax))
    Check (Abs((rMax - rMin) / IN_M - 4.75) < 0.01 And Abs((aMax - aMin) / IN_M - 1.5) < 0.01), _
          "Part size 4.75 long x 1.50 deep (measured " & Format((rMax - rMin) / IN_M, "0.000") & " x " & Format((aMax - aMin) / IN_M, "0.000") & ")"

    ' ---- the tab's end face, measured from the part: a flat face whose normal points up and right,
    '      passing through the top-right corner of the body.  Its angle drives VIEW D-D and the 60 deg check.
    Dim pp As Variant, n As Variant, nr As Double, nu As Double, nf As Double, cn As Variant, dist As Double, found As Boolean
    faceDeg = 60#
    For b = LBound(bodies) To UBound(bodies)
        faces = bodies(b).GetFaces
        For i = LBound(faces) To UBound(faces)
            Set s = faces(i).GetSurface
            If s.IsPlane And Not found Then
                pp = s.PlaneParams                                ' normal(3), point(3)
                n = Unit(Arr3(pp(0), pp(1), pp(2)))
                nr = Dot(n, vR): nu = Dot(n, vU): nf = Dot(n, vF)
                If nu < 0 Then nr = -nr: nu = -nu: nf = -nf
                If Abs(nf) < 0.01 And nr > 0.05 And nu > 0.05 Then
                    cn = PartPt(4.75, 0, 1)
                    dist = Abs((cn(0) - pp(3)) * n(0) + (cn(1) - pp(4)) * n(1) + (cn(2) - pp(5)) * n(2))
                    If dist < 0.0005 Then
                        found = True
                        faceDeg = Atn(nu / nr) * 180# / 3.14159265358979
                        vN = Unit(Add3(Scl(vR, nr), Scl(vU, nu)))
                    End If
                End If
            End If
        Next
    Next
    nr = Dot(vN, vR): nu = Dot(vN, vU)
    vA = Add3(Scl(vR, -nu), Scl(vU, nr))                         ' along the face, up and left
    pickR = 4.75 - 0.25 * nu: pickF = 1# + 0.25 * nr               ' 1/4 in up the face from the corner
    If found Then
        Check (Abs(faceDeg - 60#) < 0.5), "Tab end face measured at " & Format(faceDeg, "0.0") & ChrW(176) & _
              " from vertical through the top-right corner (reference 60" & ChrW(176) & "); VIEW D-D looks along its normal"
    Else
        Check False, "Tab end face not found in the part - assumed 60" & ChrW(176) & "; check VIEW D-D"
    End If
End Sub

' part coordinates (inches): r along the length from the left end, d back from the front face,
' f up from the bottom of the left block  ->  model point (metres)
Function PartPt(ByVal r As Double, ByVal d As Double, ByVal f As Double) As Variant
    PartPt = Add3(pOrg, Add3(Add3(Scl(vR, r * IN_M), Scl(vF, -d * IN_M)), Scl(vU, f * IN_M)))
End Function

Function SheetPt(v As Object, ByVal r As Double, ByVal d As Double, ByVal f As Double) As Variant
    Dim p As Variant
    p = PartPt(r, d, f)
    SheetPt = XfPt(v.ModelToViewTransform, p(0), p(1), p(2))
End Function

Function SX(v As Object, ByVal r As Double, ByVal d As Double, ByVal f As Double) As Double
    SX = SheetPt(v, r, d, f)(0)
End Function

Function SY(v As Object, ByVal r As Double, ByVal d As Double, ByVal f As Double) As Double
    SY = SheetPt(v, r, d, f)(1)
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
                 "then paste it here:", "HW#7C - drawing template")
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
        Check (Abs(p(2) - 1) < 0.001 And Abs(p(3) - SHEET_DEN) < 0.001), "Sheet scale 1:2"
        Check (p(4) = 0 Or p(4) = False), "Sheet projection: third angle"
    End If
End Sub

' =====================================================================================
'  Views and VIEW D-D
' =====================================================================================
Sub MakeViews()
    Dim dv As Variant, upI As Variant, rtI As Variant, n2 As Variant

    Set viewFront = OrientedView("HW7C_Front", vR, vU, vF, FRONT_X + 1.2, FRONT_Y + 0.3, StdFrontName(), False)
    If viewFront Is Nothing Then Err.Raise vbObjectError + 9, , "Could not create a correctly oriented front view."
    Check True, "Front view: groove on the left, 60 deg tab on the right"
    SetMode viewFront, DM_HIDDEN_VISIBLE, "Front"
    PlaceView viewFront, 0#, 0#, 0#, FRONT_X, FRONT_Y, True, True, "Front view placed (left end, bottom of the left block)"

    Set viewTop = Projected(viewFront, FRONT_X + 1.2, TOP_Y + 0.4)
    If viewTop Is Nothing Then Err.Raise vbObjectError + 10, , "Could not create the projected top view."
    Check FrameOK(viewTop, vR, 1, 0, vF, 0, -1), "Top view: front face toward the bottom of the sheet"
    SetMode viewTop, DM_HIDDEN_VISIBLE, "Top"
    PlaceView viewTop, 0#, 0#, 1#, 0#, TOP_Y, False, True, "Top view placed (front edge)"

    Set viewAux = MakeAux()
    If viewAux Is Nothing Then Err.Raise vbObjectError + 11, , "Could not create VIEW D-D."
    SetMode viewAux, DM_HIDDEN_REMOVED, "VIEW D-D"

    dv = Unit(Add3(Add3(vR, vU), vF))
    upI = Unit(Sub3(vU, Scl(dv, Dot(vU, dv))))
    rtI = Cross(upI, dv)
    Set viewIso = OrientedView("HW7C_Iso", rtI, upI, dv, ISO_X, ISO_Y, StdIsoName(), True)
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

' auxiliary view square to the 60 deg end face: select its edge in the front view, then create.
' Several ways are tried, because SolidWorks sometimes builds the view without handing it back.
Function MakeAux() As Object
    Dim av As Object, n2 As Variant, before As String, m As Integer, sp As Variant, how As String
    Dim ax As Variant, ay As Variant, w2 As Variant

    CollectEdges viewFront
    before = ViewNames()
    sp = SheetPt(viewFront, pickR, 0#, pickF)
    For m = 1 To 2
        swModel.ClearSelection2 True
        swDraw.ActivateView viewFront.GetName2
        If m = 1 Then
            SelPt viewFront, Array(pickR, 0#, pickF), False                    ' the edge object itself
        Else
            On Error Resume Next                                                 ' a click on the sheet
            swModel.Extension.SelectByID2 "", "EDGE", sp(0) * IN_M, sp(1) * IN_M, 0#, False, 0, Nothing, 0
            On Error GoTo 0
        End If
        If swModel.SelectionManager.GetSelectedObjectCount2(-1) > 0 Then
            On Error Resume Next
            Set av = Nothing
            Set av = swDraw.CreateAuxiliaryViewAt2(AUX_X * IN_M, AUX_Y * IN_M, 0#, False)
            If av Is Nothing Then Err.Clear: Set av = swDraw.CreateAuxiliaryViewAt(AUX_X * IN_M, AUX_Y * IN_M, 0#, False)
            On Error GoTo 0
            If av Is Nothing Then Set av = NewViewSince(before)                 ' built but not handed back
            If Not av Is Nothing Then how = IIf(m = 1, "edge selected directly", "edge picked on the sheet"): Exit For
        End If
    Next
    swModel.ClearSelection2 True

    If av Is Nothing Then
        ' last resort: an ordinary view turned to look square at the face (same picture, no D-D arrows)
        w2 = Cross(vN, vA)
        ax = Add3(Scl(vA, -Sqr(3#) / 2), Scl(w2, -0.5))                        ' face direction shown at 150 deg,
        ay = Add3(Scl(vA, 0.5), Scl(w2, -Sqr(3#) / 2))                          ' as in an aligned auxiliary view
        Set av = OrientedView("HW7C_DD", ax, ay, vN, AUX_X, AUX_Y, "", False)
        If av Is Nothing Then
            Check False, "VIEW D-D: SolidWorks would not create the auxiliary view"
            Exit Function
        End If
        Check False, "VIEW D-D made as a plain view square to the face (SolidWorks refused the auxiliary command). " & _
                     "Dimensions still go on it; the D-D arrows and label must be added by hand"
    Else
        Check True, "VIEW D-D created (" & how & ")"
    End If
    n2 = ViewDir(av, vN)                                     ' the face normal must point straight at you
    Check (Sqr(n2(0) ^ 2 + n2(1) ^ 2) < 0.1), "VIEW D-D looks square onto the 60 deg end face"
    Check True, "If the VIEW D-D arrows read A instead of D, click an arrow and change Label to D"
    Set MakeAux = av
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
    CollectEdges viewFront
    CentreMarkAt viewFront, Array(1.5, 0#, 0.567), "Front view: centre mark on the R.50 groove"
    CollectEdges viewTop
    CentreMarkAt viewTop, Array(0.375 + 0.1768, 0.75 + 0.1768, 1#), "Top view: centre mark on the " & ChrW(216) & ".50 hole"
    CollectEdges viewAux
    Dim h As Long
    h = AuxCircle(0.4375)
    If h >= 0 Then
        swModel.ClearSelection2 True
        swDraw.ActivateView viewAux.GetName2
        SelEdge viewAux, h, False
        Dim cm As Object
        On Error Resume Next
        Set cm = swDraw.InsertCenterMark3(0, False, True)
        On Error GoTo 0
        swModel.ClearSelection2 True
        Check Not cm Is Nothing, "VIEW D-D: centre mark on the " & ChrW(216) & ".875 hole"
    Else
        Check False, "VIEW D-D: " & ChrW(216) & ".875 hole not found for its centre mark"
    End If
    CentreLines viewFront
    CentreLines viewTop
    CentreLines viewAux
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

Sub DimFront()
    Dim xL As Double, xR As Double, yB As Double, yT As Double, yL As Double, xLB As Double
    CollectEdges viewFront
    xL = SX(viewFront, 0, 0, 0): xR = SX(viewFront, 4.75, 0, 0): xLB = SX(viewFront, 2.5, 0, 0)
    yB = SY(viewFront, 0, 0, 0): yT = SY(viewFront, 0, 0, 1): yL = SY(viewFront, 0, 0, -0.75)
    MakeDimP viewFront, "H", Array(0#, 0#, 0.3), Array(1.5, 0#, 0.567), xL - 0.36, yT + 0.48, 1.25, "Front   1.25", "", 2
    MakeDimP viewFront, "V", Array(1#, 0#, 0#), Array(0.35, 0#, 1#), xL - 0.25, (yB + yT) / 2, 1#, "Front   1.00", "", 2
    MakeDimP viewFront, "R", Array(1.5, 0#, 0.567), Empty, xL, yB - 0.46, 0.5, "Front   R.50", "", 2
    MakeDimP viewFront, "A", Array(4.75, 0#, 0.5), Array(pickR, 0#, pickF), xR - 0.59, yT + 0.95, 60#, "Front   60" & ChrW(176)
    MakeDimP viewFront, "V", Array(4.4, 0#, 0#), Array(3.9, 0#, -0.75), xR + 0.17, yB + 0.17, 0.75, "Front   .75  (step)", "", 2
    MakeDimP viewFront, "H", Array(2.5, 0#, -0.2), Array(4#, 0#, -0.2), xLB - 0.53, yL - 0.36, 1.5, "Front   1.50", "", 2
    MakeDimP viewFront, "H", Array(4#, 0#, -0.2), Array(4.75, 0#, 0.5), xR + 0.31, yL - 0.36, 0.75, "Front   .75  (right column)", "", 2
End Sub

Sub DimTop()
    Dim xL As Double, xR As Double, yBk As Double, yFr As Double
    CollectEdges viewTop
    xL = SX(viewTop, 0, 0, 1): xR = SX(viewTop, 4.75, 0, 1)
    yBk = SY(viewTop, 0, 1.5, 1): yFr = SY(viewTop, 0, 0, 1)
    MakeDimP viewTop, "H", Array(0#, 1.1, 1#), Array(4.75, 0.3, 1#), (xL + xR) / 2, yBk + 0.42, 4.75, "Top     4.75", "", 2
    MakeDimP viewTop, "V", Array(0.4, 1.5, 1#), Array(0.5518, 0.9268, 1#), xL - 0.26, yBk + 0.17, 0.75, "Top     .75", "", 2
    MakeDimP viewTop, "D", Array(0.5518, 0.9268, 1#), Empty, xL - 0.66, yFr - 0.28, 0.5, "Top     " & ChrW(216) & ".50", "", 2
    MakeDimP viewTop, "H", Array(0#, 1.1, 1#), Array(0.5518, 0.9268, 1#), xL + 0.7, yFr - 0.43, 0.375, "Top     .375", "", 3
    MakeDimP viewTop, "V", Array(0.4, 1.5, 1#), Array(0.4, 0#, 1#), xR + 0.4, (yBk + yFr) / 2, 1.5, "Top     1.50", "", 2
End Sub

' VIEW D-D: edges are found by measurement (their spacing), then dimensioned and checked
Sub DimAux()
    Dim sc As Double, a1 As Variant, a2 As Variant, hole As Long, big As Long, i As Long, j As Long
    Dim sh As Double, best As Long, bs As Double, s As Double, lineL As Long, xNear As Long, pLine As Long
    Dim hx As Double, hy As Double, mx As Double, my As Double, dx As Double, dy As Double, L As Double, sideSign As Double

    CollectEdges viewAux
    sc = 1#: On Error Resume Next: sc = viewAux.ScaleDecimal: On Error GoTo 0
    If sc <= 0 Then sc = 1#
    a1 = ViewDir(viewAux, vA)                     ' along the end face, on the sheet
    a2 = ViewDir(viewAux, vF)                     ' front-to-back, on the sheet
    hole = AuxCircle(0.4375)
    big = AuxCircle(0.75)
    If hole < 0 Then
        Check False, "VIEW D-D: " & ChrW(216) & ".875 hole not found - 1.00, .625 THRU, .75 and " & ChrW(216) & ".875 skipped"
        If big >= 0 Then MakeDimI viewAux, "R", big, -1, 0, 0, 0.75, "View D-D R.75", "", 2
        Exit Sub
    End If
    hx = eCX(hole): hy = eCY(hole)
    sh = (hx * a1(0) + hy * a1(1)) / sc

    ' .625 THRU: a line square to the face direction, .625 from the hole centre, on the right of the hole
    lineL = -1: bs = 1E+30
    For i = 0 To eN - 1
        If eKind(i) = "L" Then
            dx = eX2(i) - eX1(i): dy = eY2(i) - eY1(i): L = Sqr(dx * dx + dy * dy)
            If L > 0.02 Then
                If Abs((dx * a2(0) + dy * a2(1)) / L) > 0.99 Then
                    mx = (eX1(i) + eX2(i)) / 2: my = (eY1(i) + eY2(i)) / 2
                    s = (mx * a1(0) + my * a1(1)) / sc
                    If Abs(Abs(s - sh) - 0.625) < 0.01 And mx > hx Then
                        If Abs(Abs(s - sh) - 0.625) < bs Then bs = Abs(Abs(s - sh) - 0.625): lineL = i
                    End If
                End If
            End If
        End If
    Next
    If lineL >= 0 Then
        s = ((eX1(lineL) + eX2(lineL)) / 2 * a1(0) + (eY1(lineL) + eY2(lineL)) / 2 * a1(1)) / sc
        sideSign = Sgn(sh - s)
        ' 1.00: another edge (line square to the face direction, or a circle centre) 1.00 from that line, on the hole side
        xNear = -1: bs = 1E+30
        For j = 0 To eN - 1
            If j <> lineL And j <> hole Then
                If eKind(j) = "L" Then
                    dx = eX2(j) - eX1(j): dy = eY2(j) - eY1(j): L = Sqr(dx * dx + dy * dy)
                    If L > 0.02 Then
                        If Abs((dx * a2(0) + dy * a2(1)) / L) > 0.99 Then
                            mx = (eX1(j) + eX2(j)) / 2: my = (eY1(j) + eY2(j)) / 2
                            If Abs((mx * a1(0) + my * a1(1)) / sc - s - sideSign * 1#) < bs Then bs = Abs((mx * a1(0) + my * a1(1)) / sc - s - sideSign * 1#): xNear = j
                        End If
                    End If
                ElseIf eKind(j) = "C" Then
                    If Abs((eCX(j) * a1(0) + eCY(j) * a1(1)) / sc - s - sideSign * 1#) < bs Then bs = Abs((eCX(j) * a1(0) + eCY(j) * a1(1)) / sc - s - sideSign * 1#): xNear = j
                End If
            End If
        Next
        If bs > 0.01 Then xNear = -1
        MakeDimI viewAux, "S", xNear, lineL, hx + 1.48, hy + 0.79, 1#, "View D-D 1.00", "", 2
        MakeDimI viewAux, "S", hole, lineL, hx + 1.58, hy + 0.5, 0.625, "View D-D .625 THRU", "", 3, " THRU"
    Else
        Check False, "View D-D 1.00  - MISSING (no edge .625 from the hole found)"
        Check False, "View D-D .625 THRU  - MISSING (no edge .625 from the hole found)"
    End If

    ' .75: a long edge (along the face direction) .75 from the hole centre, the lower one
    pLine = -1: bs = 1E+30
    For i = 0 To eN - 1
        If eKind(i) = "L" Then
            dx = eX2(i) - eX1(i): dy = eY2(i) - eY1(i): L = Sqr(dx * dx + dy * dy)
            If L > 0.02 Then
                If Abs((dx * a1(0) + dy * a1(1)) / L) > 0.99 Then
                    mx = (eX1(i) + eX2(i)) / 2: my = (eY1(i) + eY2(i)) / 2
                    s = Abs(((mx - hx) * a2(0) + (my - hy) * a2(1)) / sc)
                    If Abs(s - 0.75) < 0.01 Then
                        If pLine < 0 Then
                            pLine = i
                        ElseIf my < (eY1(pLine) + eY2(pLine)) / 2 Then
                            pLine = i
                        End If
                    End If
                End If
            End If
        End If
    Next
    MakeDimI viewAux, "S", hole, pLine, hx + 1.37, hy, 0.75, "View D-D .75", "", 2
    MakeDimI viewAux, "R", big, -1, hx - 1.1, hy - 0.22, 0.75, "View D-D R.75", "", 2
    MakeDimI viewAux, "D", hole, -1, hx - 0.99, hy - 0.47, 0.875, "View D-D " & ChrW(216) & ".875", "", 3
End Sub

' face-on circle in the current edge table with model radius r (any arc or full circle)
Function AuxCircle(ByVal r As Double) As Long
    Dim i As Long, sc As Double
    AuxCircle = -1
    sc = 1#: On Error Resume Next: sc = viewAux.ScaleDecimal: On Error GoTo 0
    If sc <= 0 Then sc = 1#
    For i = 0 To eN - 1
        If eKind(i) = "C" Then
            If Abs(eR(i) / sc - r) < 0.01 Then AuxCircle = i: Exit Function
        End If
    Next
End Function

' kind: H horizontal, V vertical, S smart/aligned, A angle (expect in degrees), D diameter, R radius
Sub MakeDimP(v As Object, ByVal kind As String, pa As Variant, pb As Variant, ByVal tx As Double, ByVal ty As Double, _
             ByVal expect As Double, ByVal label As String, Optional ByVal prefix As String = "", Optional ByVal prec As Long = -1, _
             Optional ByVal suffix As String = "", Optional ByVal paren As Boolean = False)
    Dim need As Long
    If kind = "H" Or kind = "V" Or kind = "S" Or kind = "A" Then need = 2 Else need = 1
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    SelPt v, pa, False
    If need = 2 Then SelPt v, pb, True
    DoDim kind, need, tx, ty, expect, label, prefix, prec, suffix, paren
End Sub

' same, but with edges given by their index in the edge table (-1 = not found)
Sub MakeDimI(v As Object, ByVal kind As String, ByVal ia As Long, ByVal ib As Long, ByVal tx As Double, ByVal ty As Double, _
             ByVal expect As Double, ByVal label As String, Optional ByVal prefix As String = "", Optional ByVal prec As Long = -1, _
             Optional ByVal suffix As String = "", Optional ByVal paren As Boolean = False)
    Dim need As Long
    If kind = "H" Or kind = "V" Or kind = "S" Or kind = "A" Then need = 2 Else need = 1
    If ia < 0 Or (need = 2 And ib < 0) Then
        Check False, label & "  - MISSING (could not find the edge)"
        Exit Sub
    End If
    swModel.ClearSelection2 True
    swDraw.ActivateView v.GetName2
    SelEdge v, ia, False
    If need = 2 Then SelEdge v, ib, True
    DoDim kind, need, tx, ty, expect, label, prefix, prec, suffix, paren
End Sub

' creates the dimension from the current selection, formats it and checks its value
Sub DoDim(ByVal kind As String, ByVal need As Long, ByVal tx As Double, ByVal ty As Double, ByVal expect As Double, _
          ByVal label As String, ByVal prefix As String, ByVal prec As Long, ByVal suffix As String, ByVal paren As Boolean)
    Dim dd As Object, got As Double, nSel As Long, txt As String, ok As Boolean
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
        Case Else: Set dd = swModel.AddDimension2(tx * IN_M, ty * IN_M, 0#)
    End Select
    On Error GoTo 0
    swModel.ClearSelection2 True
    If dd Is Nothing Then Check False, label & "  - MISSING (SolidWorks did not create it)": Exit Sub

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
        If Left(txt, Len(prefix)) <> prefix Then prefix = prefix & txt Else prefix = txt
        dd.SetText 1, prefix
        txt = dd.GetText(1)
    End If
    If Len(suffix) > 0 Then dd.SetText 2, suffix
    On Error GoTo 0

    If kind = "A" Then
        got = got * 180# / 3.14159265358979              ' radians -> degrees
        ok = (Abs(got - expect) < 0.2)
    Else
        got = got / IN_M
        ok = (Abs(got - expect) < 0.003)
    End If
    If Not ok Then
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
                nt.SetText "SHEET 3 OF 4"
                If UCase(Trim(nt.GetText)) <> "SHEET 3 OF 4" Then nt.PropertyLinkedText = "SHEET 3 OF 4"
                sheetDone = (UCase(Trim(nt.GetText)) = "SHEET 3 OF 4")
            End If
            If Left(t, 6) = "SCALE:" And Not scaleDone Then              ' sheet is 1:2
                If InStr(t, "1:2") = 0 Then nt.SetText "SCALE: 1:2"
                If InStr(UCase(nt.GetText), "1:2") = 0 Then nt.PropertyLinkedText = "SCALE: 1:2"
                scaleDone = (InStr(UCase(nt.GetText), "1:2") > 0)
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

    AddSheetNote "1018 CRS", mX + 0.06, mY - 0.17, "Material: 1018 CRS"
    AddSheetNote Format(Date, "mm-dd-yyyy"), dX + 0.4, dY - 0.18, "Date: " & Format(Date, "mm-dd-yyyy")
    AddSheetNote "IME 1140 ENGINEERING DRAWINGS I", tX + 0.44, tY, "Title line 1"
    AddSheetNote "IN-LAB ASSIGNMENT #7C", tX + 0.44, tY - 0.17, "Title line 2"
    Check scaleDone, "SCALE: 1:2" & IIf(scaleDone, "", " - edit the scale note by hand")
    Check sheetDone, "SHEET 3 OF 4" & IIf(sheetDone, "", IIf(sheetSeen, " - the note would not change; edit it by hand", " - sheet note not found; edit it by hand"))
End Sub

Sub FinishSheet()
    ' leave the drawing active, in normal sheet mode, with all four views shown and redrawn
    Dim vs As Variant, n As Long, i As Long, v As Variant
    On Error Resume Next
    ActivateDrawing
    swDraw.EditSheet
    swModel.ClearSelection2 True
    For Each v In Array(viewFront, viewTop, viewAux, viewIso)
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
    If n >= 0 Then Check (n >= 4), n & " views on the sheet (expected 4)"
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
    lastName = Trim(InputBox("Last name for the file names  HW#7C<LastName>.SLDDRW  and  .PDF" & vbCrLf & _
                             "(leave empty to skip saving)", "Save HW#7C", "Perhach"))
    If Len(lastName) = 0 Then WriteReport SaveFolder() & "\": Exit Sub
    folder = Left(partPath, InStrRev(partPath, "\"))
    drw = folder & "HW#7C" & lastName & ".SLDDRW"
    ok = TrySave(drw)
    If Not ok Then
        folder = SaveFolder() & "\"
        drw = folder & "HW#7C" & lastName & ".SLDDRW"
        ok = TrySave(drw)
    End If
    If Not ok Then
        MsgBox "Could not save the drawing automatically. Use File > Save As and name it  HW#7C" & lastName & _
               ", then File > Save As > PDF with the same name.", vbExclamation, "HW#7C"
        Exit Sub
    End If
    pdf = folder & "HW#7C" & lastName & ".PDF"
    FinishSheet                                         ' views visible + drawing active before the PDF export
    okPdf = TrySavePdf(pdf)
    WriteReport folder
    MsgBox "Saved:" & vbCrLf & drw & vbCrLf & IIf(okPdf, pdf, "PDF NOT saved - use File > Save As > PDF"), _
           IIf(okPdf, vbInformation, vbExclamation), "HW#7C"
End Sub

Sub WriteReport(ByVal folder As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open folder & "HW7C_checks.txt" For Output As #f
    Print #f, "HW#7C drawing checks  (" & Now & ")"
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
