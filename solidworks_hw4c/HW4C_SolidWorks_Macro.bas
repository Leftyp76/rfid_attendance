' =====================================================================================
'  HW#4C  -  Arch bracket  -  SolidWorks build macro   (MMGS)
'
'  HOW TO RUN
'   1. Open SolidWorks.  Tools > Macro > New...  save it as  HW4C_Build.swp
'   2. The VBA editor opens with an empty "Sub main()".  Select ALL the text in that
'      module, delete it, and paste the whole content of this file in its place.
'   3. Press F5 (Run).  A new part is built step by step; each feature, sketch and
'      reference plane is named "StepNN - ..." so the FeatureManager tree reads
'      like a recipe.  Roll back the tree (drag the rollback bar) to replay any step.
'   4. At the end the macro shows the model volume (target: 779,719 mm^3), then asks
'      for your last name and saves  HW#4C<LastName>.SLDPRT  on your Desktop.
'   5. Confirm the unit system shows "MMGS" in the status bar (bottom-right corner).
'
'  ORIENTATION (so the Isometric view matches the figure)
'   - Front Plane  = back face of the arch wall; the part grows toward the viewer (+Z)
'   - Top Plane    = bottom face of the base plate (the 6 mm pad hangs below it)
'   - Right Plane  = plane of symmetry (everything is mirrored about it)
'   - Origin       = centre of the back edge of the base-plate bottom face
'
'  v2: sketches are exited and selected before each extrude/cut.
'  Every sketch is drawn in MODEL coordinates (mm) and converted with the sketch's
'  own ModelToSketchTransform, so the geometry is exact whatever the plane orientation.
' =====================================================================================
Option Explicit

Dim swApp As Object          ' SldWorks.SldWorks
Dim swModel As Object        ' SldWorks.ModelDoc2
Dim swMath As Object         ' SldWorks.MathUtility
Dim swXform As Object        ' MathTransform: model -> active sketch
Dim PL_FRONT As String, PL_TOP As String, PL_RIGHT As String
Dim curStep As String
Dim curSketch As Object      ' Feature of the sketch that was just finished

' ---- SolidWorks API constants (numeric, so no type-library reference is required) ----
Const EC_BLIND = 0           ' swEndCondBlind
Const EC_THRU_ALL = 1        ' swEndCondThroughAll
Const EC_MIDPLANE = 6        ' swEndCondMidPlane
Const SC_SKETCHPLANE = 0     ' swStartSketchPlane
Const RP_DISTANCE = 8        ' swRefPlaneReferenceConstraint_Distance

' ---- Part dimensions (mm) --------------------------------------------------------------
Const BASE_W = 222#          ' base plate width  (X)
Const BASE_D = 96#           ' base plate depth  (Z)
Const BASE_T = 20#           ' base plate thickness -> base top face at Y = 20
Const PAD_W = 120#           ' pad under the base
Const PAD_T = 6#
Const HOLE22_Z = 52#         ' dia 22 hole: 44 mm from the FRONT edge  (96 - 44)
Const CY = 55#               ' arch centre height  (20 + 35)
Const WALL_T = 14#           ' arch wall thickness      Z  0..14
Const BAND_T = 14#           ' raised band + boss       Z 14..28
Const R_OUT = 59#            ' band outer radius
Const R_IN = 24#             ' band inner radius   (slot R48/R35 sits centred in the band)
Const LOBE_X = 41.5          ' centre of the rounded band ends  = (59+24)/2
Const LOBE_R = 17.5          ' radius of the rounded band ends  = (59-24)/2
Const SLOT_RO = 48#          ' through-slot outer radius
Const SLOT_RI = 35#          ' through-slot inner radius
Const SLOT_LR = 6.5          ' radius of the rounded slot ends  = (48-35)/2
Const BOSS_R = 16#           ' dia 32 boss
Const HOLE16_R = 8#          ' dia 16 through hole
Const GUS_XMID = 93#         ' gusset mid-plane   (inner face 88, outer face 98)
Const GUS_T = 10#
Const GUS_ZTOP = 40#         ' gusset slope reaches the top (Y=55) 56 mm behind the front edge
Const EAR_XMID = 117#        ' ear mid-plane      (inner face 111, outer face 123)
Const EAR_T = 12#
Const EAR_ZMIN = 64#         ' ear spans Z 64..96 (32 mm deep)
Const EAR_CZ = 80#           ' ear hole / R16 centre, Z
Const EAR_CY = -16#          ' ear hole / R16 centre, Y  (36 below the base top)
Const EAR_R = 16#
Const EAR_HOLE_R = 8#

' =====================================================================================
Sub main()
    On Error GoTo Fail
    Dim f As Object, TX As Double, TY As Double

    Set swApp = Application.SldWorks
    Set swMath = swApp.GetMathUtility

    curStep = "new part": NewPart

    ' ---------- Step 01 : base plate 222 x 96 x 20, sketched on the Top Plane ----------
    curStep = "Step01 base plate"
    SketchOnPlane PL_TOP
    MRect -BASE_W / 2, 0, 0, BASE_W / 2, 0, BASE_D
    EndSketchEntities
    Set f = Boss(BASE_T, False)
    NameFeat f, "Step01 - Base plate 222x96x20", "Sk01 - Base 222x96 (Top Plane)"
    EnsureExtent f, 1, 0, BASE_T                         ' must grow UP: Y 0..20

    ' ---------- Step 02 : pad 120 x 96 x 6 under the base ----------
    curStep = "Step02 pad"
    SketchOnFace 0, 0, BASE_D / 2                        ' bottom face of the base
    MRect -PAD_W / 2, 0, 0, PAD_W / 2, 0, BASE_D
    EndSketchEntities
    Set f = Boss(PAD_T, False)
    NameFeat f, "Step02 - Pad 120x96x6 under base", "Sk02 - Pad 120 wide (base bottom face)"
    EnsureExtent f, 1, -PAD_T, 0                         ' must grow DOWN: Y -6..0

    ' ---------- Step 03 : dia 22 THRU hole in the base, 44 from the front edge ----------
    curStep = "Step03 dia22 hole"
    SketchOnFace -105, BASE_T, 50                        ' top face of the base
    MCircle 0, BASE_T, HOLE22_Z, 11
    EndSketchEntities
    Set f = CutThruAll()
    NameFeat f, "Step03 - Hole dia22 THRU (base)", "Sk03 - dia22 at 44 from front"

    ' ---------- Step 04 : arch wall, 14 thick, sketched on the Front Plane ----------
    '  outline: 196 wide, vertical to 35 above the base, then straight lines tangent
    '  to the R59 arc (tangent points computed exactly)
    curStep = "Step04 arch wall"
    TX = R_OUT * (R_OUT / 98#)                            ' 35.52
    TY = CY + R_OUT * Sqr(1 - (R_OUT / 98#) ^ 2)          ' 102.11
    SketchOnPlane PL_FRONT
    MLine -98, BASE_T, 0, 98, BASE_T, 0
    MLine 98, BASE_T, 0, 98, CY, 0
    MLine 98, CY, 0, TX, TY, 0
    MArc3 TX, TY, 0, -TX, TY, 0, 0, CY + R_OUT, 0        ' R59 crown
    MLine -TX, TY, 0, -98, CY, 0
    MLine -98, CY, 0, -98, BASE_T, 0
    EndSketchEntities
    Set f = Boss(WALL_T, False)
    NameFeat f, "Step04 - Arch wall 14 thick (196 wide, R59)", "Sk04 - Arch outline R59 (Front Plane)"
    EnsureExtent f, 2, 0, WALL_T                         ' must grow toward viewer: Z 0..14

    ' ---------- Step 05 : raised horseshoe band R59/R24 + dia32 boss, 14 high ----------
    curStep = "Step05 band + boss"
    SketchOnFace 80, 30, WALL_T                          ' front face of the arch wall
    MArc3 -R_OUT, CY, WALL_T, R_OUT, CY, WALL_T, 0, CY + R_OUT, WALL_T              ' outer R59
    MArc3 R_OUT, CY, WALL_T, R_IN, CY, WALL_T, LOBE_X, CY - LOBE_R, WALL_T          ' right rounded end R17.5
    MArc3 R_IN, CY, WALL_T, -R_IN, CY, WALL_T, 0, CY + R_IN, WALL_T                 ' inner R24
    MArc3 -R_IN, CY, WALL_T, -R_OUT, CY, WALL_T, -LOBE_X, CY - LOBE_R, WALL_T       ' left rounded end R17.5
    MCircle 0, CY, WALL_T, BOSS_R                                                    ' dia 32 boss
    EndSketchEntities
    Set f = Boss(BAND_T, False)
    NameFeat f, "Step05 - Horseshoe band R59-R24 + boss dia32, 14 high", "Sk05 - Band + dia32 (wall front face)"
    EnsureExtent f, 2, WALL_T, WALL_T + BAND_T           ' Z 14..28

    ' ---------- Step 06 : curved slot R48/R35 with R6.5 ends, THRU ALL ----------
    curStep = "Step06 slot"
    SketchOnFace 0, 110, WALL_T + BAND_T                 ' front face of the band
    MArc3 -SLOT_RO, CY, 28, SLOT_RO, CY, 28, 0, CY + SLOT_RO, 28
    MArc3 SLOT_RO, CY, 28, SLOT_RI, CY, 28, LOBE_X, CY - SLOT_LR, 28
    MArc3 SLOT_RI, CY, 28, -SLOT_RI, CY, 28, 0, CY + SLOT_RI, 28
    MArc3 -SLOT_RI, CY, 28, -SLOT_RO, CY, 28, -LOBE_X, CY - SLOT_LR, 28
    EndSketchEntities
    Set f = CutThruAll()
    NameFeat f, "Step06 - Slot R48-R35 THRU", "Sk06 - Slot R48 R35 (band front face)"

    ' ---------- Step 07 : dia 16 THRU hole in the boss ----------
    curStep = "Step07 dia16 hole"
    SketchOnFace 10, CY, WALL_T + BAND_T                 ' front face of the boss
    MCircle 0, CY, WALL_T + BAND_T, HOLE16_R
    EndSketchEntities
    Set f = CutThruAll()
    NameFeat f, "Step07 - Hole dia16 THRU (boss)", "Sk07 - dia16 (boss front face)"

    ' ---------- Step 08/09/10 : gussets (10 thick) on a plane 93 from the Right Plane ----------
    curStep = "Step08 gusset plane"
    OffsetPlane PL_RIGHT, GUS_XMID, "Step08 - Gusset plane (93 from Right)"
    curStep = "Step09 gusset"
    SketchOnPlane "Step08 - Gusset plane (93 from Right)"
    '  profile in the side view: front-bottom corner -> up the slope -> flat top -> down the back
    MLine GUS_XMID, BASE_T, BASE_D, GUS_XMID, CY, GUS_ZTOP           ' slope: rises 35 over 56
    MLine GUS_XMID, CY, GUS_ZTOP, GUS_XMID, CY, 0                     ' flat top at 35 above base
    MLine GUS_XMID, CY, 0, GUS_XMID, BASE_T, 0                        ' back edge
    MLine GUS_XMID, BASE_T, 0, GUS_XMID, BASE_T, BASE_D               ' along the base top
    EndSketchEntities
    Set f = Boss(GUS_T, True)                                          ' mid-plane -> X 88..98
    NameFeat f, "Step09 - Gusset 10 thick (mid-plane)", "Sk09 - Gusset profile (gusset plane)"
    curStep = "Step10 mirror gusset"
    MirrorAboutRight "Step09 - Gusset 10 thick (mid-plane)", "Step10 - Mirror gusset (Right Plane)"

    ' ---------- Step 11/12/13 : ears (12 thick) on a plane 117 from the Right Plane ----------
    curStep = "Step11 ear plane"
    OffsetPlane PL_RIGHT, EAR_XMID, "Step11 - Ear plane (117 from Right)"
    curStep = "Step12 ear"
    SketchOnPlane "Step11 - Ear plane (117 from Right)"
    MLine EAR_XMID, BASE_T, BASE_D, EAR_XMID, BASE_T, EAR_ZMIN        ' top edge, flush with base top
    MLine EAR_XMID, BASE_T, EAR_ZMIN, EAR_XMID, EAR_CY, EAR_ZMIN      ' back edge, 32 behind the front
    MArc3 EAR_XMID, EAR_CY, EAR_ZMIN, EAR_XMID, EAR_CY, BASE_D, EAR_XMID, EAR_CY - EAR_R, EAR_CZ   ' R16 round
    MLine EAR_XMID, EAR_CY, BASE_D, EAR_XMID, BASE_T, BASE_D          ' front edge
    EndSketchEntities
    Set f = Boss(EAR_T, True)                                          ' mid-plane -> X 111..123
    NameFeat f, "Step12 - Ear 12 thick R16 (mid-plane)", "Sk12 - Ear profile R16 (ear plane)"
    curStep = "Step13 mirror ear"
    MirrorAboutRight "Step12 - Ear 12 thick R16 (mid-plane)", "Step13 - Mirror ear (Right Plane)"

    ' ---------- Step 14 : 2 x dia 16 THRU holes in the ears (one cut through both) ----------
    curStep = "Step14 ear holes"
    SketchOnPlane "Step11 - Ear plane (117 from Right)"
    MCircle EAR_XMID, EAR_CY, EAR_CZ, EAR_HOLE_R
    EndSketchEntities
    Set f = CutThruAll()
    NameFeat f, "Step14 - 2x hole dia16 THRU (ears)", "Sk14 - dia16 ear hole (ear plane)"

    ' ---------- finish ----------
    curStep = "finish"
    swModel.ClearSelection2 True
    swModel.ForceRebuild3 False
    swModel.ShowNamedView2 "*Isometric", 7
    swModel.ViewZoomtofit2

    Dim mp As Object, vol As Double
    Set mp = swModel.Extension.CreateMassProperty
    vol = mp.Volume * 1000000000#
    MsgBox "HW#4C built." & vbCrLf & vbCrLf & _
           "Volume = " & Format(vol, "#,##0") & " mm^3" & vbCrLf & _
           "Target  = 779,719 mm^3" & vbCrLf & vbCrLf & _
           "IMPORTANT: click the unit box in the bottom-right corner of the SolidWorks window" & vbCrLf & _
           "and choose MMGS before submitting.", vbInformation, "HW#4C"

    Dim lastName As String, path As String, errs As Long, warns As Long
    lastName = InputBox("Enter your last name to save as  HW#4C<LastName>.SLDPRT  on the Desktop" & vbCrLf & _
                        "(leave empty to skip saving)", "Save HW#4C")
    If Len(Trim(lastName)) > 0 Then
        path = Environ("USERPROFILE") & "\Desktop\HW#4C" & Trim(lastName) & ".SLDPRT"
        If swModel.Extension.SaveAs(path, 0, 1, Nothing, errs, warns) Then
            MsgBox "Saved: " & path, vbInformation
        Else
            MsgBox "Could not save automatically (error " & errs & "). Use File > Save As and name it HW#4C" & Trim(lastName), vbExclamation
        End If
    End If
    Exit Sub

Fail:
    MsgBox "The macro stopped at: " & curStep & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical, "HW#4C macro"
End Sub

' =====================================================================================
'  Document helpers
' =====================================================================================
Sub NewPart()
    Dim tmpl As String
    tmpl = swApp.GetUserPreferenceStringValue(8)           ' swDefaultTemplatePart
    Set swModel = Nothing
    If Len(tmpl) > 0 Then Set swModel = swApp.NewDocument(tmpl, 0, 0#, 0#)
    If swModel Is Nothing Then Set swModel = swApp.NewPart
    If swModel Is Nothing Then Err.Raise vbObjectError + 1, , "Could not create a new part. Check Tools > Options > Default Templates."
    On Error Resume Next
    swModel.Extension.SetUserPreferenceInteger 0, 0, 0     ' swUnitsLinear, no option, swMM
    On Error GoTo 0
    GetPlaneNames
End Sub

Sub GetPlaneNames()
    ' the first three RefPlane features of a new part are Front, Top and Right (any language)
    Dim ft As Object, n As Integer
    Set ft = swModel.FirstFeature
    Do While Not ft Is Nothing
        If ft.GetTypeName2 = "RefPlane" Then
            n = n + 1
            If n = 1 Then PL_FRONT = ft.Name
            If n = 2 Then PL_TOP = ft.Name
            If n = 3 Then PL_RIGHT = ft.Name: Exit Do
        End If
        Set ft = ft.GetNextFeature
    Loop
    If n < 3 Then Err.Raise vbObjectError + 2, , "Could not find the three default planes."
End Sub

Sub NameFeat(f As Object, ByVal featName As String, ByVal sketchName As String)
    Dim sk As Object
    f.Name = featName
    Set sk = f.GetFirstSubFeature
    If Not sk Is Nothing Then sk.Name = sketchName
End Sub

' =====================================================================================
'  Sketch helpers  (all coordinates are MODEL coordinates in mm)
' =====================================================================================
Sub SketchOnPlane(ByVal planeName As String)
    swModel.ClearSelection2 True
    If Not swModel.Extension.SelectByID2(planeName, "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0) Then
        Err.Raise vbObjectError + 5, , "Could not select plane '" & planeName & "'"
    End If
    StartSketch
End Sub

Sub SketchOnFace(ByVal x As Double, ByVal y As Double, ByVal z As Double)
    swModel.ClearSelection2 True
    If Not swModel.Extension.SelectByID2("", "FACE", x / 1000#, y / 1000#, z / 1000#, False, 0, Nothing, 0) Then
        Err.Raise vbObjectError + 6, , "Could not select the face at (" & x & ", " & y & ", " & z & ") mm"
    End If
    StartSketch
End Sub

Sub StartSketch()
    swModel.ViewZoomtofit2                                   ' sane zoom so inferencing behaves
    swModel.SketchManager.InsertSketch True
    swModel.SketchManager.AddToDB = False                    ' normal drawing: endpoints merge
    Set swXform = swModel.SketchManager.ActiveSketch.ModelToSketchTransform
End Sub

Sub EndSketchEntities()
    ' exit the sketch so it is solved, then select it for the feature that follows
    swModel.ClearSelection2 True
    swModel.SketchManager.InsertSketch True
    Set curSketch = swModel.FeatureByPositionReverse(0)      ' the sketch just finished
    swModel.ClearSelection2 True
    If Not swModel.Extension.SelectByID2(curSketch.Name, "SKETCH", 0#, 0#, 0#, False, 0, Nothing, 0) Then
        Err.Raise vbObjectError + 9, , "Could not select sketch '" & curSketch.Name & "'"
    End If
End Sub

Function SketchInfo() As String
    On Error Resume Next
    Dim sk As Object, n As Long
    Set sk = curSketch.GetSpecificFeature2
    n = sk.GetSketchSegmentCount
    SketchInfo = " [sketch '" & curSketch.Name & "' has " & n & " entities]"
End Function

Function SK(ByVal x As Double, ByVal y As Double, ByVal z As Double) As Variant
    ' model point (mm) -> sketch coordinates (m)
    Dim a(2) As Double, v As Variant, p As Object
    a(0) = x / 1000#: a(1) = y / 1000#: a(2) = z / 1000#
    v = a
    Set p = swMath.CreatePoint(v)
    Set p = p.MultiplyTransform(swXform)
    SK = p.ArrayData
End Function

Sub MLine(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
    Dim a As Variant, b As Variant
    a = SK(x1, y1, z1): b = SK(x2, y2, z2)
    swModel.SketchManager.CreateLine a(0), a(1), 0#, b(0), b(1), 0#
End Sub

Sub MArc3(ByVal xs As Double, ByVal ys As Double, ByVal zs As Double, _
          ByVal xe As Double, ByVal ye As Double, ByVal ze As Double, _
          ByVal xm As Double, ByVal ym As Double, ByVal zm As Double)
    ' arc from start (s) to end (e) passing through the mid point (m)
    Dim s As Variant, e As Variant, m As Variant
    s = SK(xs, ys, zs): e = SK(xe, ye, ze): m = SK(xm, ym, zm)
    swModel.SketchManager.Create3PointArc s(0), s(1), 0#, e(0), e(1), 0#, m(0), m(1), 0#
End Sub

Sub MCircle(ByVal xc As Double, ByVal yc As Double, ByVal zc As Double, ByVal r As Double)
    Dim c As Variant
    c = SK(xc, yc, zc)
    swModel.SketchManager.CreateCircleByRadius c(0), c(1), 0#, r / 1000#
End Sub

Sub MRect(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)
    Dim a As Variant, b As Variant
    a = SK(x1, y1, z1): b = SK(x2, y2, z2)
    swModel.SketchManager.CreateCornerRectangle a(0), a(1), 0#, b(0), b(1), 0#
End Sub

' =====================================================================================
'  Feature helpers
' =====================================================================================
Function Boss(ByVal depthMM As Double, ByVal midPlane As Boolean) As Object
    Dim t1 As Long, d As Double, f As Object
    If midPlane Then t1 = EC_MIDPLANE Else t1 = EC_BLIND
    d = depthMM / 1000#
    On Error Resume Next
    Set f = swModel.FeatureManager.FeatureExtrusion3(True, False, False, t1, EC_BLIND, d, d, _
            False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, SC_SKETCHPLANE, 0#, False)
    If f Is Nothing Then
        Err.Clear
        Set f = swModel.FeatureManager.FeatureExtrusion2(True, False, False, t1, EC_BLIND, d, d, _
                False, False, False, False, 0#, 0#, False, False, False, False, True, True, True)
    End If
    On Error GoTo 0
    If f Is Nothing Then Err.Raise vbObjectError + 3, , "Boss extrude failed (" & curStep & ")" & SketchInfo()
    Set Boss = f
End Function

Function CutThruAll() As Object
    ' cut "Through All" in BOTH directions, so it never depends on the sketch normal
    Dim f As Object
    On Error Resume Next
    Set f = swModel.FeatureManager.FeatureCut4(False, False, False, EC_THRU_ALL, EC_THRU_ALL, 0.01, 0.01, _
            False, False, False, False, 0#, 0#, False, False, False, False, False, True, True, True, True, False, _
            SC_SKETCHPLANE, 0#, False, False)
    If f Is Nothing Then
        Err.Clear
        Set f = swModel.FeatureManager.FeatureCut3(False, False, False, EC_THRU_ALL, EC_THRU_ALL, 0.01, 0.01, _
                False, False, False, False, 0#, 0#, False, False, False, False, False, True, True, True, True, False, _
                SC_SKETCHPLANE, 0#, False)
    End If
    On Error GoTo 0
    If f Is Nothing Then Err.Raise vbObjectError + 7, , "Cut failed (" & curStep & ")" & SketchInfo()
    Set CutThruAll = f
End Function

Sub OffsetPlane(ByVal baseName As String, ByVal distMM As Double, ByVal newName As String)
    Dim f As Object
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 baseName, "PLANE", 0#, 0#, 0#, False, 0, Nothing, 0
    Set f = swModel.FeatureManager.InsertRefPlane(RP_DISTANCE, distMM / 1000#, 0, 0#, 0, 0#)
    If f Is Nothing Then Err.Raise vbObjectError + 4, , "Could not create reference plane '" & newName & "'"
    f.Name = newName
    swModel.ClearSelection2 True
End Sub

Sub MirrorAboutRight(ByVal featName As String, ByVal newName As String)
    Dim f As Object
    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 PL_RIGHT, "PLANE", 0#, 0#, 0#, False, 2, Nothing, 0
    swModel.Extension.SelectByID2 featName, "BODYFEATURE", 0#, 0#, 0#, True, 1, Nothing, 0
    On Error Resume Next
    Set f = swModel.FeatureManager.InsertMirrorFeature2(False, False, False, False, 1)   ' 1 = swFeatureScope_AllBodies
    If f Is Nothing Then
        Err.Clear
        Set f = swModel.FeatureManager.InsertMirrorFeature(False, False, False, False)
    End If
    On Error GoTo 0
    If f Is Nothing Then Err.Raise vbObjectError + 8, , "Mirror failed for '" & featName & "'"
    f.Name = newName
    swModel.ClearSelection2 True
End Sub

' ---- direction guard: make sure an extrude went the intended way (axis 0=X 1=Y 2=Z) ----
Sub EnsureExtent(f As Object, ByVal axis As Integer, ByVal loMM As Double, ByVal hiMM As Double)
    Dim lo As Double, hi As Double, tries As Integer
    For tries = 1 To 2
        If Not FeatExtent(f, axis, lo, hi) Then Exit Sub
        If Abs(lo - loMM) < 0.5 And Abs(hi - hiMM) < 0.5 Then Exit Sub
        If tries = 1 Then ReverseExtrude f
    Next
    MsgBox "Please check the direction of '" & f.Name & "': it spans " & Format(lo, "0.0") & " .. " & _
           Format(hi, "0.0") & " mm (expected " & loMM & " .. " & hiMM & ").", vbExclamation
End Sub

Function FeatExtent(f As Object, ByVal axis As Integer, lo As Double, hi As Double) As Boolean
    Dim faces As Variant, i As Long, b As Variant
    lo = 1E+30: hi = -1E+30
    faces = f.GetFaces
    If IsEmpty(faces) Then Exit Function
    If Not IsArray(faces) Then Exit Function
    For i = LBound(faces) To UBound(faces)
        b = faces(i).GetBox
        If b(axis) * 1000# < lo Then lo = b(axis) * 1000#
        If b(axis + 3) * 1000# > hi Then hi = b(axis + 3) * 1000#
    Next
    FeatExtent = True
End Function

Sub ReverseExtrude(f As Object)
    Dim d As Object
    Set d = f.GetDefinition
    d.AccessSelections swModel, Nothing
    d.ReverseDirection = Not d.ReverseDirection
    f.ModifyDefinition d, swModel, Nothing
End Sub
