' NOTES block, top left: each number and each text is its own note so the columns line up exactly
Sub AddNotes()
    On Error GoTo Oops
    Dim lines As Variant, i As Long, y As Double, ok As Boolean
    lines = Array("USE THE IN-LAB ASSIGNMENT #8D PART MODEL ON CANVAS TO RE-CREATE THE ENGINEERING DRAWING", _
                  "MAKE SURE YOUR DRAWING'S TOLERANCES MATCH THE DRAWING TOLERANCES SHOWN BELOW.", _
                  "ENGRAVE THE MANUFACTURING ENGINEERING LOGO ON THE PART .003"" DEEP")
    ok = PlaceNote("NOTES:", NOTE_X, NOTE_Y)
    For i = 0 To 2
        y = NOTE_Y - (i + 1) * NOTE_DY
        ok = PlaceNote((i + 1) & ".", NOTE_X, y) And ok
        ok = PlaceNote(CStr(lines(i)), NOTE_X + NOTE_TAB, y) And ok
    Next
    Check ok, "NOTES 1-3 top left (each read back)"
    Exit Sub
Oops:
    Check False, curStep & " - stopped by error " & Err.Number & ": " & Err.Description & " (the rest of the drawing carries on)"
    On Error Resume Next
    swModel.ClearSelection2 True
End Sub
