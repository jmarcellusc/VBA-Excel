Sub SpellCheckUpdate()

' Sourced Online (edited for each word)

Dim cel As Range, CellLen As Long, CurChr As Long, TheString As String

For Each cel In Selection
    For CurChr = 1 To Len(cel.Value)
        If Asc(mid(cel.Value, CurChr, 1)) = 32 Then
            If InStr(CurChr + 1, cel.Value, " ") = 0 Then
                TheString = mid(cel.Value, CurChr + 1, Len(cel.Value) - CurChr)
            Else
                TheString = mid(cel.Value, CurChr + 1, InStr(CurChr + 1, cel.Value, " ") - CurChr)
            End If
            If Not Application.CheckSpelling(word:=TheString) Then
                cel.Characters(CurChr + 1, Len(TheString)).Font.Color = RGB(255, 0, 0)
            Else
                cel.Characters(CurChr + 1, Len(TheString)).Font.Color = RGB(0, 0, 0)
            End If
            TheString = ""
        End If
    Next CurChr
Next cel

End Sub




