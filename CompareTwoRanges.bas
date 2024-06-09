Sub CompareTwoRanges()

' Compares two selected ranges (vertical columns), the small (initial) is the sample column and large (compare) is the reference

Dim initalRange As Range
Dim compareRange As Range


Set initalRange = Application.InputBox("Select the Smaller Range to Compare", "Initial Comparison", Type:=8)
Set compareRange = Application.InputBox("Select the Comparison Range", "Secondary Comparison", Type:=8)


For Each cell In initalRange
    If Application.WorksheetFunction.CountIf(compareRange, cell) = 0 Then
        cell.Interior.ColorIndex = 38
    End If
    
    Next cell


End Sub
