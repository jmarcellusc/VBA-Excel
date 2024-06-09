Sub RemoveCarrageReturns()

' Removed carrage returns, substitutes it with a space

    Dim selectedRange As Range
    Dim cel As Range
    
    Set selectedRange = Application.Selection

    For Each cel In selectedRange.Cells
        cel.Value = Replace(cel.Value, Chr(10), " ")
    Next cel

End Sub

