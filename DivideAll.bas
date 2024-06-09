Public Sub DivideAll()

    'Varibles
    Dim cellRange As Range
    Dim selectedRange As Range
    Dim divValue As Variant
     
     ' Prompt user for input
     divValue = InputBox("Please enter a number to divide (Thats Not 0):", "Dividing Factor", 100)
     
     ' Validate if the input is a number
     While Not IsNumeric(divValue) And divValue <> 0
         ' Display an error message
         MsgBox "Invalid input. Please enter a valid number."
         
         ' Prompt user for input again
         divValue = InputBox("Please enter a number to divide:")
     Wend
    
    
    ' Process Selected
    Set selectedRange = Application.Selection
    
    For Each cell In selectedRange.Cells
        ' Process on Numeric Only
        Dim cellItem As Variant
        
        If IsNumeric(cell.Value) And divValue <> 0 Then
            ' Extract
            cellItem = cell.Value
            
            'Process Application
            cellItem = cellItem / divValue
            
            ' Insert back
            cell.Value = cellItem
        End If
        
    
    Next cell

End Sub
