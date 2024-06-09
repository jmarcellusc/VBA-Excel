Sub sheetToCSV()

    Dim sFolder As String
    Dim myCSVFileName As String
    Dim tempWB As Workbook
    Dim sheetName As String
    Dim strFileName As String
    Dim splitFile
    
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
    sheetName = ActiveSheet.Name
    strFileName = ActiveWorkbook.Name
    splitFile = Split(strFileName, ".")
    strFileName = splitFile(0)
    ActiveSheet.Select
    ActiveSheet.Copy
    sheetName = ActiveSheet.Name
    Set tempWB = ActiveWorkbook
            
    myCSVFileName = sFolder & "\" & "CSV_ws" & sheetName & "__" & VBA.Format(VBA.Now, "dd-MMM") & ".csv"
    msCSVFileName2 = sFolder & "\" & strFileName & ".csv"
    MsgBox (msCSVFileName2)
    
    With tempWB
    .SaveAs Filename:=msCSVFileName2, FileFormat:=xlCSVUTF8, CreateBackup:=False
    .Close
    End With

    End If

End Sub
