' This code work in Microsoft excel objects(Execute from this workbook sheet)
Private Sub Workbook_Open()
  Dim wksUserLog As Worksheet
  Dim lngNextRow As Long
  
  On Error Resume Next
  Set wksUserLog = ThisWorkbook.Worksheets("User_Log")
  
  On Error GoTo ExitProc
  If wksUserLog Is Nothing Then
    Set wksUserLog = ThisWorkbook.Worksheets.Add
    wksUserLog.Name = "User Log"
    wksUserLog.Range("A1:B1").Value = Array("Username", "Datetime")
  End If
  
  With wksUserLog
    lngNextRow = .Cells(.Rows.Count, "A").End(xlUp).Row + 1
    .Cells(lngNextRow, "A").Value = Application.username
    .Cells(lngNextRow, "B").Value = Now()
    .Columns("A:B").AutoFit
  End With
  
ExitProc:
End Sub
