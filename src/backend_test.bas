Attribute VB_Name = "backend_test"
   Sub Test()
   
    Dim last_row As String
    Dim last_col As String
    Dim office_str As String
    
    Dim lookupVal As String
    Dim i As Long
    
    With data_test
        
        Application.ScreenUpdating = True
        Worksheets("data_test").Activate
        Application.ScreenUpdating = False
        
        'This is used to get the last used row and the location of the column
        last_row = Range("A2").End(xlDown).Row
        office_str = getColStr("office_code")
    
        ' do not update screen when the code runs
        With Application
            .ScreenUpdating = False
            .EnableEvents = False
            .CutCopyMode = False
        End With
        
        ' This is used to loop in all rows in a range.
        For i = 2 To last_row
'            lookupVal = Cells(i, "D")
            lookupVal = Cells(i, office_str)
            ' loop over values in "details"
            If lookupVal = Office_Code Then
'                Cells(i, "D").EntireRow.Copy
                Cells(i, office_str).EntireRow.Copy
                Client_Finder.result.Paste
            End If
        Next i
    
         With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .CutCopyMode = True
        End With
    End With
End Sub
