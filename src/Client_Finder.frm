VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Client_Finder 
   Caption         =   "Client Finder"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16140
   OleObjectBlob   =   "Client_Finder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Client_Finder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is used to insert a logo
Private Sub logo_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Set picture = ThisWorkbook.Worksheets("logo").Shapes.Range(seperate_cellsay("Picture 2"))
    Sheets("logo").Visible = True
    With Sheets("logo").Shapes("logo").Fill
        .Visible = True
        .UserPicture picture
        .TextureTile = True
        .RotateWithObject = True
    End With
End Sub

' This is used to retrieve user ID and Commit ID
Public Sub UserForm_initialize()
     Me.CSA_hostID.Value = Environ("UserName")
     Me.CSA_username.Value = Application.username
End Sub

' This the code for Office Code drop down lists
Public Sub Office_Code_Change()
    Office_Code.List = Sheets(9).Range("B2:B73").Value
End Sub

Public Sub Search_Bar_Click()
    Dim last_row As String
    Dim last_col As String
    Dim office_str As String
    Dim lookupVal As String
    Dim i As Long
    Dim seperate_cells, cell_rng As Range
    Dim r As Range
    Dim row_str As String
    
    With contacts
        
        Application.ScreenUpdating = True
        Worksheets("contacts").Activate
        Application.ScreenUpdating = False
        
        'This is used to get the last used row and the location of the column
        last_row = Range("A2").End(xlDown).Row
        office_str = getColStr("office_code")
    
        ' do not update screen when the code runs
        With Application
            .ScreenUpdating = Falses
            .EnableEvents = False
            .CutCopyMode = False
        End With
        
        'This is used to loop in all rows in a range.
        For i = 2 To last_row
            lookupVal = Cells(i, office_str)
            ' Compare ComboBox with the range from the spreadsheet
            If lookupVal = Office_Code Then
                Set cell_rng = Rows(i & ":" & i).SpecialCells(xlCellTypeConstants)
                'Set a range which will return all cells value in the row, except the empty ones
                seperate_cells = cellsSeparator(cell_rng)
                'call a function able to make an seperate_cellsfrom the range set in the above line
                Client_Finder.result.Text = Client_Finder.result.Text & vbLf & Join(seperate_cells, vbLf) 'add the text obtained by joinning the seperate_cellsay to the next line of existing text
            End If
        Next i
        

        
         With Application
            .ScreenUpdating = True
            .EnableEvents = True
            .CutCopyMode = True
        End With
    End With
End Sub
' This  function to clear the Combobox selection and textbox result.
Public Sub clear_result_Click()
    Office_Code.Value = Null
    result.Value = ""
End Sub

