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

' This is used to retrieve user ID and Commit ID
Public Sub UserForm_initialize()
     Me.CSA_hostID.Value = Environ("UserName")
     Me.CSA_username.Value = Application.username
End Sub

' This the code for Office Code drop down lists
Public Sub Office_Code_Change()
    Dim cr_rng As String
    
    With office_codes
        Application.ScreenUpdating = True
        Worksheets("office_codes").Activate
        Application.ScreenUpdating = False
        
        'This is used to get the last used row
        cr_rng = getColRangeFunction("3CR")
    
        Office_Code.List = Sheets(10).Range(cr_rng).Value
    End With
    ' This is used to make it searchable by within row.
    Dim n As Long
    Static found As Boolean '<--| this will be used as this sub "footprint" and avoid its recursive and useless calls

    If found Then '<-- we're here just after the text update made by the sub itself, so we must do nothing but erase our "footprint" and have this sub run at the next user combobox change
        found = False '<--| erase our "footprint"
        Exit Sub '<--| exit sub
    End If

    With Me.Office_Code '<--| reference userform combobox
        If .Text = "" Then Exit Sub '<--| exit if no text has been typed in
        For n = 0 To .ListCount - 1 '<--|loop through its list
            If InStr(.List(n), .Text) > 0 Then '<--| if current list value contains typed in text...
                found = True '<--| leave our "footprint"
                .Text = .List(n) '<--| change text to the current list value. this will trigger this sub again but our "footprint" will make it exit
                Exit For '<--| exit loop
            End If
        Next n
    End With
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

