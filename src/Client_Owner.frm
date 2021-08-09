VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Client_Owner 
   Caption         =   "Client Owner"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11220
   OleObjectBlob   =   "Client_Owner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Client_Owner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' This is used to retrieve user ID and Commit ID
Public Sub UserForm_initialize()
     Me.owner_ID = Environ("UserName")
     Me.owner_name.Value = Application.username
End Sub

' This is used for combox box selection
Public Sub company_name_Change()
   
    Dim client_rng As String
    
    With monthly_source_csa
        Application.ScreenUpdating = True
        Worksheets("monthly_source_csa").Activate
        Application.ScreenUpdating = False
        
        'This is used to get the last used row
        client_rng = getColRangeFunction("csa_clients")
        
        company_name.List = Sheets(9).Range(client_rng).Value
    End With
    
    ' This is used to make it searchable by within row.
    Dim n As Long
    Static found As Boolean '<--| this will be used as this sub "footprint" and avoid its recursive and useless calls

    If found Then '<-- we're here just after the text update made by the sub itself, so we must do nothing but erase our "footprint" and have this sub run at the next user combobox change
        found = False '<--| erase our "footprint"
        Exit Sub '<--| exit sub
    End If

    With Me.company_name '<--| reference userform combobox
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
    Dim company_str As String
    Dim lookupVal As String
    Dim i As Long

    
    With monthly_source_csa
        
        Application.ScreenUpdating = True
        Worksheets("monthly_source_csa").Activate
        Application.ScreenUpdating = False
        
        'This is used to get the last used row and the location of the column
        last_row = Range("A2").End(xlDown).Row
        company_str = getColStr("csa_clients")
    
        ' do not update screen when the code runs
        With Application
            .ScreenUpdating = Falses
            .EnableEvents = False
            .CutCopyMode = False
        End With
        
        ' This is used to loop in all rows in a range.
        For i = 2 To last_row
            lookupVal = Cells(i, company_str)
            ' Compare ComboBox with the range from the spreadsheet
            If lookupVal = company_name Then
                ' Copy and paste the entire row if the value matches in the TextBox-UserForum
                Cells(i, company_str).EntireRow.Copy
                Client_Owner.end_result.Paste
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
    company_name.Value = Null
    end_result.Value = ""
End Sub
