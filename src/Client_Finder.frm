VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Client_Finder 
   Caption         =   "Client Finder"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690.001
   OleObjectBlob   =   "Client_Finder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Client_Finder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub UserForm_Initialize()
'     Me.CSA_username.Value = Environ("UserName")
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
    
    With contacts
        
        Application.ScreenUpdating = True
        Worksheets("contacts").Activate
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
            lookupVal = Cells(i, office_str)
            ' Compare ComboBox with the range from the spreadsheet
            If lookupVal = Office_Code Then
                ' Copy and paste the entire row if the value matches in the TextBox-UserForum
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

' This  function to clear the Combobox selection and textbox result.
Public Sub clear_result_Click()
    Office_Code.Value = Null
    result.Value = ""
End Sub

