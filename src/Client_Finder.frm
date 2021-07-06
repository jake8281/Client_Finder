VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Client_Finder 
   Caption         =   "Client Finder"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15135
   OleObjectBlob   =   "Client_Finder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Client_Finder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    
    With data_test
        
        Application.ScreenUpdating = True
        Worksheets("data_test").Activate
        Application.ScreenUpdating = False
        
        last_row = Range("A2").End(xlDown).Row
        last_col = Cells(1, Columns.Count).End(xlToLeft).Column
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
