VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Client_Finder 
   Caption         =   "Client Finder"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5325
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
    Office_Code.List = Sheets(4).Range("D2:D117").Value
End Sub

Private Sub result_Change()
Dim last_row As String
Dim last_col As String
Dim office_str As String

Dim lCol As Long
Dim lRow As Long

With data_test

    last_row = Range("A2").End(xlDown).Row
    last_col = Cells(1, Columns.Count).End(xlToLeft).Column

    For lRow = 2 To last_row
        'Check to see if A(lRow) = TextBox.  Exact match required
        If ws.Cells(lRow, "D").ComboBox = Office_Code.ComboBox Then
            MsgBox ("Match Found for Office Code #: " & Office_Code.ComboBox)
            For lCol = 1 To last_col      'Loop through columns A-Z, Copy lRow to Row 1
                Cells(1, lCol) = Cells(lRow, lCol)
            Next lCol
        Else
            MsgBox ("No match Found for Office Code #: " & Office_Code.ComboBox)
        End If
    Next lRow
End With
End Sub

'Public Sub Search_Bar_Click()
'
'Dim office_rng As String
'Dim item_rng As String
'
'With data_test
'
'    office_rng = getColRangeFunction("office_code")
'
'    For Each i In Range(office_rng)
'        If i = Office_Code_Change Then:
'
'        End If
'    Next
'
'End With
'End Sub




Public Sub Search_Bar_Click()
    
End Sub
