Attribute VB_Name = "Get_Col_RangeFunction"
'Created By Jake Ayoub 6/30/2021

Public Function getColRangeFunction(colName As String) As String

    'create variables that will be used in this function
    Dim first As String
    Dim last As String
    Dim col As String
    Dim first_row As Integer
    Dim first_str As String
    Dim last_col As String
    Dim last_row As Integer
    Dim last_str As String
     
    'loop to check if colname is equal in range between columns A and X, easy to change below
    For Each i In Range("A1:X1")
        If i = colName Then
            'catches column, first and last rows
            col = Split(i.Address(1, 0), "$")(0)
            last_row = Range("A2").End(xlDown).Row
            first_row = 2
            
            'make first and last addresses as strings
            first_str = "" & col & first_row
            last_str = "" & first_col & last_row
            
            'function ouput in the next line is a combination of above two strings
            getColRangeFunction = "" & first_str & ":" & col & last_str
        End If
    Next

End Function

