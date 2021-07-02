Attribute VB_Name = "get_col_str_function"
'Created By Jake Ayoub 6/30/2021
'this function gives us the column letter as a string if you give it something to compare to

Public Function getColStr(colName As String) As String

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
            
            'function ouput in the next line is a combination of above two strings
            getColStr = "" & col
        End If
    Next
    
End Function
