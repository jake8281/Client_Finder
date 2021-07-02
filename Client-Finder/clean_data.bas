Attribute VB_Name = "clean_data"
'This function group data in one row if they sharing same value in specificed multiple ranges
'Created By Jake Ayoub 7/2/2021

Public Sub clean_data()

    Dim i As Long
    Dim last_row As String
    Dim company_str As String 'Col A
    Dim item_str As String  ' Col E
    Dim name_str As String ' Col F
    Dim email_str As String ' Col G
    Dim phone_str As String ' Col H
    
    With contacts
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("contacts").Activate
        Application.ScreenUpdating = False
         
        ' Big Nest to combine rows based on duplicate values in mulitple ranges
        'First Part: This is used to sort the filtered data based on the first col
        'Second Part:This is used to count the rows upwards from the bottom, and the--> Chr(10) to start from new line in the same cell
        
        last_row = Range("A2").End(xlDown).Row
        company_str = getColStr("Company")
        item_str = getColStr("item_type")
        name_str = getColStr("contact_name")
        email_str = getColStr("contact_email")
        phone_str = getColStr("contact_phone")
        
        'First Part
        With Range(Cells(2, company_str), Cells(last_row, phone_str))
            .Sort key1:=.Cells(1, 1), order1:=xlAscending, _
                  key2:=.Cells(1, 2), order2:=xlAscending, _
                  Header:=xlNo
        End With
        
        'Second Part
        For i = last_row - 1 To 2 Step -1
            If Cells(i, company_str).Value = Cells(i + 1, company_str).Value And _
               Cells(i, item_str).Value = Cells(i + 1, item_str).Value Then
                Cells(i, name_str).Value = Join(Array(Cells(i, name_str).Value, Cells(i + 1, 6).Value), Chr(10))
                Cells(i, email_str).Value = Join(Array(Cells(i, email_str).Value, Cells(i + 1, 7).Value), Chr(10))
                Cells(i, phone_str).Value = Join(Array(Cells(i, phone_str).Value, Cells(i + 1, 8).Value), Chr(10))
                Rows(i + 1).Delete
            End If
        Next i
    End With
End Sub



