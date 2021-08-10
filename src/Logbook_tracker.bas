' This code work in Microsoft excel objects(Execute from this workbook sheet)
'This event code starts when the workbook is opened.
Private Sub Workbook_Open()

'Dimension variable and declare data type
Dim Lrow As Single
Dim log As Range
Dim i As Long

With Audit_log_book
    Application.ScreenUpdating = True
    Worksheets("Audit_log_book").Activate
    Application.ScreenUpdating = False
    
    'Save the row of the first empty cell in column A to variable Lrow.
    Lrow = Worksheets("Audit_log_book").Range("A" & Rows.Count).End(xlUp).Row + 1

    'Save text value "Open Workbook" to the first empty cell
    Worksheets("Audit_log_book").Range("A" & Lrow).Value = "Open workbook"

    'Save date and time to the corresponding cell in column B.
    Worksheets("Audit_log_book").Range("B" & Lrow).Value = Now
    
    
    'Set the range in column A you want to loop through
    Set log = Range("A2:A10000")
    For Each cell In log
        'test if cell is empty
        If cell.Value <> "" Then
            'write to adjacent cell
            cell.Offset(0, 2).Value = getUserName()
        End If
    Next
End With
End Sub




'This event code begins right before the user closes the workbook.
Private Sub Workbook_BeforeClose(Cancel As Boolean)

'Dimension variable and declare data type
Dim Lrow As Single
Dim log As Range
Dim i As Long



With Audit_log_book
    Application.ScreenUpdating = True
    Worksheets("Audit_log_book").Activate
    Application.ScreenUpdating = False

    'Save the row of the first empty cell in column A to variable Lrow.
    Lrow = Worksheets("Audit_log_book").Range("A" & Rows.Count).End(xlUp).Row + 1

    'Check if cell above equals text value "Close Workbook",
    'if true then withdraw value in Lrow with 1 and then save the result to Lrow
    If Worksheets("Audit_log_book").Range("A" & Lrow - 1).Value = "Close workbook" Then Lrow = Lrow - 1

    'Save text value "Close Workbook" to cell
    Worksheets("Audit_log_book").Range("A" & Lrow).Value = "Close workbook"

    'Save date and time to the corresponding cell in column B.
    Worksheets("Audit_log_book").Range("B" & Lrow).Value = Now
    
    'Set the range in column A you want to loop through
    Set log = Range("A2:A10000")
    For Each cell In log
        'test if cell is empty
        If cell.Value <> "" Then
            'write to adjacent cell
            cell.Offset(0, 2).Value = getUserName()
        End If
    Next
    
End With
End Sub
