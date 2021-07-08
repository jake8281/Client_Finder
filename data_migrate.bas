'This function is used to organize data, populate vlookup referencing multiple spreadsheets.
'Created By Jake Ayoub 6/30/2021
'Updated 7/8/2021

Sub data_migrate()

    Dim contact_maincsa As String
    Dim contact_backupcsa As String
    Dim separator_rng As String
    Dim separator_str As String
    Dim ItemType_str As String
    
    Dim office_rng As String
    Dim contact_border As Range
    Dim silly_filter As String
    
    Dim client_maincsa As String
    Dim client_backupcsa As String
    
    
    With contacts
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("contacts").Activate
        Application.ScreenUpdating = False
        
        ' This is used to insert & populate the column with Vlookup referencing another worksheet
        Columns("B:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
        Range("B1").Value = "main_csa"
    
        contact_maincsa = getColRangeFunction("main_csa")
        Range(contact_maincsa).Formula = "=VLOOKUP(A2,'monthly_source_csa'!$A$1:$C$76,2,0)"
        Range(contact_maincsa).Select
        
        
        ' This is used to insert & populate the column with Vlookup referencing another worksheet
        Columns("C:C").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
        Range("C1").Value = "csa_backup"
    
        contact_backupcsa = getColRangeFunction("csa_backup")
        Range(contact_backupcsa).Formula = "=VLOOKUP(A2,'monthly_source_csa'!$A$1:$C$76,3,0)"
        Range(contact_backupcsa).Select
        
        ' This is used to insert & populate the column with Vlookup referencing another worksheet
        Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromRightOrBelow
        Range("D1").Value = "office_code"
    
        office_rng = getColRangeFunction("office_code")
        Range(office_rng).Formula = "=VLOOKUP(A2,'office_codes'!$A$1:$B$96,2,0)"
        Range(office_rng).Select
        
        ' This is used to insert separator column with Chr(10) for later use & hide column
        Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrBelow
        Range("I1").Value = "separator"
        
        separator_rng = getColRangeFunction("separator")
        With Range(separator_rng)
            .FormulaR1C1 = "" & Chr(10) & ""
        End With
        
        ' This is used to build colored borders around multiple ranges but it's not dynamic-needs improvement.
        Set contact_border = Range("A1:I229")
        With contact_border.Borders
            .LineStyle = xlContinuous
            .ColorIndex = 14 ' Navy border color
            .Weight = xlThin
        End With
        
        ' This is used to custom color fill the first row
        With Range("A1:I1")
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 49407
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        End With
        
        ' Add Filter to the worksheet  & Wrap text
        silly_filter = Range("A1:I1").AutoFilter
        Cells.EntireColumn.AutoFit
    End With
    
    With contacts
        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("contacts").Activate
        Application.ScreenUpdating = False
        
        'This is used to adjust column width
        ItemType_str = getColStr("item_type")
        With Columns(ItemType_str)
            .ColumnWidth = 56.43
        End With
        
        separator_str = getColStr("separator")
        With Columns(separator_str)
            .EntireColumn.Hidden = True
        End With
    End With
    
    With client_list

        ' This function let you run the code from any worksheet in the workbook
        Application.ScreenUpdating = True
        Worksheets("client_list").Activate
        Application.ScreenUpdating = False

        ' This is used to insert & populate the column with Vlookup referencing another worksheet
        client_maincsa = getColRangeFunction("csa_main")
        Range(client_maincsa).Formula = "=VLOOKUP(C2,'monthly_source_csa'!$A$1:$C$76,2,0)"
        Range(client_maincsa).Select

        ' This is used to insert & populate the column with Vlookup referencing another worksheet
        client_backupcsa = getColRangeFunction("csa_backup")
        Range(client_backupcsa).Formula = "=VLOOKUP(C2,'monthly_source_csa'!$A$1:$C$76,3,0)"
        Range(client_backupcsa).Select

        ' This is used to build colored borders around multiple ranges but it's not dynamic-needs improvement.
        Set contact_border = Range("A1:H229")
        With contact_border.Borders
            .LineStyle = xlContinuous
            .ColorIndex = 14 ' Navy border color
            .Weight = xlThin
        End With

        ' This is used to custom color fill the first row
        With Range("A1:H1")
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 49407
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With .Font
                .Color = -16776961
                .TintAndShade = 0
            End With
        End With

        ' Add Filter to the worksheet  & Wrap text
        silly_filter = Range("A1:H1").AutoFilter
        Cells.EntireColumn.AutoFit
    End With

End Sub


