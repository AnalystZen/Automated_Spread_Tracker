Attribute VB_Name = "ModReports"
Sub CreateSearchReport(Optional ByVal WbNewReport As Workbook)
Attribute CreateSearchReport.VB_ProcData.VB_Invoke_Func = " \n14"
'// Dynamic report based on user search values in list box.
'// Created by "" on 3/2/2025.

    '// Handle errors.
    On Error GoTo ErrorHandler
    
    '// Turn on screen updating.
    Application.ScreenUpdating = False

    '// Declare variables.
    Dim LastRow As Long
    Dim RngTitleDisplay As Range
    Dim RngTableHeader As Range
    
    '// Assign values
    LastRow = Range("A" & Rows.Count).ROW
    Set RngTitleDisplay = Range("A1:L5")
    Set RngTableHeader = Range("A6:L6")

    '// Change color of header.
    WbNewReport.Activate
    Windows(ActiveWorkbook.Name).DisplayGridlines = False
    With RngTitleDisplay.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.149998474074526
        .PatternTintAndShade = 0
    End With
    
    '// Merge and change text.
    With RngTitleDisplay
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Value = "DENVER SPREADS"
        .Font.Name = "Bahnschrift SemiBold SemiConden"
        .Font.Size = 36
    End With
    
    '// Copy header range.
    ThisWorkbook.Activate
    ShHome.Range("B10:M10").Copy
   
    '// Change font of imported data in sheet.
    WbNewReport.Activate
    RngTableHeader.PasteSpecial (xlPasteValues)
    RngTableHeader.Font.Name = "Bahnschrift Light Condensed"
    RngTableHeader.Font.Size = 15
    Range("A1").Select
    
    '// Change row heights for report sheet.
    Range("A6:A10000").RowHeight = 25
    Range("A7:L10000").Font.Size = 12
    Range("A6:L10000").Font.Name = "Bahnschrift SemiBold SemiConden"
    Range("J7:J10000").HorizontalAlignment = xlRight
    Range("H7:H10000").HorizontalAlignment = xlRight
   
    '// Autofit and format columns.
    WbNewReport.Activate
    Columns("A:L").EntireColumn.AutoFit
    Columns("C:C").NumberFormat = "m/d/yyyy"
    Columns("D:D").NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
    Columns("H:H").NumberFormat = "0"
    Columns("L:L").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    
    '// Install formulas into sheet.
    Range("A999999").End(xlUp).Offset(1).Value = "Totals  :"
    Range("L999999").End(xlUp).Offset(1).FormulaR1C1 = WorksheetFunction.Sum(Range("L7", Range("L999999").End(xlUp)))
    Range("K999999").End(xlUp).Offset(1).FormulaR1C1 = WorksheetFunction.Sum(Range("K7", Range("K999999").End(xlUp)))
    
    '// Make total row font larger.
    Range("A999999").End(xlUp).Resize(1, 12).Font.Size = 11
    Range("A999999").End(xlUp).Resize(1, 12).Font.Bold = True
    
    '// Declare variables for row highlight.
    Dim MyRange As Range
    Dim RngRow As Long
    Dim RngCell As Range
    
    '// Assign values
    Set MyRange = Range("A7", Range("A999999").End(xlUp))
    RngRow = 1
    
    '// Start loop to highlight rows.
    For Each RngCell In MyRange
        If RngCell.ROW Mod 2 <> 0 Then
            Range("A" & 6 + RngRow, "L" & 6 + RngRow).Interior.Color = RGB(238, 239, 242)
        End If
            RngRow = RngRow + 1
    Next RngCell
    
    '// Make all columns fit on sheet.
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .CenterHeader = "Prepared by " & Application.UserName
        .RightHeader = "Page &P"
        .CenterFooter = "&T" & Chr(10) & "&D"
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.15)
    End With
    Application.PrintCommunication = True
    
    '// Activate sourcebook.
    ShHome.Activate
    
    '// Clean exit
    Exit Sub
    
ErrorHandler:

    '// user update on failure.
    MsgBox "Creation of report failed. Please try again"
    
    '// Turn on screen updating.
    Application.ScreenUpdating = True
    
End Sub
