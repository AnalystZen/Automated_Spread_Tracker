Attribute VB_Name = "ModDataCleaning"
'// All procedure that are related to data scrubbing paste here.

Sub UniqueMaterialDocuments()
Attribute UniqueMaterialDocuments.VB_ProcData.VB_Invoke_Func = " \n14"
'// This procedure will get the unique document numbers from the imported data from SAP.
'// Created by  on 2/24/2025.

    '// Declare variables
    Dim RngUnique As Range
    Dim RngDestination As Range
    
    '// Assign values
    Set RngUnique = Range("B4", Range("B1048575").End(xlUp))
    Set RngDestination = Range("U4")

    '// Dont continue if there is no data.
    If Range("B4").Value = "" Then Exit Sub

    '// Filter unique material documents for calculations
    RngUnique.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=RngDestination, Unique:=True
    
End Sub

Sub CleanNegatives()
'// This procedure will clean the negative values that are imported from SAP. This is needed as SAP format for negatives is "xx-".
'// Created by "" on 2/24/2205.

    '// Declare variables.
    Dim RngNegativeQty As Range
    Dim RngNegativeVal As Range
    Dim DataRowCount As Long
    Dim RngColumn As String
    Dim RngDataCleaned As Range
    
    '// Assign Values
    Set RngNegativeQty = Range("P4")
    Set RngNegativeVal = Range("Q4")
    Set RngDataCleaned = Range("L4")
    DataRowCount = Range("L4", Range("L1048576").End(xlUp)).Rows.Count + 3
    RngColumn = "Q"
    
    '// Dont run if there is no data.
    If RngDataCleaned = "" Then Exit Sub
    
    '// Scrub negative imported data by installing formulas.
    RngNegativeQty.FormulaR1C1 = "=TRIM(LEFT(RC[-4],LEN(RC[-4])-1))*-1"
    RngNegativeQty.AutoFill Destination:=Range(RngNegativeQty, RngNegativeVal), Type:=xlFillDefault
    
    '// Autofill formulas with dynamic row count.
    If RngDataCleaned.Offset(1).Value <> "" Then
        Range(RngNegativeQty, RngNegativeVal).AutoFill Destination:=Range(RngNegativeQty, RngColumn & DataRowCount), Type:=xlFillDefault
    End If

    '// Replace imported data that was imported from SAP with different format.
    Range(RngNegativeQty, RngColumn & DataRowCount).Copy
    RngDataCleaned.PasteSpecial xlPasteValues
    
    '// Remove formulas due to volume of data.
    Range(RngNegativeQty, RngColumn & DataRowCount).Clear
    Application.CutCopyMode = False
    
End Sub

Sub ArrRetrieveData()
'// This procedure will retrieve the unique material document data information. This was created to replace lookup formulas.
'// Created by "" on 2/24/2025.

    '// Declare variables.
    Dim MyArr() As Variant: MyArr = Range("B4", Range("K1048575").End(xlUp))
    Dim RngMaterialDocs As String: RngMaterialDocs = "U"
    Dim X As Long: X = 5
    Dim I As Long

    '// Dont run if there is no data.
    If Range("K4").Value = "" Then Exit Sub
    
    '// Start loop through array for data retrievel.
    For I = LBound(MyArr, 1) To UBound(MyArr, 1)
        If MyArr(I, 1) = Range(RngMaterialDocs & X) Then
            Range(RngMaterialDocs & X).Offset(, 1) = MyArr(I, 2)
            Range(RngMaterialDocs & X).Offset(, 2) = MyArr(I, 3)
            Range(RngMaterialDocs & X).Offset(, 3) = MyArr(I, 4)
            Range(RngMaterialDocs & X).Offset(, 4) = MyArr(I, 5)
            Range(RngMaterialDocs & X).Offset(, 5) = MyArr(I, 6)
            Range(RngMaterialDocs & X).Offset(, 6) = MyArr(I, 7)
            Range(RngMaterialDocs & X).Offset(, 7) = MyArr(I, 8)
            Range(RngMaterialDocs & X).Offset(, 8) = MyArr(I, 9)
            Range(RngMaterialDocs & X).Offset(, 9) = MyArr(I, 10)
            X = X + 1
        End If
    Next I
    
    '// Clear Arr from memory data. Not really needed.
    Erase MyArr

End Sub

Sub CalculateSums()
Attribute CalculateSums.VB_ProcData.VB_Invoke_Func = " \n14"
'// This procedure will install formulas for sums and lookups of imported data then remove fomulas.
'// Created by "" on 2/23/2025.

    '// Declare variables
    Dim RngRowCount As Long
    Dim RngFormula As Range
    Dim RngFormulaTwo As Range
    
    '// Assign values
    Set RngFormula = Range("AF5")
    Set RngFormulaTwo = Range("AE5")
    '// Get number of rows from filtered documents.
    RngRowCount = Range("U5", Range("U1048575").End(xlUp)).Rows.Count + 4
    
    '// Dont run if there is no data.
    If Range("H4").Value = "" Then Exit Sub
    
    '// Install formulas in to data sheet.
    RngFormula.FormulaR1C1 = "=IFERROR(SUMIF(C[-30],RC[-11],C[-19]),0)"
    RngFormulaTwo.FormulaR1C1 = "=SUMIF(C[-29],RC[-10],C[-19])"
    
    '// Autofill formula with dynamic range.
    If Range("U6").Value <> "" Then
        RngFormula.AutoFill Destination:=Range("AF5:AF" & RngRowCount)
        RngFormulaTwo.AutoFill Destination:=Range("AE5:AE" & RngRowCount)
    End If
    '// Remove formula cells by pasting values.
    ActiveSheet.Cells.Copy
    Range("A1").PasteSpecial xlValues
    Application.CutCopyMode = False
    
End Sub

Sub FilterData(Optional ByVal FilterValue)
'// filter and relocate data that was scrubbed based on a designated value amount.

    '// Declare variables
    Dim RngToFilter As Range
    Dim LocalCur As Range
    Dim AmountToFilterBy As Long
    Dim NegativeAmountToFilterby As Long
    Dim X As Long
    
    On Error Resume Next
    '// Assign Values
    Set RngToFilter = Range("U5").CurrentRegion.Resize(, 1).Offset(, 11)
    AmountToFilterBy = FilterValue
    '// Check for 0 in filter string value and adjust to 3000.
    If AmountToFilterBy = 0 Then AmountToFilterBy = 3000
    NegativeAmountToFilterby = AmountToFilterBy * -1
    X = 5
    On Error GoTo 0
    
    '// Dont run if there is no data
    If Range("U5").Value = "" Then Exit Sub
        
    '// Start Loop and set destination range.
    For Each LocalCur In RngToFilter
        If LocalCur.Value >= AmountToFilterBy Or LocalCur.Value <= NegativeAmountToFilterby Then
            Range("U" & LocalCur.ROW, Range("AF" & LocalCur.ROW)).Copy Range("AH" & X)
            X = X + 1
        End If
    Next LocalCur

End Sub

Sub ArrFilterToTbl()
'// This sub will filter the cleaned data via an array from the raw data sheets and filter it to an output array.
'// Created by "" on 2/25/2025.

    '// Declare variables.
    Dim ArrCleanedData() As Variant
    Dim TblArray() As Variant
    Dim ArrOutput() As Variant
    Dim DataExists As Boolean
    Dim CurrentRow As Long
    Dim I As Long
    Dim X As Long
    Dim J As Long
    
    '// Exit if there is no filtered data.
    If Range("AH5").Value = "" Then Exit Sub
    
    '// Assign values.
    ArrCleanedData = Range("AH5").CurrentRegion
    TblArray = ShHome.Range("B11").CurrentRegion.Resize(, 1)
    ReDim ArrOutput(1 To UBound(ArrCleanedData, 1), 1 To UBound(ArrCleanedData, 2))
    DataExists = False
    CurrentRow = 0
    
    '// Start Loop through arrays.
    For I = LBound(ArrCleanedData, 1) To UBound(ArrCleanedData, 1)
        For X = LBound(TblArray, 1) To UBound(TblArray, 1)
        
            '// Check to see if material document exists.
            If ArrCleanedData(I, 1) = TblArray(X, 1) Then
                DataExists = True
            End If
            
            '// If material document does not exist in table add data to output array.
            If DataExists = False And X = UBound(TblArray, 1) Then
                CurrentRow = CurrentRow + 1
                For J = 1 To UBound(ArrCleanedData, 2)
                    ArrOutput(CurrentRow, J) = ArrCleanedData(I, J)
                Next J
            End If
            
            '// If material document did exist then turn boolean to false to restart loop.
            If DataExists = True And X = UBound(TblArray, 1) Then
                DataExists = False
            End If
            
        Next X
    Next I
   
    '// Add unique material documents to table.
    If CurrentRow = 0 Then CurrentRow = 1
    ShHome.Range("B1048575").End(xlUp).Offset(1).Resize(CurrentRow, UBound(ArrOutput, 2)) = ArrOutput

End Sub

Sub SortTbl()
'// This procedure will sort the data table to the most recent date.
'// Created by "" on 2/25/2025.

    '// Activate Table.
    ShHome.Activate
    
    '// Clear and sort data table to newest entry.
    With ShHome.ListObjects("TblZ15").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("TblZ15[Pstng Date]"), SortOn:=xlSortOnValues, Order:=xlDescending
        .Header = xlYes
        .Apply
    End With
    
    '// Clear filter button if is showing.
    ShHome.ListObjects("TblZ15").ShowAutoFilterDropDown = False
    
End Sub

Sub Z15Chain(Optional ByVal FilterValue As String)
'// This procedure will chain all the data scrubbing subs together to for the Z15 Import.
'// Created by "" on 3/1/2025.

    '// Activate sheet to clean.
    With ShZ15
        .Visible = xlSheetVisible
        .Activate
    End With
     
    '// Get unique material list.
    If Range("F4").Value = "" Then
        MsgBox _
        Prompt:="There was no Z15 Spread data for the date selected." _
        , Buttons:=vbExclamation + vbOKOnly _
        , Title:="Input Different Date"
        
        '// Hide sheet if there is no data.
        ShZ15.Visible = xlSheetHidden
        ShHome.Activate
        Exit Sub
        Else
        UniqueMaterialDocuments
    End If
    
    '// Clean negative format from SAP.
    If Right(Range("L4"), 1) = "-" Then
        CleanNegatives
        Else
        '// Dont run
    End If
    
    '// Get material document information.
    ArrRetrieveData
    
    '// Calculate the sums.
    CalculateSums
    
    '// Filter values over user selected input.
    If Range("U5").Value = "" Then
        MsgBox _
        Prompt:="There was no data that matched the  Z15 value filter criteria." _
        , Buttons:=vbExclamation + vbOKOnly _
        , Title:="No Data At Spread Limit"
        
        '// Hide sheet if there is no data.
        ShZ15.Visible = xlSheetHidden
        ShHome.Activate
        Exit Sub
        Else
        FilterData FilterValue
    End If
    
    '// Add filtered data to table.
    If Range("AH5").Value = "" Then
        MsgBox _
        Prompt:="There was no data that matched the Z15 value filter criteria." _
        , Buttons:=vbExclamation + vbOKOnly, _
        Title:="No Data At Spread Limit"
        
        '// Hide sheet if there is no data.
        ShZ15.Visible = xlSheetHidden
        ShHome.Activate
        Exit Sub
        Else
        ArrFilterToTbl
    End If
    
    '// Sort table by newest date.
    SortTbl

    With ShZ15
        .Visible = xlSheetHidden
    End With
    
    '// Activate Table.
    ShHome.Activate

End Sub

Sub Z16Chain(Optional ByVal FilterValue As String)
'// This procedure will chain all the data scrubbing subs together to for the Z16 Import.
'// Created by "" on 3/1/2025.

     '// Activate sheet to clean.
    With ShZ16
        .Visible = xlSheetVisible
        .Activate
    End With
     
    '// Get unique material list.
    If Range("F4").Value = "" Then
        MsgBox _
        Prompt:="There was no Z16 spread data for the date selected." _
        , Buttons:=vbExclamation + vbOKOnly _
        , Title:="Input Different Date"
        
        '// Hide sheet if there is no data.
        ShZ16.Visible = xlSheetHidden
        ShHome.Activate
        Exit Sub
        Else
        UniqueMaterialDocuments
    End If
    
    '// Get material document information.
    ArrRetrieveData
    
    '// Calculate the sums.
    CalculateSums
    
    '// Filter vaules over user selected input.
    If Range("U5").Value = "" Then
        MsgBox _
        Prompt:="There was no data that matched the Z16 value filter criteria." _
        , Buttons:=vbExclamation + vbOKOnly _
        , Title:="No Data At Spread Limit"
        
        '// Hide sheet if there is no data.
        ShZ16.Visible = xlSheetHidden
        ShHome.Activate
        Exit Sub
    Else
        FilterData FilterValue
    End If
    
    '// Add filtered data to table.
    If Range("AH5").Value = "" Then
        MsgBox _
        Prompt:="There was no data that matched the Z16 value filter criteria." _
        , Buttons:=vbExclamation + vbOKOnly _
        , Title:="No Data At Spread Limit"
        
        '// Hide sheet if there is no data.
        ShZ16.Visible = xlSheetHidden
        ShHome.Activate
        Exit Sub
    Else
        ArrFilterToTbl
    End If
    
    '// Sort table by newest date.
    SortTbl

    '// Hide sheet after clean.
    With ShZ16
        .Visible = xlSheetHidden
    End With

    '// Activate Table.
    ShHome.Activate

End Sub

