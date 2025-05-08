Attribute VB_Name = "ModAdavancedFilter"
Sub FilterTransactions()
'// Filter table for serach user form.
    
    '// Declare variables
    Dim LastRow As Long
    Dim RngCrit As Range
    
    '// Assign values
    LastRow = ShHome.Range("B999999").End(xlUp).ROW
    
    '// Exit if no data is present.
    If LastRow < 11 Then Exit Sub
    
    '// Start advanced filter. Filter criteria hidden on sheet starting at Range("BA1")
    On Error Resume Next
    With ShHome
        .Range("A10:M" & LastRow).AdvancedFilter xlFilterCopy _
        , CriteriaRange:=.Range("BC2:BC3") _
        , CopyToRange:=.Range("BF2") _
        , Unique:=True
    End With
    On Error GoTo 0
    
    '// Reset Slicer and remove filter arrows. Slicers errors with advance filter
    ActiveWorkbook.SlicerCaches("Slicer_MvT").RequireManualUpdate = False
    ActiveSheet.ListObjects("TblZ15").ShowAutoFilterDropDown = False
        
End Sub
