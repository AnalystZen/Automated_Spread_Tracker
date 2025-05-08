VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSearch 
   ClientHeight    =   10140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22485
   OleObjectBlob   =   "FrmSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// All subs related to the search form are listed here.

Private Sub UserForm_Initialize()
'// Load the combo and list box data when the form starts up. Data needed in list box row source to prevent errors.
'// Created by "" on 3/2/2025.

    '// Put useform in middle of sheet.
    Me.Top = Application.Height / 2 - (Me.Height / 2)
    Me.Left = Application.Width / 2 - (Me.Width / 2)

    '// Run Filter to populate values at start.
    ShHome.Range("BC2") = "Plnt"
    ShHome.Range("BC3") = "4014"

    '// Refilter based on input by user
    FilterTransactions

    '// Add table headers to combo box list from sheet.
    With Me.CmbBoxHeaders
        .RowSource = ShHome.Range("BA2", Range("BA13")).Address
    End With
    
    '// Source data from advanced filter range located on sheet.
    With Me.LstBoxSearchData
        .ColumnCount = 12
        .RowSource = ShHome.Range("BG3", Range("BR10000").End(xlUp)).Address
        .ColumnHeads = True
        .ListIndex = 0
        .Selected(0) = True
    End With
    
    '// Set selection focus to search box
    Me.TxtBoxSearch.SetFocus
    
End Sub

Private Sub CmbBoxHeaders_Change()
'// Change shhet value for advance filter criteria.
'// Created by "" on 3/2/2025.
    
    '// Change range for advanced filter.
    ShHome.Range("BC2").Value = Me.CmbBoxHeaders.Value

End Sub

Private Sub TxtBoxSearch_Change()
'// Filter string or numeric values based on user input. Change sheet value for advance filter.
'// Created by "" on 3/2/2025.
    
    '// Deaclare variables.
    Dim ColNum As Long
    
    '// Assign values. Start list index to match column count.
    ColNum = Me.CmbBoxHeaders.ListIndex + 1

    '// Filter for numerical values. Need parenethesis to evaluate if/or statement like below.
    If (ColNum = 1) Or (ColNum = 3) Or (ColNum = 4) Or (ColNum = 6) Or (ColNum = 7) Or (ColNum = 8) Or (ColNum = 10) Or (ColNum = 11) Or (ColNum = 12) Then
        ShHome.Range("BC3").Value = Me.TxtBoxSearch.Value
    Else
        '// Filter string values with wildcard.
        ShHome.Range("BC3").Value = "*" & Me.TxtBoxSearch.Value & "*"
    End If
    
        '// Refilter based on input by user
        FilterTransactions
        
End Sub

Private Sub CmdButtonClear_Click()
'// Clear search values that user input for search box.

    '// Clear values
    Me.TxtBoxSearch.Value = ""

End Sub

Private Sub CmdButtonReport_Click()
'// This will create a custom report based on the visible data in the list box.
'// Created by "" on 3/2/2025.

   '// Dont run if there is no active data.
   If Me.LstBoxSearchData.List(1, 1) = "" Then Exit Sub
      
    '// Declare variables
    Dim WbNewReport As Workbook
    Dim I As Long
    Dim X As Long
    
    '// Turn off Screen updating.
    TurnOffApps
    
    '// Assign values
    Set WbNewReport = Workbooks.Add
      
    '// Loop through  list box data and paste values.
    For I = 0 To Me.LstBoxSearchData.ListCount - 1
        For X = 1 To 12
            WbNewReport.Worksheets(1).Range("A" & I + 7).Offset(, X - 1) _
            = Me.LstBoxSearchData.List(I, X - 1)
        Next X
    Next I
    
    '// Format Report
    CreateSearchReport WbNewReport
    
    '// Remove copy ranges highlight.
     Application.CutCopyMode = False
     
    '// Turn on Screen updating.
    TurnOnApps

End Sub

Private Sub CmdButtonClose_Click()
'// If user selects  close, refill advance filter sourcerow before start.Sourcerow will error on blank data.
'// Created by "" on 3/1/2025.
    
    '// Unload form
    Unload Me
    
    '// Run Filter to populate values at start.
    ShHome.Range("BC2") = "Plnt"
    ShHome.Range("BC3") = "4014"

    '// Refilter based on input by user
    FilterTransactions
    
    '// Goto table.
    Application.GoTo ShHome.Range("B3"), True

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'// Happens first.If user selects  terminste button "X", refill advance filter sourcerow before start.Sourcerow will error on blank data.
'// Created by "" on 3/1/2025.

    If CloseMode = vbFormControlMenu Then
        CmdButtonClose_Click
    End If
    
End Sub

Private Sub UserForm_Terminate()
'// If user selects  terminste button "X", refill advance filter sourcerow before start.Sourcerow will error on blank data.
'// Created by "" on 3/1/2025.

    '// Run Filter to populate values at start.
    ShHome.Range("BC2") = "Plnt"
    ShHome.Range("BC3") = "4014"

    '// Refilter based on input by user
    FilterTransactions
    
    '// Goto table.
    Application.GoTo ShHome.Range("B3"), True

End Sub
