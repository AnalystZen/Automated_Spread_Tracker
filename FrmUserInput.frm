VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmUserInput 
   ClientHeight    =   8280.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6300
   OleObjectBlob   =   "FrmUserInput.frx":0000
End
Attribute VB_Name = "FrmUserInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// All subs related to the user input form are listed here.

Private Sub UserForm_Initialize()
'// User form for retrieval data.
'// Created by "" on 2/26/2025.


    '// Put useform in middle of sheet.
    Me.Top = Application.Height / 2 - (Me.Height / 2)
    Me.Left = Application.Width / 2 - (Me.Width / 2)

    '// Install txt into labels for form on intiatilization.
    Me.LblWelcomeUser.Caption = "Welcome" & " " & Application.UserName
    
    '// Install Label txt for search.
    Me.LblSearchOne.Caption = "Input Search Date :"
    Me.TxtSearchOne.Value = "MM/DD/YYYY"
    Me.TxtSearchOne.SetFocus
    
    '// Install Label txt for optional search.
    Me.LblSearchTwo.Caption = "(Optional) Search Date :"
    
    '// Install Filter label and Values
    Me.LblFilterValue = "Select Value Filter :"
    
    '// Insatll Combobox vales.
    Me.CmbFilterValues.List = Array("5000", "4000", "3000", "2000", "1000")
    Me.CmbFilterValues.ListIndex = 2
    
End Sub

Private Sub BtnCancel_Click()
'// When user hits cancel unload form.

    '// Begin to unload form from memory.
    Unload Me

End Sub

Private Sub BtnImportData_Click()
'// When the user clicks the import button run main sub.
    
    On Error Resume Next
    '// Declare variables
    Dim FirstDate As Date
    Dim SecondDate As Date
    Dim FilterValue As String
    
    '// Assign Values.
    FirstDate = CDate(TxtSearchOne)
    SecondDate = CDate(TxtSearchTwo)
    FilterValue = CmbFilterValues.Value
    On Error GoTo 0
    
    '// Check if dates are valid to run.
    If FirstDate = #12:00:00 AM# Then
    MsgBox _
        Prompt:="Hello, " & Application.UserName & "." & vbNewLine & vbNewLine _
        & "Please insert a valid date and try again." _
        , Buttons:=vbExclamation + vbOKOnly, _
        Title:="Insert A Valid Date"
        Exit Sub
        ElseIf SecondDate = #12:00:00 AM# Then
        SecondDate = FirstDate
    End If

    '// Begin to unload form from memory.
    Unload Me

    '// Run sub to get data.
    ActionManager FirstDate, SecondDate, FilterValue
    
End Sub

