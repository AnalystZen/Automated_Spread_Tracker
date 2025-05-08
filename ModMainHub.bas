Attribute VB_Name = "ModMainHub"
'// Action manager handles the structure off the app and most of error handling.

Sub ActionManager(Optional ByVal FirstDate As Date, Optional ByVal SecondDate As Date, Optional ByVal FilterValue As String)
'// Main mod for all the  call subs related to this project.
'// Created by "" on 2/27/2025.

    '// Handle Errors.
    On Error GoTo ErrorHandler
    
    '// Turn off screen updating.
    TurnOffApps
    
    '// First call to import data from SAP.
    ChainAllImports FirstDate, SecondDate
     
    '//Start the Z15 import and cleaning.
    Z15Chain FilterValue

    '//Start the Z16 import and cleaning.
    Z16Chain FilterValue

    '// Refresh Data.
    ThisWorkbook.RefreshAll

    '// Turn on screen updating.
    TurnOnApps
    
    '// Notification of failure
    MsgBox Prompt:="Data import ran succesfully.", Buttons:=vbInformation, Title:="Import Finished"
    
    '// Refresh data
    ThisWorkbook.RefreshAll
    
    '// Clean Exit.
    Exit Sub

ErrorHandler:

    '// Hide sheets on failure.
    ShZ15.Visible = xlSheetHidden
    ShZ16.Visible = xlSheetHidden

    '// Turn on screen updating.
    TurnOnApps
    
    '// Notification of failure
    MsgBox _
    Prompt:=Application.UserName & ", " & "Something went wrong." & vbNewLine & vbNewLine & "Please verify a session of SAP is open and a valid date was selected." _
    , Buttons:=vbCritical _
    , Title:="Import Failed"
    
End Sub
