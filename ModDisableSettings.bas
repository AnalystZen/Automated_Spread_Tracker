Attribute VB_Name = "ModDisableSettings"
'// All subs related to app functionality list here.

Sub TurnOffApps()
' This sub will turn off app settings so macros can run faster.
' Created by  on 8/24/2024.
    With ThisWorkbook
        .Application.ScreenUpdating = False
        .Application.DisplayAlerts = False
        .Application.EnableEvents = False
        .ActiveSheet.DisplayPageBreaks = False
    End With
End Sub
Sub TurnOnApps()
' This sub will turn on app settings after macros have executed.
' Created by "" on 8/24/2024.
    With ThisWorkbook
        .Application.ScreenUpdating = True
        .Application.DisplayAlerts = True
        .Application.EnableEvents = True
    End With
End Sub


