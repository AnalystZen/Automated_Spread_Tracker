Attribute VB_Name = "ModImportData"
'// All subs related to importing of data are to be listed here.

Sub ChainAllImports(ByVal FirstDate As Date, Optional ByVal SecondDate As Date)
'// This procedure will import the main SAP retieval subs together. Created to stop GUI confirmations.
'// Created by "" on 2/26/2025.

    '// Declare variables.
    Dim UserName As String
    Dim ShData As Worksheet

    '// Assign values.
    UserName = Range("User")
    Set ShData = ShZ15
    
    '// Handle errors.
    On Error GoTo ErrHandler:
        
    '// This procedure will import the Z15 sap movement. It imports to the clipboard and pastes in the designated sheet.
    '// Created by "" on 01/17/2025.
    
    '// This establishes the SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
       
    '// Sap t code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "MB51"
    
     '// Clear sap bug that inputs random text into system.
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKUNNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBWART-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSOBKZ-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtEBELN-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtINSMK-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKDAUF-LOW").Text = ""
    session.findById("wnd[0]/usr/txtKDPOS-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKOSTL-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSAKTO-LOW").Text = ""
    session.findById("wnd[0]/usr/txtWEMPF-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtVGART-LOW").Text = ""
    session.findById("wnd[0]/usr/txtBKTXT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCPUDT-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXABLN-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    
    '// Selects the created variant in sap.
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LASSALAN"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").pressToolbarButton "&FIND"
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "DENVERZ15"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[2]").sendVKey 12
    '// CurrentCellRow property to selelct find function highlight when there is a layout problem.
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow
    session.findById("wnd[1]").sendVKey 2
    
    '// Input date.
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = FirstDate
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = SecondDate
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKUNNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSOBKZ-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtEBELN-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtINSMK-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKDAUF-LOW").Text = ""
    session.findById("wnd[0]/usr/txtKDPOS-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKOSTL-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSAKTO-LOW").Text = ""
    session.findById("wnd[0]/usr/txtWEMPF-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtVGART-LOW").Text = ""
    session.findById("wnd[0]/usr/txtBKTXT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCPUDT-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXABLN-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXBLNR-LOW").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[8]").press

    '// Begin export to clipboard.
    session.findById("wnd[0]").sendVKey 9
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]").sendVKey 0
    
    '// Got to home screen and re- select global layout.
'    session.findById("wnd[0]/tbar[0]/btn[3]").press
'    session.findById("wnd[0]/usr/radRHIER_L").Select
'    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
'    session.findById("wnd[0]").sendVKey 0
    session.StartTransaction "MB51"
    session.findById("wnd[0]/usr/radRHIER_L").SetFocus
    session.findById("wnd[0]/usr/radRHIER_L").Select
    session.findById("wnd[0]/tbar[0]/btn[3]").press


    '// Paste data in to worksheet.
    With ShData
        .Visible = xlSheetVisible
        .Activate
        .Cells.Clear
        .Range("A1").PasteSpecial
    End With
    
    '// Format Z15 with | delimiter. This installs Data into columns.
    Columns("A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
    
    '// Hide raw data sheet.
    ShData.Visible = xlSheetHidden
    
    '// Got to home page.
    ShHome.Activate

    '// This procedure will import the Z16 sap movement. It imports to the clipboard and pastes in the designated sheet.
    '// Created by "" on 01/17/2025.

    '// Declare variables.
    Dim ShDataTwo As Worksheet
    
    '// Assign values.
    UserName = Range("User")
    Set ShDataTwo = ShZ16
    
    '// Sap t code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "MB51"
    
    '// Clear sap bug that inputs random text into system.
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKUNNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBWART-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSOBKZ-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtEBELN-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtINSMK-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKDAUF-LOW").Text = ""
    session.findById("wnd[0]/usr/txtKDPOS-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKOSTL-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSAKTO-LOW").Text = ""
    session.findById("wnd[0]/usr/txtWEMPF-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtVGART-LOW").Text = ""
    session.findById("wnd[0]/usr/txtBKTXT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCPUDT-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXABLN-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    
    '// Selects the created variant in sap.
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LASSALAN"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").pressToolbarButton "&FIND"
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "DENVERZ16"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[2]").sendVKey 12
    
    '// CurrentCellRow property to selelct find function highlight when there is a layout problem.
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow
    session.findById("wnd[1]").sendVKey 2
    
    '// Input date.
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = FirstDate
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = SecondDate
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCHARG-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtLIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKUNNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSOBKZ-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtEBELN-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtINSMK-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKDAUF-LOW").Text = ""
    session.findById("wnd[0]/usr/txtKDPOS-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtKOSTL-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtSAKTO-LOW").Text = ""
    session.findById("wnd[0]/usr/txtWEMPF-LOW").Text = ""
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtVGART-LOW").Text = ""
    session.findById("wnd[0]/usr/txtBKTXT-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtCPUDT-LOW").Text = ""
    session.findById("wnd[0]/usr/txtMBLNR-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXABLN-LOW").Text = ""
    session.findById("wnd[0]/usr/txtXBLNR-LOW").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    '// Begin export to clipboard.
    session.findById("wnd[0]").sendVKey 9
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]").sendVKey 0
    
    '// Got to home screen and re- select global layout.
'    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
'    session.findById("wnd[0]").sendVKey 0
    session.StartTransaction "MB51"
    session.findById("wnd[0]/usr/radRHIER_L").SetFocus
    session.findById("wnd[0]/usr/radRHIER_L").Select
    session.findById("wnd[0]/tbar[0]/btn[3]").press

    '// Paste data in to worksheet.
    With ShDataTwo
        .Visible = xlSheetVisible
        .Activate
        .Cells.Clear
        .Range("A1").PasteSpecial
    End With
    
    '// Format Z16 with | delimiter. This installs Data into columns.
    Columns("A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
    
    '// Hide sheet.
    ShDataTwo.Visible = xlSheetHidden
    
    '// Got to home page.
    ShHome.Activate
    
    '// Exit sub if all good.
    Exit Sub
    
ErrHandler:
    '// Got to home page.
    ShHome.Activate

    '// Raise an error to stop main Parent sub.
    Err.Raise 1004
End Sub


Sub ImportZ15Data()
'// This procedure will import the Z15 sap movement. It imports to the clipboard and pastes in the designated sheet.
'// Created by "" on 01/17/2025.

    '// Declare variables.
    Dim FirstDate As Date
    Dim SecondDate As Date
    Dim UserName As String
    Dim ShData As Worksheet
    
    '// Assign values.
    FirstDate = Range("DateEntry")
    SecondDate = Range("SecondEntry")
    UserName = Range("User")
    Set ShData = ShZ15
    
    '// This establishes the SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
       
    '// Sap t code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "MB51"
    
    '// Selects the created variant in sap.
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LASSALAN"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").pressToolbarButton "&FIND"
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "DENVERZ15"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[2]").sendVKey 12
    '// CurrentCellRow property to selelct find function highlight when there is a layout problem.
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow
    session.findById("wnd[1]").sendVKey 2
    
    '// Clear sap bug that inputs random text into system.
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = FirstDate
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = SecondDate
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[8]").press
'    session.findById("wnd[0]/usr/lbl[1,1]").SetFocus
'    session.findById("wnd[0]").sendVKey 31
    
    '// Begin export to clipboard.
    session.findById("wnd[0]").sendVKey 9
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]").sendVKey 0
    
    '// Got to home screen.
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    '// Paste data in to worksheet.
    ShData.Activate
    ShData.Cells.ClearContents
    ShData.Range("A1").PasteSpecial
    
    '// Format Z15 with | delimiter. This installs Data into columns.
    Columns("A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
    
    '// Got to home page.
    ShHome.Activate

End Sub

Sub ImportZ16Data()
'// This procedure will import the Z16 sap movement. It imports to the clipboard and pastes in the designated sheet.
'// Created by "" on 01/17/2025.

    '// Declare variables.
    Dim FirstDate As Date
    Dim SecondDate As Date
    Dim UserName As String
    Dim ShDataTwo As Worksheet
    
    '// Assign values.
    FirstDate = Range("DateEntry")
    SecondDate = Range("SecondEntry")
    UserName = Range("User")
    Set ShDataTwo = ShZ16
    
    '// This establishes the SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
       
    '// Sap t code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "MB51"
    
    '// Selects the created variant in sap.
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LASSALAN"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").pressToolbarButton "&FIND"
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "DENVERZ16"
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    session.findById("wnd[2]").sendVKey 12
    
    '// CurrentCellRow property to selelct find function highlight when there is a layout problem.
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").CurrentCellRow
    session.findById("wnd[1]").sendVKey 2
    
    '// Clear sap bug that inputs random text into system.
    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = FirstDate
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = SecondDate
    session.findById("wnd[0]/usr/txtUSNAM-LOW").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[8]").press
'    session.findById("wnd[0]/usr/lbl[1,1]").SetFocus
'    session.findById("wnd[0]").sendVKey 31
    
    '// Begin export to clipboard.
    session.findById("wnd[0]").sendVKey 9
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]").sendVKey 0
    
    '// Got to home screen.
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    '// Paste data in to worksheet.
    ShDataTwo.Activate
    ShDataTwo.Cells.ClearContents
    ShDataTwo.Range("A1").PasteSpecial
    
    '// Format Z16 with | delimiter. This installs Data into columns.
    Columns("A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
    
    '// Got to home page.
    ShHome.Activate
    
End Sub
