Attribute VB_Name = "ModTesting"
'// Create and test any subs or functions here.

Sub RngTest()

    '// Test for table range. Dialing in offset range of table.
    Range("B1048575").End(xlUp).Offset(1).Resize(1, 12).Select

End Sub


