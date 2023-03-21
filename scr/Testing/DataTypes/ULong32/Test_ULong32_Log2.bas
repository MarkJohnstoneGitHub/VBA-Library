Attribute VB_Name = "Test_ULong32_Log2"
'@Folder("Testing.VBACorLib.DataTypes.ULong32")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 22, 2023

Option Explicit

Private Sub TestingULong32_Log2()
    Dim result As ULong
    Dim val  As ULong
    
    val = ULong32.CreateChecked(1325)
    result = ULong32.Log2(val)
    DisplayLog2 val, result

    val = ULong32.CreateChecked(4294967295#)
    result = ULong32.Log2(val)
    DisplayLog2 val, result

    val = ULong32.CreateChecked(4967295)
    result = ULong32.Log2(val)
    DisplayLog2 val, result

    val = ULong32.CreateChecked(294967295)
    result = ULong32.Log2(val)
    DisplayLog2 val, result

    val = ULong32.CreateChecked(394967295)
    result = ULong32.Log2(val)
    DisplayLog2 val, result
End Sub

Private Sub DisplayLog2(ByRef val As ULong, ByRef result As ULong)
    Debug.Print "Log2(" & ULong32.ToString(val) & ")" & " = " & ULong32.ToString(result)
End Sub
