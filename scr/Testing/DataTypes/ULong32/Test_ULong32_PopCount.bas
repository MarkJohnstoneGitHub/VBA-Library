Attribute VB_Name = "Test_ULong32_PopCount"
'@Folder("Testing.VBACorLib.DataTypes.ULong32")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 23, 2023

Option Explicit

Private Sub TestingULong32_PopCount()
    Dim result As ULong
    Dim val  As ULong
    
    val = ULong32.CreateChecked(0)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
    
    val = ULong32.CreateChecked(4294967295#)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
    
    val = ULong32.CreateChecked(1325)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result

    val = ULong32.CreateChecked(4967295)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result

    val = ULong32.CreateChecked(294967295)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result

    val = ULong32.CreateChecked(394967295)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
End Sub

Private Sub DisplayPopCount(ByRef val As ULong, ByRef result As ULong)
    Debug.Print "PopCount(" & ULong32.ToString(val) & ")" & " = " & ULong32.ToString(result)
End Sub
