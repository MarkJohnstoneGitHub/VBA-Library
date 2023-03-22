Attribute VB_Name = "Test_ULong32_LeadingZeroCount"
'@Folder("Testing.VBACorLib.DataTypes.ULong32")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 22, 2023

Option Explicit

Private Sub TestingULong32_LeadingZeroCount()
    Dim result As ULong
    Dim val  As ULong
    
    val = ULong32.CreateChecked(0)
    result = ULong32.LeadingZeroCount(val)
    DisplayLeadingZeroCount val, result
    
    val = ULong32.CreateChecked(4294967295#)
    result = ULong32.LeadingZeroCount(val)
    DisplayLeadingZeroCount val, result
    
    val = ULong32.CreateChecked(1325)
    result = ULong32.LeadingZeroCount(val)
    DisplayLeadingZeroCount val, result

    val = ULong32.CreateChecked(4967295)
    result = ULong32.LeadingZeroCount(val)
    DisplayLeadingZeroCount val, result

    val = ULong32.CreateChecked(294967295)
    result = ULong32.LeadingZeroCount(val)
    DisplayLeadingZeroCount val, result

    val = ULong32.CreateChecked(394967295)
    result = ULong32.LeadingZeroCount(val)
    DisplayLeadingZeroCount val, result
End Sub

Private Sub DisplayLeadingZeroCount(ByRef val As ULong, ByRef result As ULong)
    Debug.Print "LeadingZeroCount(" & ULong32.ToString(val) & ")" & " = " & ULong32.ToString(result)
End Sub
