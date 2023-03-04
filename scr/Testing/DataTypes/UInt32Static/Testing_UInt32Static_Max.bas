Attribute VB_Name = "Testing_UInt32Static_Max"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.1 March 4, 2023
'@LastModified March 4, 2023

Option Explicit

Private Sub TestingUInt32StaticMax()
    Dim lhs  As ULong
    Dim rhs As ULong
    Dim result As ULong
    
    lhs.Value = &HF6F2F1F0
    rhs.Value = &H1F3&
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
    
    lhs.Value = &H1F3&
    rhs.Value = &HF6F2F1F0
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
    
    lhs.Value = &HF6F2F1F0
    rhs.Value = &HF6F2F1F0
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
    
    lhs.Value = &HF0
    rhs.Value = &HF6F2F1F0
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
End Sub

Private Sub DisplayMax(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & ", " & UInt32Static.ToString(rhs) & " Max = " & UInt32Static.ToString(result)
End Sub
