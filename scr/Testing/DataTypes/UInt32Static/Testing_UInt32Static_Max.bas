Attribute VB_Name = "Testing_UInt32Static_Max"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32StaticMax()
    Dim lhs  As ULong
    Dim rhs As ULong
    Dim result As ULong
    
    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&H1F3&)
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
    
    lhs = CBytesUInt32(&H1F3&)
    rhs = CBytesUInt32(&HF6F2F1F0)
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
    
    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
    
    lhs = CBytesUInt32(&HF0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    result = UInt32Static.Max(lhs, rhs)
    DisplayMax lhs, rhs, result
End Sub

Private Sub DisplayMax(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & ", " & UInt32Static.ToString(rhs) & " Max = " & UInt32Static.ToString(result)
End Sub
