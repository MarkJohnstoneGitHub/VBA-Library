Attribute VB_Name = "Testing_UInt32Static_Min"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32StaticMin()
    Dim lhs  As ULong
    Dim rhs As ULong
    Dim result As ULong
    
    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&H1F3&)
    result = UInt32Static.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
    
    lhs = CBytesUInt32(&H1F3&)
    rhs = CBytesUInt32(&HF6F2F1F0)
    result = UInt32Static.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
    
    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    result = UInt32Static.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
    
    lhs = CBytesUInt32(&HF0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    result = UInt32Static.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
End Sub

Private Sub DisplayMin(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & ", " & UInt32Static.ToString(rhs) & " Min = " & UInt32Static.ToString(result)
End Sub
