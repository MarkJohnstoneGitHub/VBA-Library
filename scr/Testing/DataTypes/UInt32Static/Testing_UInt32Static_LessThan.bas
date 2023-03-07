Attribute VB_Name = "Testing_UInt32Static_LessThan"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 7, 2023
'@LastModified March 7, 2023

Option Explicit

Private Sub TestingUInt32LessThan()
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&H1F3&)
    DisplayLessThan lhs, rhs, UInt32Static.LessThan(lhs, rhs)
    
    lhs = CBytesUInt32(&H1F3&)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayLessThan lhs, rhs, UInt32Static.LessThan(lhs, rhs)
    
    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayLessThan lhs, rhs, UInt32Static.LessThan(lhs, rhs)
    
    lhs = CBytesUInt32(&H0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayLessThan lhs, rhs, UInt32Static.LessThan(lhs, rhs)
    
    lhs = CBytesUInt32(&H0)
    rhs = CBytesUInt32(&H0)
    DisplayLessThan lhs, rhs, UInt32Static.LessThan(lhs, rhs)
    
    lhs = CBytesUInt32(&HFFFFFFFF)
    rhs = CBytesUInt32(&HFFFFFFFF)
    DisplayLessThan lhs, rhs, UInt32Static.LessThan(lhs, rhs)
    
    lhs = CBytesUInt32(&H0&)
    rhs = CBytesUInt32(&HFFFFFFFF)
    DisplayLessThan lhs, rhs, UInt32Static.LessThan(lhs, rhs)
End Sub

Private Sub DisplayLessThan(ByRef lhs As ULong, ByRef rhs As ULong, ByVal result As Boolean)
    Debug.Print UInt32Static.ToString(lhs) & " < " & UInt32Static.ToString(rhs) & " : " & result
End Sub
