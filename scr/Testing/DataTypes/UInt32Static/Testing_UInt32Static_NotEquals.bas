Attribute VB_Name = "Testing_UInt32Static_NotEquals"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 7, 2023
'@LastModified March 7, 2023

Option Explicit

Private Sub TestingUInt32NotEquals()
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&H1F3&)
    DisplayNotEquals lhs, rhs, UInt32Static.NotEquals(lhs, rhs)
    
    lhs = CBytesUInt32(&H1F3&)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayNotEquals lhs, rhs, UInt32Static.NotEquals(lhs, rhs)
    
    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayNotEquals lhs, rhs, UInt32Static.NotEquals(lhs, rhs)
    
    lhs = CBytesUInt32(&H0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayNotEquals lhs, rhs, UInt32Static.NotEquals(lhs, rhs)
    
    lhs = CBytesUInt32(&H0)
    rhs = CBytesUInt32(&H0)
    DisplayNotEquals lhs, rhs, UInt32Static.NotEquals(lhs, rhs)
    
    lhs = CBytesUInt32(&HFFFFFFFF)
    rhs = CBytesUInt32(&HFFFFFFFF)
    DisplayNotEquals lhs, rhs, UInt32Static.NotEquals(lhs, rhs)
    
    lhs = CBytesUInt32(&H0&)
    rhs = CBytesUInt32(&HFFFFFFFF)
    DisplayNotEquals lhs, rhs, UInt32Static.NotEquals(lhs, rhs)
End Sub

Private Sub DisplayNotEquals(ByRef lhs As ULong, ByRef rhs As ULong, ByVal result As Boolean)
    Debug.Print UInt32Static.ToString(lhs) & " <> " & UInt32Static.ToString(rhs) & " : " & result
End Sub
