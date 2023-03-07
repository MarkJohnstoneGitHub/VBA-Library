Attribute VB_Name = "Testing_UInt32Static_CompareTo"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 7, 2023

Option Explicit

Private Sub TestingUInt32CompareTo()
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&H1F3&)
    DisplayCompareTo lhs, rhs, UInt32Static.CompareTo(lhs, rhs)
    
    lhs = CBytesUInt32(&H1F3&)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayCompareTo lhs, rhs, UInt32Static.CompareTo(lhs, rhs)
    
    
    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayCompareTo lhs, rhs, UInt32Static.CompareTo(lhs, rhs)
    
    lhs = CBytesUInt32(&H0)
    rhs = CBytesUInt32(&HF6F2F1F0)
    DisplayCompareTo lhs, rhs, UInt32Static.CompareTo(lhs, rhs)
    
    lhs = CBytesUInt32(&H0)
    rhs = CBytesUInt32(&H0)
    DisplayCompareTo lhs, rhs, UInt32Static.CompareTo(lhs, rhs)
    
    lhs = CBytesUInt32(&HFFFFFFFF)
    rhs = CBytesUInt32(&HFFFFFFFF)
    DisplayCompareTo lhs, rhs, UInt32Static.CompareTo(lhs, rhs)
    
    lhs = CBytesUInt32(&H0&)
    rhs = CBytesUInt32(&HFFFFFFFF)
    DisplayCompareTo lhs, rhs, UInt32Static.CompareTo(lhs, rhs)
End Sub

Private Sub DisplayCompareTo(ByRef lhs As ULong, ByRef rhs As ULong, ByVal compareResult As Long)
    Select Case compareResult
        Case 0
            Debug.Print UInt32Static.ToString(lhs) & " = " & UInt32Static.ToString(rhs)
        Case 1
            Debug.Print UInt32Static.ToString(lhs) & " > " & UInt32Static.ToString(rhs)
        Case -1
            Debug.Print UInt32Static.ToString(lhs) & " < " & UInt32Static.ToString(rhs)
    End Select
End Sub

