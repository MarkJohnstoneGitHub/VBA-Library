Attribute VB_Name = "Testing_UInt32Static_Compare"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 26, 2023
'@LastModified February 26, 2023

Option Explicit

Private Sub TestingUInt32Compare()
 Dim compareResult As Long

    Dim lhs  As ULong
    Dim rhs As ULong

    lhs.Value = &HF6F2F1F0
    rhs.Value = &H1F3&
    DisplayCompare lhs, rhs, UInt32Static.Compare(lhs, rhs)
    
    lhs.Value = &H1F3&
    rhs.Value = &HF6F2F1F0
    DisplayCompare lhs, rhs, UInt32Static.Compare(lhs, rhs)
    
    
    lhs.Value = &HF6F2F1F0
    rhs.Value = &HF6F2F1F0
    DisplayCompare lhs, rhs, UInt32Static.Compare(lhs, rhs)
    
    lhs.Value = &H0
    rhs.Value = &HF6F2F1F0
    DisplayCompare lhs, rhs, UInt32Static.Compare(lhs, rhs)
    
    lhs.Value = &H0
    rhs.Value = &H0
    DisplayCompare lhs, rhs, UInt32Static.Compare(lhs, rhs)
    
    lhs.Value = &HFFFFFFFF
    rhs.Value = &HFFFFFFFF
    DisplayCompare lhs, rhs, UInt32Static.Compare(lhs, rhs)
End Sub

Private Sub DisplayCompare(ByRef lhs As ULong, ByRef rhs As ULong, ByVal compareResult As Long)
    Select Case compareResult
        Case 0
            Debug.Print UInt32Static.ToString(lhs) & " = " & UInt32Static.ToString(rhs)
        Case 1
            Debug.Print UInt32Static.ToString(lhs) & " > " & UInt32Static.ToString(rhs)
        Case -1
            Debug.Print UInt32Static.ToString(lhs) & " < " & UInt32Static.ToString(rhs)
    End Select
    

End Sub


