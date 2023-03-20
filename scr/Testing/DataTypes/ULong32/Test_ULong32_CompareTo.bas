Attribute VB_Name = "Test_ULong32_CompareTo"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingUInt32CompareTo()
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H1F3&)
    DisplayCompareTo lhs, rhs, ULong32.CompareTo(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H1F3&)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayCompareTo lhs, rhs, ULong32.CompareTo(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayCompareTo lhs, rhs, ULong32.CompareTo(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayCompareTo lhs, rhs, ULong32.CompareTo(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0)
    rhs = ULong32.CreateTruncating(&H0)
    DisplayCompareTo lhs, rhs, ULong32.CompareTo(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&HFFFFFFFF)
    rhs = ULong32.CreateTruncating(&HFFFFFFFF)
    DisplayCompareTo lhs, rhs, ULong32.CompareTo(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0&)
    rhs = ULong32.CreateTruncating(&HFFFFFFFF)
    DisplayCompareTo lhs, rhs, ULong32.CompareTo(lhs, rhs)
End Sub

Private Sub DisplayCompareTo(ByRef lhs As ULong, ByRef rhs As ULong, ByVal compareResult As Long)
    Select Case compareResult
        Case 0
            Debug.Print ULong32.ToString(lhs) & " = " & ULong32.ToString(rhs)
        Case 1
            Debug.Print ULong32.ToString(lhs) & " > " & ULong32.ToString(rhs)
        Case -1
            Debug.Print ULong32.ToString(lhs) & " < " & ULong32.ToString(rhs)
    End Select
End Sub

