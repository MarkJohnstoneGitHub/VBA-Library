Attribute VB_Name = "Test_ULong32_Min"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingULong32Min()
    Dim lhs  As ULong
    Dim rhs As ULong
    Dim result As ULong
    
    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H1F3&)
    result = ULong32.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&H1F3&)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    result = ULong32.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    result = ULong32.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HF0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    result = ULong32.Min(lhs, rhs)
    DisplayMin lhs, rhs, result
End Sub

Private Sub DisplayMin(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print ULong32.ToString(lhs) & ", " & ULong32.ToString(rhs) & " Min = " & ULong32.ToString(result)
End Sub
