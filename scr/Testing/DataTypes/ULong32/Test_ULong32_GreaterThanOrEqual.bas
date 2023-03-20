Attribute VB_Name = "Test_ULong32_GreaterThanOrEqual"
'@Folder("Testing.VBACorLib.DataTypes.ULong32")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingULong32GreaterThanOrEqual()
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H1F3&)
    DisplayGreaterThanOrEqual lhs, rhs, ULong32.GreaterThanOrEqual(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H1F3&)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayGreaterThanOrEqual lhs, rhs, ULong32.GreaterThanOrEqual(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayGreaterThanOrEqual lhs, rhs, ULong32.GreaterThanOrEqual(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayGreaterThanOrEqual lhs, rhs, ULong32.GreaterThanOrEqual(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0)
    rhs = ULong32.CreateTruncating(&H0)
    DisplayGreaterThanOrEqual lhs, rhs, ULong32.GreaterThanOrEqual(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&HFFFFFFFF)
    rhs = ULong32.CreateTruncating(&HFFFFFFFF)
    DisplayGreaterThanOrEqual lhs, rhs, ULong32.GreaterThanOrEqual(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0&)
    rhs = ULong32.CreateTruncating(&HFFFFFFFF)
    DisplayGreaterThanOrEqual lhs, rhs, ULong32.GreaterThanOrEqual(lhs, rhs)
End Sub

Private Sub DisplayGreaterThanOrEqual(ByRef lhs As ULong, ByRef rhs As ULong, ByVal result As Boolean)
    Debug.Print ULong32.ToString(lhs) & " >= " & ULong32.ToString(rhs) & " : " & result
End Sub

