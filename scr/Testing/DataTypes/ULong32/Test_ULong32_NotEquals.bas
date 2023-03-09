Attribute VB_Name = "Test_ULong32_NotEquals"
'@Folder("Testing.VBACorLib.DataTypes.ULong32")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32NotEquals()
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H1F3&)
    DisplayNotEquals lhs, rhs, ULong32.NotEquals(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H1F3&)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayNotEquals lhs, rhs, ULong32.NotEquals(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayNotEquals lhs, rhs, ULong32.NotEquals(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0)
    rhs = ULong32.CreateTruncating(&HF6F2F1F0)
    DisplayNotEquals lhs, rhs, ULong32.NotEquals(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0)
    rhs = ULong32.CreateTruncating(&H0)
    DisplayNotEquals lhs, rhs, ULong32.NotEquals(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&HFFFFFFFF)
    rhs = ULong32.CreateTruncating(&HFFFFFFFF)
    DisplayNotEquals lhs, rhs, ULong32.NotEquals(lhs, rhs)
    
    lhs = ULong32.CreateTruncating(&H0&)
    rhs = ULong32.CreateTruncating(&HFFFFFFFF)
    DisplayNotEquals lhs, rhs, ULong32.NotEquals(lhs, rhs)
End Sub

Private Sub DisplayNotEquals(ByRef lhs As ULong, ByRef rhs As ULong, ByVal result As Boolean)
    Debug.Print ULong32.ToString(lhs) & " <> " & ULong32.ToString(rhs) & " : " & result
End Sub
