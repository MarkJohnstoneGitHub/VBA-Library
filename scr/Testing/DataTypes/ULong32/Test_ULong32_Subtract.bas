Attribute VB_Name = "Test_ULong32_Subtract"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32Subtract()
    Dim result As ULong
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H1F3&)
    result = ULong32.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&HF6FFF0)
    result = ULong32.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result

    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H0&)
    result = ULong32.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&H0&)
    rhs = ULong32.CreateTruncating(&H0&)
    result = ULong32.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HFFFFFFFF)
    rhs = ULong32.CreateTruncating(&H1&)
    result = ULong32.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HFFFFFFFF)
    rhs = ULong32.CreateTruncating(&HF5FFEFF2)
    result = ULong32.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
End Sub

Private Sub DisplaySubtract(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print ULong32.ToString(lhs) & " - " & ULong32.ToString(rhs) & " = " & ULong32.ToString(result)
End Sub

Private Sub TestingULong32PerformanceSubtract()
    Dim result As ULong
    Dim dTime As Double

    Dim lhs  As ULong
    Dim rhs As ULong
    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H1F3)

    'Perform initial subtraction so overhead of initiliasing ULong32 isn't included in timer calculations
    result = ULong32.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        result = ULong32.Subtract(lhs, rhs)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Subtraction duration for 1,000,000 calculations : " & dTime
End Sub

