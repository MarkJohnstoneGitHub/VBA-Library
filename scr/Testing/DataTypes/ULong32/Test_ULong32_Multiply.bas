Attribute VB_Name = "Test_ULong32_Multiply"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingULong32Multiply()
    Dim result As ULong
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = ULong32.CreateTruncating(&HF6F2F1F)
    rhs = ULong32.CreateTruncating(&HF&)
    result = ULong32.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HF62)
    rhs = ULong32.CreateTruncating(&HF6)
    result = ULong32.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result

    lhs = ULong32.CreateTruncating(&HF6F2F1F0)
    rhs = ULong32.CreateTruncating(&H0&)
    result = ULong32.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&H0&)
    rhs = ULong32.CreateTruncating(&H0&)
    result = ULong32.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HFFFFFFFF)
    rhs = ULong32.CreateTruncating(&H1&)
    result = ULong32.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = ULong32.CreateTruncating(&HF72)
    rhs = ULong32.CreateTruncating(&H1F2)
    result = ULong32.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
End Sub

Private Sub DisplayMultiply(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print ULong32.ToString(lhs) & " * " & ULong32.ToString(rhs) & " = " & ULong32.ToString(result)
End Sub

Private Sub TestingUInt32PerformanceMultiply()
    Dim result As ULong
    Dim dTime As Double

    Dim lhs  As ULong
    Dim rhs As ULong
    lhs = ULong32.CreateTruncating(&HF62)
    rhs = ULong32.CreateTruncating(&HF6)

    'Perform so overhead of initiliasing ULong32 isn't included in timer calculations
    result = ULong32.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        result = ULong32.Multiply(lhs, rhs)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Multiply duration for 1,000,000 calculations : " & dTime
End Sub


