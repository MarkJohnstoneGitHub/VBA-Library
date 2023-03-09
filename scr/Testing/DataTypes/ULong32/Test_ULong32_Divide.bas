Attribute VB_Name = "Test_ULong32_Divide"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32Divide()
    Dim result As ULong
    Dim dividend  As ULong
    Dim divisor As ULong

    dividend = ULong32.CreateTruncating(&HF6F2F1F)
    divisor = ULong32.CreateTruncating(&HF&)
    result = ULong32.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    dividend = ULong32.CreateTruncating(&HF62)
    divisor = ULong32.CreateTruncating(&HF6)
    result = ULong32.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result

    dividend = ULong32.CreateTruncating(&HF6F2F1F0)
    divisor = ULong32.CreateTruncating(&H7&)
    result = ULong32.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
        
    dividend = ULong32.CreateTruncating(&HFFFFFFFF)
    divisor = ULong32.CreateTruncating(&H2&)
    result = ULong32.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    dividend = ULong32.CreateTruncating(&HF72)
    divisor = ULong32.CreateTruncating(&H1F2)
    result = ULong32.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
End Sub

Private Sub DisplayDivide(ByRef dividend As ULong, ByRef divisor As ULong, ByRef result As ULong)
    Debug.Print ULong32.ToString(dividend) & " / " & ULong32.ToString(divisor) & " = " & ULong32.ToString(result)
End Sub

Private Sub TestingUInt32StaticPerformanceDivide()
    Dim result As ULong
    Dim dTime As Double

    Dim dividend  As ULong
    Dim divisor As ULong
    dividend = ULong32.CreateTruncating(&HF62)
    divisor = ULong32.CreateTruncating(&HF6)

    'Perform initial subtraction so overhead of initiliasing ULong32 isn't included in timer calculations
    result = ULong32.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        result = ULong32.Divide(dividend, divisor)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Divide duration for 1,000,000 calculations : " & dTime
End Sub

