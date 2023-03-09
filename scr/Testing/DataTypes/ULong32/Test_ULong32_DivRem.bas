Attribute VB_Name = "Test_ULong32_DivRem"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32DivRem()
    Dim quotient As ULong
    Dim dividend  As ULong
    Dim divisor As ULong
    Dim remainder As ULong

    dividend = ULong32.CreateTruncating(&HF6F2F1F)
    divisor = ULong32.CreateTruncating(&HF&)
    quotient = ULong32.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
    
    dividend = ULong32.CreateTruncating(&HF62)
    divisor = ULong32.CreateTruncating(&HF6)
    quotient = ULong32.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder

    dividend = ULong32.CreateTruncating(&HF6F2F1F0)
    divisor = ULong32.CreateTruncating(&H7&)
    quotient = ULong32.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
        
    dividend = ULong32.CreateTruncating(&HFFFFFFFF)
    divisor = ULong32.CreateTruncating(&H2&)
    quotient = ULong32.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
    
    dividend = ULong32.CreateTruncating(&HF72)
    divisor = ULong32.CreateTruncating(&H1F2)
    quotient = ULong32.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
End Sub

Private Sub DisplayDivRem(ByRef dividend As ULong, ByRef divisor As ULong, ByRef quotient As ULong, ByRef remainder As ULong)
    Debug.Print ULong32.ToString(dividend) & " / " & ULong32.ToString(divisor) & " = " & ULong32.ToString(quotient) & " Remainder " & ULong32.ToString(remainder)
End Sub

Private Sub TestingULong32PerformanceDivRem()
    Dim quotient As ULong
    Dim dTime As Double

    Dim dividend  As ULong
    Dim divisor As ULong
    Dim remainder As ULong
    
    dividend = ULong32.CreateTruncating(&HF62)
    divisor = ULong32.CreateTruncating(&HF6)

    'Perform initial subtraction so overhead of initiliasing ULong32 isn't included in timer calculations
    quotient = ULong32.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        quotient = ULong32.DivRem(dividend, divisor, remainder)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Divide duration for 1,000,000 calculations : " & dTime
End Sub
