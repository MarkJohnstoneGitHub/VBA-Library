Attribute VB_Name = "Testing_UInt32Static_DivRem"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 28, 2023
'@LastModified February 28, 2023

Option Explicit

Private Sub TestingUInt32StaticDivRem()
    Dim quotient As ULong
    Dim dividend  As ULong
    Dim divisor As ULong
    Dim remainder As ULong

    dividend.Value = &HF6F2F1F
    divisor.Value = &HF&
    quotient = UInt32Static.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
    
    dividend.Value = &HF62
    divisor.Value = &HF6
    quotient = UInt32Static.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder

    dividend.Value = &HF6F2F1F0
    divisor.Value = &H7&
    quotient = UInt32Static.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
        
    dividend.Value = &HFFFFFFFF
    divisor.Value = &H2&
    quotient = UInt32Static.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
    
    dividend.Value = &HF72
    divisor.Value = &H1F2
    quotient = UInt32Static.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
End Sub

Private Sub DisplayDivRem(ByRef dividend As ULong, ByRef divisor As ULong, ByRef quotient As ULong, ByRef remainder As ULong)
    Debug.Print UInt32Static.ToString(dividend) & " / " & UInt32Static.ToString(divisor) & " = " & UInt32Static.ToString(quotient) & " Remainder " & UInt32Static.ToString(remainder)
End Sub

Private Sub TestingUInt32StaticPerformanceDivRem()
    Dim quotient As ULong
    Dim dTime As Double

    Dim dividend  As ULong
    Dim divisor As ULong
    Dim remainder As ULong
    
    dividend.Value = &HF62
    divisor.Value = &HF6

    'Perform initial subtraction so overhead of initiliasing UInt32Static isn't included in timer calculations
    quotient = UInt32Static.DivRem(dividend, divisor, remainder)
    DisplayDivRem dividend, divisor, quotient, remainder
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        quotient = UInt32Static.DivRem(dividend, divisor, remainder)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Divide duration for 1,000,000 calculations : " & dTime
End Sub
