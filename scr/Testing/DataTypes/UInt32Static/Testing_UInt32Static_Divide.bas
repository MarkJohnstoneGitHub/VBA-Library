Attribute VB_Name = "Testing_UInt32Static_Divide"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.1 March 4, 2023
'@LastModified March 4, 2023

Option Explicit

Private Sub TestingUInt32StaticDivide()
    Dim result As ULong
    Dim dividend  As ULong
    Dim divisor As ULong

    dividend.Value = &HF6F2F1F
    divisor.Value = &HF&
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    dividend.Value = &HF62
    divisor.Value = &HF6
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result

    dividend.Value = &HF6F2F1F0
    divisor.Value = &H7&
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
        
    dividend.Value = &HFFFFFFFF
    divisor.Value = &H2&
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    dividend.Value = &HF72
    divisor.Value = &H1F2
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
End Sub

Private Sub DisplayDivide(ByRef dividend As ULong, ByRef divisor As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(dividend) & " / " & UInt32Static.ToString(divisor) & " = " & UInt32Static.ToString(result)
End Sub

Private Sub TestingUInt32StaticPerformanceDivide()
    Dim result As ULong
    Dim dTime As Double

    Dim dividend  As ULong
    Dim divisor As ULong
    dividend.Value = &HF62
    divisor.Value = &HF6

    'Perform initial subtraction so overhead of initiliasing UInt32Static isn't included in timer calculations
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        result = UInt32Static.Divide(dividend, divisor)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Divide duration for 1,000,000 calculations : " & dTime
End Sub

