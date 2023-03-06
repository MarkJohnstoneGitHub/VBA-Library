Attribute VB_Name = "Testing_UInt32Static_Divide"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32StaticDivide()
    Dim result As ULong
    Dim dividend  As ULong
    Dim divisor As ULong

    dividend = CBytesUInt32(&HF6F2F1F)
    divisor = CBytesUInt32(&HF&)
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    dividend = CBytesUInt32(&HF62)
    divisor = CBytesUInt32(&HF6)
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result

    dividend = CBytesUInt32(&HF6F2F1F0)
    divisor = CBytesUInt32(&H7&)
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
        
    dividend = CBytesUInt32(&HFFFFFFFF)
    divisor = CBytesUInt32(&H2&)
    result = UInt32Static.Divide(dividend, divisor)
    DisplayDivide dividend, divisor, result
    
    dividend = CBytesUInt32(&HF72)
    divisor = CBytesUInt32(&H1F2)
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
    dividend = CBytesUInt32(&HF62)
    divisor = CBytesUInt32(&HF6)

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

