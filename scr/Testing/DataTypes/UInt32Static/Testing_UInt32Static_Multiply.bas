Attribute VB_Name = "Testing_UInt32Static_Multiply"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32Multiply()
    Dim result As ULong
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs = CBytesUInt32(&HF6F2F1F)
    rhs = CBytesUInt32(&HF&)
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = CBytesUInt32(&HF62)
    rhs = CBytesUInt32(&HF6)
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result

    lhs = CBytesUInt32(&HF6F2F1F0)
    rhs = CBytesUInt32(&H0&)
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = CBytesUInt32(&H0&)
    rhs = CBytesUInt32(&H0&)
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = CBytesUInt32(&HFFFFFFFF)
    rhs = CBytesUInt32(&H1&)
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs = CBytesUInt32(&HF72)
    rhs = CBytesUInt32(&H1F2)
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
End Sub

Private Sub DisplayMultiply(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & " * " & UInt32Static.ToString(rhs) & " = " & UInt32Static.ToString(result)
End Sub

Private Sub TestingUInt32PerformanceMultiply()
    Dim result As ULong
    Dim dTime As Double

    Dim lhs  As ULong
    Dim rhs As ULong
    lhs = CBytesUInt32(&HF62)
    rhs = CBytesUInt32(&HF6)

    'Perform initial subtraction so overhead of initiliasing UInt32Static isn't included in timer calculations
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        result = UInt32Static.Multiply(lhs, rhs)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Multiply duration for 1,000,000 calculations : " & dTime
End Sub

