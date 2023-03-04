Attribute VB_Name = "Testing_UInt32Static_Subtract"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.1 March 4, 2023
'@LastModified March 4, 2023

Option Explicit

Private Sub TestingUInt32Subtract()
    Dim result As ULong
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs.Value = &HF6F2F1F0
    rhs.Value = &H1F3&
    result = UInt32Static.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs.Value = &HF6F2F1F0
    rhs.Value = &HF6FFF0
    result = UInt32Static.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result

    lhs.Value = &HF6F2F1F0
    rhs.Value = &H0&
    result = UInt32Static.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs.Value = &H0&
    rhs.Value = &H0&
    result = UInt32Static.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs.Value = &HFFFFFFFF
    rhs.Value = &H1&
    result = UInt32Static.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    lhs.Value = &HFFFFFFFF
    rhs.Value = &HF5FFEFF2
    result = UInt32Static.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
End Sub

Private Sub DisplaySubtract(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & " - " & UInt32Static.ToString(rhs) & " = " & UInt32Static.ToString(result)
End Sub

Private Sub TestingUInt32PerformanceSubtract()
    Dim result As ULong
    Dim dTime As Double

    Dim lhs  As ULong
    Dim rhs As ULong
    lhs.Value = &HF6F2F1F0
    rhs.Value = &H1F3

    'Perform initial subtraction so overhead of initiliasing UInt32Static isn't included in timer calculations
    result = UInt32Static.Subtract(lhs, rhs)
    DisplaySubtract lhs, rhs, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        result = UInt32Static.Subtract(lhs, rhs)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Subtraction duration for 1,000,000 calculations : " & dTime
End Sub

