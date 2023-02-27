Attribute VB_Name = "Testing_UInt32Static_Addition"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 27, 2023
'@LastModified February 27, 2023

Option Explicit

Private Sub TestingUInt32Add()
    Dim result As ULong

    Dim t1  As ULong
    Dim t2 As ULong

    t1.Value = &HF6F2F1F0
    t2.Value = &H1F3&

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)

    t1.Value = &HFF2F1F
    t2.Value = &H1F364

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)
    
    t1.Value = &HFF2F1F
    t2.Value = &H0&

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)
    
    t1.Value = &H0
    t2.Value = &HFF2F1F

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)
End Sub

Private Sub DisplayAddition(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & " + " & UInt32Static.ToString(rhs) & " = " & UInt32Static.ToString(result)
End Sub


Private Sub TestingUInt32PerformanceAddition()
    Dim result As ULong
    Dim dTime As Double

    Dim lhs  As ULong
    Dim rhs As ULong
    lhs.Value = &HF62
    rhs.Value = &HF6

    'Perform initial subtraction so overhead of initiliasing UInt32Static isn't included in timer calculations
    result = UInt32Static.Add(lhs, rhs)
    DisplayAddition lhs, rhs, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        result = UInt32Static.Add(lhs, rhs)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Addition duration for 1,000,000 calculations : " & dTime
End Sub

