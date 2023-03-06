Attribute VB_Name = "Testing_UInt32Static_Addition"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32Add()
    Dim result As ULong

    Dim t1  As ULong
    Dim t2 As ULong
    
    t1 = CBytesUInt32(&HF6F2F1F0)
    t2 = CBytesUInt32(&H1F3&)
    result = UInt32Static.Add(t1, t2)
    DisplayAddition t1, t2, result

    t1 = CBytesUInt32(&HF6F2F1F0)
    t2 = CBytesUInt32(&H1F3&)
    result = UInt32Static.Add(t1, t2)
    DisplayAddition t1, t2, result

    t1 = CBytesUInt32(&HFF2F1F)
    t2 = CBytesUInt32(&H1F364)
    result = UInt32Static.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    t1 = CBytesUInt32(&HFF2F1F)
    t2 = CBytesUInt32(&H0&)
    result = UInt32Static.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    t1 = CBytesUInt32(&H0)
    t2 = CBytesUInt32(&HFF2F1F)
    result = UInt32Static.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    t1 = CBytesUInt32(&HFFFFFFFE)
    t2 = CBytesUInt32(&H1)
    result = UInt32Static.Add(t1, t2)
    DisplayAddition t1, t2, result
End Sub

Private Sub TestUInt32AddOverflow()
    Dim result As ULong
    Dim t1  As ULong
    Dim t2 As ULong
    
    t1 = CBytesUInt32(&HFFFFFFFF)
    t2 = CBytesUInt32(&H1)
    result = UInt32Static.Add(t1, t2)
    DisplayAddition t1, t2, result
End Sub

Private Sub DisplayAddition(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & " + " & UInt32Static.ToString(rhs) & " = " & UInt32Static.ToString(result)
End Sub

Private Sub TestingUInt32PerformanceAddition()
    Dim result As ULong
    Dim dTime As Double

    Dim lhs  As ULong
    Dim rhs As ULong
    lhs = CBytesUInt32(&HF62)
    rhs = CBytesUInt32(&HF6)

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
    Debug.Print VBA.vbNewLine & "Addition duration for ADD 1,000,000 calculations : " & dTime
End Sub

