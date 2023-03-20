Attribute VB_Name = "Test_ULong32_Addition"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingULong32Add()
    Dim result As ULong

    Dim t1  As ULong
    Dim t2 As ULong
    
    t1 = ULong32.CreateTruncating(&HF6F2F1F0)
    t2 = ULong32.CreateTruncating(&H1F3&)
    result = ULong32.Add(t1, t2)
    DisplayAddition t1, t2, result

    t1 = ULong32.CreateTruncating(&HF6F2F1F0)
    t2 = ULong32.CreateTruncating(&H1F3&)
    result = ULong32.Add(t1, t2)
    DisplayAddition t1, t2, result

    t1 = ULong32.CreateTruncating(&HFF2F1F)
    t2 = ULong32.CreateTruncating(&H1F364)
    result = ULong32.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    t1 = ULong32.CreateTruncating(&HFF2F1F)
    t2 = ULong32.CreateTruncating(&H0&)
    result = ULong32.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    t1 = ULong32.CreateTruncating(&H0)
    t2 = ULong32.CreateTruncating(&HFF2F1F)
    result = ULong32.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    t1 = ULong32.CreateTruncating(&HFFFFFFFE)
    t2 = ULong32.CreateTruncating(&H1)
    result = ULong32.Add(t1, t2)
    DisplayAddition t1, t2, result
End Sub

Private Sub TestULong32AddOverflow()
    Dim result As ULong
    Dim t1  As ULong
    Dim t2 As ULong
    
    t1 = ULong32.CreateTruncating(&HFFFFFFFF)
    t2 = ULong32.CreateTruncating(&H1)
    result = ULong32.Add(t1, t2)
    DisplayAddition t1, t2, result
End Sub

Private Sub DisplayAddition(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print ULong32.ToString(lhs) & " + " & ULong32.ToString(rhs) & " = " & ULong32.ToString(result)
End Sub

Private Sub TestingULong32PerformanceAddition()
    Dim result As ULong
    Dim dTime As Double

    Dim lhs  As ULong
    Dim rhs As ULong
    lhs = ULong32.CreateTruncating(&HF6200000)
    rhs = ULong32.CreateTruncating(&HF6)

    'Perform so overhead of initiliasing ULong32 isn't included in timer calculations
    result = ULong32.Add(lhs, rhs)
    DisplayAddition lhs, rhs, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 10000000
        result = ULong32.Add(lhs, rhs)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Addition duration for ADD 100,000,000 calculations : " & dTime
End Sub


