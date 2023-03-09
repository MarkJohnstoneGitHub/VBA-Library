Attribute VB_Name = "Test_UInt32_Add"
'@Folder "Testing.VBACorLib.DataTypes.UInt32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.0 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingUInt32Add()
    Dim result As UInt32

    Dim t1  As UInt32
    Dim t2 As UInt32
    
    Set t1 = UInt32.CreateTruncating(&HF6F2F1F0)
    Set t2 = UInt32.CreateTruncating(&H1F3&)
    Set result = UInt32.Add(t1, t2)
    DisplayAddition t1, t2, result

    Set t1 = UInt32.CreateTruncating(&HF6F2F1F0)
    Set t2 = UInt32.CreateTruncating(&H1F3&)
    Set result = UInt32.Add(t1, t2)
    DisplayAddition t1, t2, result

    Set t1 = UInt32.CreateTruncating(&HFF2F1F)
    Set t2 = UInt32.CreateTruncating(&H1F364)
    Set result = UInt32.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    Set t1 = UInt32.CreateTruncating(&HFF2F1F)
    Set t2 = UInt32.CreateTruncating(&H0&)
    Set result = UInt32.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    Set t1 = UInt32.CreateTruncating(&H0)
    Set t2 = UInt32.CreateTruncating(&HFF2F1F)
    Set result = UInt32.Add(t1, t2)
    DisplayAddition t1, t2, result
    
    Set t1 = UInt32.CreateTruncating(&HFFFFFFFE)
    Set t2 = UInt32.CreateTruncating(&H1)
    Set result = UInt32.Add(t1, t2)
    DisplayAddition t1, t2, result
End Sub

Private Sub TestUInt32AddOverflow()
    Dim result As UInt32
    Dim t1  As UInt32
    Dim t2 As UInt32
    
    Set t1 = UInt32.CreateTruncating(&HFFFFFFFF)
    Set t2 = UInt32.CreateTruncating(&H1)
    Set result = UInt32.Add(t1, t2)
    DisplayAddition t1, t2, result
End Sub

Private Sub DisplayAddition(ByRef lhs As UInt32, ByRef rhs As UInt32, ByRef result As UInt32)
    Debug.Print lhs.ToString() & " + " & rhs.ToString() & " = " & result.ToString()
End Sub

Private Sub TestingUInt32PerformanceAddition()
    Dim result As UInt32
    Dim dTime As Double

    Dim lhs  As UInt32
    Dim rhs As UInt32
    Set lhs = UInt32.CreateTruncating(&HF62)
    Set rhs = UInt32.CreateTruncating(&HF6)

    'Perform initial subtraction so overhead of initiliasing UInt32 isn't included in timer calculations
    Set result = UInt32.Add(lhs, rhs)
    DisplayAddition lhs, rhs, result
    
    Dim i As Long
    ' Initialize
    dTime = MicroTimer

    For i = 1 To 1000000
        UInt32.Add lhs, rhs
        Set result = UInt32.Add(lhs, rhs)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print VBA.vbNewLine & "Addition duration for ADD 1,000,000 calculations : " & dTime
End Sub



