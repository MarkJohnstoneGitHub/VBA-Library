Attribute VB_Name = "Test_ULong32_PopCount"
'@Folder("Testing.VBACorLib.DataTypes.ULong32")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.41 March 21, 2023
'@LastModified  March 24, 2023

Option Explicit

Private Sub TestingULong32_PopCount()
    Dim result As ULong
    Dim val  As ULong
    
    val = ULong32.CreateChecked(0)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
    
    val = ULong32.CreateChecked(4294967295#)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
    
    val = ULong32.CreateChecked(1325)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result

    val = ULong32.CreateChecked(4967295)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result

    val = ULong32.CreateChecked(294967295)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result

    val = ULong32.CreateChecked(394967295)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
    
    val = ULong32.CreateChecked(255)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
    
    val = ULong32.CreateChecked(32767)
    result = ULong32.PopCount(val)
    DisplayPopCount val, result
End Sub

Private Sub DisplayPopCount(ByRef val As ULong, ByRef result As ULong)
    Debug.Print "PopCount(" & ULong32.ToString(val) & ")" & " = " & ULong32.ToString(result)
End Sub

Private Sub TestingULong32Performance_PopCount()

    'Perform so overhead of initiliasing ULong32 isn't included in timer calculations
    Dim val  As ULong
    val = ULong32.CreateChecked(4294967295#)
    Dim result As ULong
    result = ULong32.PopCount(val)

    Dim i As Long
    ' Initialize
    Dim dTime As Double
    dTime = MicroTimer
    For i = 1 To 1000000
        result = ULong32.PopCount(val)
    Next i

    ' Calculate duration.
    dTime = MicroTimer - dTime
    Debug.Print dTime & ","
End Sub















