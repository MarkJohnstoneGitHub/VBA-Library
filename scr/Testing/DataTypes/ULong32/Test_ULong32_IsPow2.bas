Attribute VB_Name = "Test_ULong32_IsPow2"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32IsPow2()
    Dim val  As ULong
    Dim result As Boolean
    
    val = ULong32.CreateTruncating(&HFFFFFFF8)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = ULong32.CreateTruncating(&H8000&)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = ULong32.CreateTruncating(&H1)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = ULong32.CreateTruncating(&H10)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = ULong32.CreateTruncating(&H0)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = ULong32.CreateTruncating(&H80000000)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = ULong32.CreateTruncating(&H80000)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = ULong32.CreateTruncating(&HF6F2F1F1)
    result = ULong32.IsPow2(val)
    DisplayIsPow2 val, result
End Sub

Private Sub DisplayIsPow2(ByRef val As ULong, ByVal result As Boolean)
    Debug.Print ULong32.ToString(val) & " is power 2 " & result
End Sub
