Attribute VB_Name = "Test_ULong32_IsEvenInteger"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32IsEvenInteger()
    Dim val  As ULong
    Dim result As Boolean
    
    val = ULong32.CreateTruncating(&HF6F2F1F0)
    result = ULong32.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = ULong32.CreateTruncating(&HF6F2F1F1)
    result = ULong32.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = ULong32.CreateTruncating(&H1)
    result = ULong32.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = ULong32.CreateTruncating(&H10)
    result = ULong32.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = ULong32.CreateTruncating(&H0)
    result = ULong32.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
End Sub

Private Sub DisplayIsEvenInteger(ByRef val As ULong, ByVal result As Boolean)
    Debug.Print ULong32.ToString(val) & " is even " & result
End Sub
