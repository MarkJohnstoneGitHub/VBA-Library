Attribute VB_Name = "Testing_UInt32Static_IsEvenInteger"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32StaticIsEvenInteger()
    Dim val  As ULong
    Dim result As Boolean
    
    val = CBytesUInt32(&HF6F2F1F0)
    result = UInt32Static.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = CBytesUInt32(&HF6F2F1F1)
    result = UInt32Static.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = CBytesUInt32(&H1)
    result = UInt32Static.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = CBytesUInt32(&H10)
    result = UInt32Static.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
    
    val = CBytesUInt32(&H0)
    result = UInt32Static.IsEvenInteger(val)
    DisplayIsEvenInteger val, result
End Sub

Private Sub DisplayIsEvenInteger(ByRef val As ULong, ByVal result As Boolean)
    Debug.Print UInt32Static.ToString(val) & " is even " & result
End Sub
