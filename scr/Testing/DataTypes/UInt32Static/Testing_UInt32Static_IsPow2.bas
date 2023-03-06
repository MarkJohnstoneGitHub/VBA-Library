Attribute VB_Name = "Testing_UInt32Static_IsPow2"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32StaticIsPow2()
    Dim val  As ULong
    Dim result As Boolean
    
    val = CBytesUInt32(&HFFFFFFF8)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = CBytesUInt32(&H8000&)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = CBytesUInt32(&H1)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = CBytesUInt32(&H10)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = CBytesUInt32(&H0)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = CBytesUInt32(&H80000000)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = CBytesUInt32(&H80000)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
    
    val = CBytesUInt32(&HF6F2F1F1)
    result = UInt32Static.IsPow2(val)
    DisplayIsPow2 val, result
End Sub

Private Sub DisplayIsPow2(ByRef val As ULong, ByVal result As Boolean)
    Debug.Print UInt32Static.ToString(val) & " is power 2 " & result
End Sub
