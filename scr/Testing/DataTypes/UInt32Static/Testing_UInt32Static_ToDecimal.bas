Attribute VB_Name = "Testing_UInt32Static_ToDecimal"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 7, 2023

Option Explicit

Private Sub TestingUInt32ToDecimal()
    Dim decResult As Variant
    Dim t1  As ULong
    
    t1 = CBytesUInt32(&HF6F2F1F0)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = CUInt32("&HF6F2F1F0")
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = UInt32Static.Parse("&HF6F2F1F0")
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = CBytesUInt32(&HFF2F1F)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = CBytesUInt32(&HFF2F1FFF)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = CBytesUInt32(&H0)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = CBytesUInt32(&H107)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = CBytesUInt32(&HFFFFFFFE)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = CBytesUInt32(&HFFFFFFFF)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
End Sub
