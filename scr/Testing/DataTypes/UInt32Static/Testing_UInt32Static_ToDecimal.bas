Attribute VB_Name = "Testing_UInt32Static_ToDecimal"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.1 March 4, 2023
'@LastModified March 4, 2023

Option Explicit

Private Sub TestingUInt32ToDecimal()
    Dim decResult As Variant
    Dim t1  As ULong
    
    t1.Value = &HF6F2F1F0
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    

    t1 = UInt32Static.CUInt32(&HF6F2F1F0)
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult

    t1.Value = &HFF2F1F
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1.Value = &HFF2F1FFF
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1.Value = &H0
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1.Value = &H107
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1.Value = &HFFFFFFFE
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
    
    t1.Value = &HFFFFFFFF
    decResult = UInt32Static.ToDecimal(t1)
    Debug.Print decResult
End Sub
