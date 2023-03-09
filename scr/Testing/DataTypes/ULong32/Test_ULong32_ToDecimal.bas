Attribute VB_Name = "Test_ULong32_ToDecimal"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32ToDecimal()
    Dim decResult As Variant
    Dim t1  As ULong
    
    t1 = ULong32.CreateTruncating(&HF6F2F1F0)
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = ULong32.Parse("&HF6F2F1F0")
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = ULong32.CreateTruncating(&HFF2F1F)
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = ULong32.CreateTruncating(&HFF2F1FFF)
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = ULong32.CreateTruncating(&H0)
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = ULong32.CreateTruncating(&H107)
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = ULong32.CreateTruncating(&HFFFFFFFE)
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
    
    t1 = ULong32.CreateTruncating(&HFFFFFFFF)
    decResult = ULong32.ToDecimal(t1)
    Debug.Print decResult
End Sub
