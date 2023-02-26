Attribute VB_Name = "Testing_UInt32Static_Multiply"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 27, 2023
'@LastModified February 27, 2023

Option Explicit

Private Sub TestingUInt32Multiply()
    Dim result As ULong
    Dim lhs  As ULong
    Dim rhs As ULong

    lhs.Value = &HF6F2F1F
    rhs.Value = &HF&
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs.Value = &HF62
    rhs.Value = &HF6
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result

    lhs.Value = &HF6F2F1F0
    rhs.Value = &H0&
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs.Value = &H0&
    rhs.Value = &H0&
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs.Value = &HFFFFFFFF
    rhs.Value = &H1&
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
    
    lhs.Value = &HF72
    rhs.Value = &H1F2
    result = UInt32Static.Multiply(lhs, rhs)
    DisplayMultiply lhs, rhs, result
End Sub

Private Sub DisplayMultiply(ByRef lhs As ULong, ByRef rhs As ULong, ByRef result As ULong)
    Debug.Print UInt32Static.ToString(lhs) & " * " & UInt32Static.ToString(rhs) & " = " & UInt32Static.ToString(result)
End Sub
