Attribute VB_Name = "Test_Int32Bitwise"
'@Folder("VBACorLib.DataTypes")
Option Explicit

Private Sub TestInt32Bitshift()
    Dim int32Input As Long
    Dim binary As String
    Dim result As Long

    Dim numbits As Byte
    
    numbits = 1
    int32Input = "32431"
    
    binary = Int32Bitwise.ToBinary(int32Input, True)
    Debug.Print "Input           :  "; int32Input, binary
    
    result = Int32Bitwise.ShiftRight(int32Input, numbits)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift right bits: " & numbits; result, binary
    
    result = Int32Bitwise.ShiftLeft(result, numbits)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift left bits : " & numbits; result, binary
    Debug.Print
    
    
    int32Input = 129
    
    binary = Int32Bitwise.ToBinary(int32Input, True)
    Debug.Print "Input           :  "; int32Input, binary
    
    result = Int32Bitwise.ShiftLeft(int32Input, numbits)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift right bits: " & numbits; result, binary
    
    result = Int32Bitwise.ShiftRight(result, numbits)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift left bits : " & numbits; result, binary
    Debug.Print
End Sub
