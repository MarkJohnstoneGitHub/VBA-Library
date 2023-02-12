Attribute VB_Name = "Test_Int32Bitwise"
'@Folder("Testing.VBACorLib.DataTypes")
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


Private Sub Testing_Int32Bitwise_SignNegative()
    Dim val As Long
    
    val = -100
    Debug.Print "Value: " & val, "Sign negative : " & Int32Bitwise.SignNegative(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)

    val = 100
    Debug.Print "Value: " & val, "Sign negative : " & Int32Bitwise.SignNegative(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)
    
    val = 0
    Debug.Print "Value: " & val, "Sign negative : " & Int32Bitwise.SignNegative(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)
    
    val = &HFFFFFFFF
    Debug.Print "Value: " & val, "Sign negative : " & Int32Bitwise.SignNegative(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)
End Sub


Private Sub Testing_Int32Bitwise_SignsOpposite()
    Dim t1 As Long
    Dim t2 As Long
    
    t1 = 0
    t2 = 0
    Debug.Print "Int32 Signs Opposite: " & t1 & ", "; t2 & " : " & Int32Bitwise.SignsOpposite(t1, t2)
    
    t1 = -1
    t2 = 0
    Debug.Print "Int32 Signs Opposite: " & t1 & ", "; t2 & " : " & Int32Bitwise.SignsOpposite(t1, t2)
    
    t1 = 2342134
    t2 = 544
    Debug.Print "Int32 Signs Opposite: " & t1 & ", "; t2 & " : " & Int32Bitwise.SignsOpposite(t1, t2)
    
    t1 = -1
    t2 = 23355
    Debug.Print "Int32 Signs Opposite: " & t1 & ", "; t2 & " : " & Int32Bitwise.SignsOpposite(t1, t2)
    
    
    t1 = 234234
    t2 = -34324
    Debug.Print "Int32 Signs Opposite: " & t1 & ", "; t2 & " : " & Int32Bitwise.SignsOpposite(t1, t2)
    
End Sub
