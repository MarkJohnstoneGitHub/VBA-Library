Attribute VB_Name = "Test_Int32Bitwise"
'@Folder("Testing.VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.03 February 15, 2023
'@LastModified  February 15, 2023

Option Explicit

Private Sub TestInt32Bitshift()
    Dim int32Input As Long
    Dim binary As String
    Dim result As Long

    Dim numbits As Long
    
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
    Debug.Print "Value: " & val, "Sign flag : " & Int32Bitwise.SignFlag(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)

    val = 100
    Debug.Print "Value: " & val, "Sign flag : " & Int32Bitwise.SignFlag(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)
    
    val = 0
    Debug.Print "Value: " & val, "Sign flag : " & Int32Bitwise.SignFlag(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)
    
    val = &HFFFFFFFF
    Debug.Print "Value: " & val, "Sign flag : " & Int32Bitwise.SignFlag(val) & VBA.vbTab & Int32Bitwise.ToBinary(val, True)
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

Private Sub Testing_Int32Bitwise_TwosComplement()
    Dim val As Long
    Dim result As Long
    
    val = 0
    result = Int32Bitwise.TwosComplement(val)
    Debug.Print "Two's complement"
    Debug.Print "Value:  " & val & VBA.vbTab; Int32Bitwise.ToBinary(val, True)
    Debug.Print "Result: " & result & VBA.vbTab; Int32Bitwise.ToBinary(result, True)
    
    val = -5433
    result = Int32Bitwise.TwosComplement(val)
    Debug.Print "Two's complement"
    Debug.Print "Value:  " & val & VBA.vbTab; Int32Bitwise.ToBinary(val, True)
    Debug.Print "Result: " & result & VBA.vbTab; Int32Bitwise.ToBinary(result, True)
End Sub

'@References
' Conversion of signed and unsigned longs note use 8 character hex values
' https://www.binaryconvert.com/convert_signed_int.html
' https://www.binaryconvert.com/convert_unsigned_int.html
Private Sub Testing_Int32Bitwise_CompareUnsigned()
    Dim ulngT1 As Long
    Dim ulngT2 As Long
    Dim result As Long
    
    ulngT1 = &HFFFFFFFE     'i.e.unsigned = 4294967294, signed = -2
    ulngT2 = "&H0FFFFFFE"   'i.e.unsigned = 268435454,  signed = 268435454
    Debug.Print "Comparing two unsigned longs"
    result = Int32Bitwise.CompareUnsigned(ulngT1, ulngT2)
    Debug.Print "Result: " & result & " " & "values:  " & "4294967294" & ", " & "268435454"
    
    ulngT1 = "&H0FFFFFFE"   'i.e.unsigned = 268435454,  signed = 268435454
    ulngT2 = &HFFFFFFFE     'i.e.unsigned = 4294967294, signed = -2
    Debug.Print "Comparing two unsigned longs"
    result = Int32Bitwise.CompareUnsigned(ulngT1, ulngT2)
    Debug.Print "Result: " & result & " " & "values:  " & "268435454" & ", " & "4294967294"
    
    ulngT1 = &HFFFFFFFE    'i.e.unsigned = 4294967294,  signed = -2
    ulngT2 = &HFFFFFFFE     'i.e.unsigned = 4294967294, signed = -2
    Debug.Print "Comparing two unsigned longs"
    result = Int32Bitwise.CompareUnsigned(ulngT1, ulngT2)
    Debug.Print "Result: " & result & " " & "values:  " & "4294967294" & ", " & "4294967294"
End Sub


