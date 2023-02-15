Attribute VB_Name = "Test_Int32Bitwise"
'@Folder("Testing.VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.04 February 16, 2023
'@LastModified  February 16, 2023

Option Explicit

Private Sub TestingInt32Bitshift()
    Dim value As Long
    Dim binary As String
    Dim result As Long

    Dim offset As Long
    
    offset = 1
    value = "32431"
    
    binary = Int32Bitwise.ToBinary(value, True)
    Debug.Print "Value           :  "; value, binary
    
    result = Int32Bitwise.ShiftRight(value, offset)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift right bits: " & offset; result, binary
    
    result = Int32Bitwise.ShiftLeft(result, offset)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift left bits : " & offset; result, binary
    Debug.Print
    
    value = 129
    binary = Int32Bitwise.ToBinary(value, True)
    Debug.Print "Value           :  "; value, binary
    
    result = Int32Bitwise.ShiftLeft(value, offset)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift right bits: " & offset; result, binary
    
    result = Int32Bitwise.ShiftRight(result, offset)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift left bits : " & offset; result, binary
    Debug.Print
        
    value = &HFF67FF2F
    offset = 30
    Debug.Print "Value:       " & value & " " & VBA.vbTab & Int32Bitwise.ToBinary(value, True)
    result = Int32Bitwise.ShiftRight(value, offset)
    binary = Int32Bitwise.ToBinary(result, True)
    Debug.Print "Shift left bits : " & offset & " " & " " & VBA.vbTab & binary
    Debug.Print

End Sub

Private Sub TestingInt32BitshiftAll()
    Dim value As Long
    Dim offset As Long
    Dim binary As String
    Dim result As Long
    
    value = &H2F67FF2F
    
    For offset = 0 To 32
        Debug.Print "Value:    " & value & " " & VBA.vbTab & Int32Bitwise.ToBinary(value, True)
        result = Int32Bitwise.ShiftRight(value, offset)
        binary = Int32Bitwise.ToBinary(result, True)
        Debug.Print "Shift left bits : " & offset & " " & " " & VBA.vbTab & binary
        Debug.Print
    Next

End Sub

Private Sub TestingInt32BitwiseSignFlag()
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


Private Sub TestingInt32BitwiseSignsOpposite()
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

Private Sub TestingInt32BitwiseTwosComplement()
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
Private Sub TestingInt32BitwiseCompareUnsigned()
    Dim ulngT1 As Long
    Dim ulngT2 As Long
    Dim result As Long
    
    ulngT1 = &HFFFFFFFE     'i.e.unsigned = 4294967294, signed = -2
    ulngT2 = "&H0FFFFFFE"   'i.e.unsigned = 268435454,  signed = 268435454
    Debug.Print "Comparing two unsigned longs"
    result = Int32Bitwise.CompareUnsigned(ulngT1, ulngT2)
    Debug.Print "Result: " & result & " " & "values:  " & "4294967294" & ", " & "268435454"
    Debug.Print
    
    ulngT1 = "&H0FFFFFFE"   'i.e.unsigned = 268435454,  signed = 268435454
    ulngT2 = &HFFFFFFFE     'i.e.unsigned = 4294967294, signed = -2
    Debug.Print "Comparing two unsigned longs"
    result = Int32Bitwise.CompareUnsigned(ulngT1, ulngT2)
    Debug.Print "Result: " & result & " " & "values:  " & "268435454" & ", " & "4294967294"
    Debug.Print
    
    ulngT1 = &HFFFFFFFE    'i.e.unsigned = 4294967294,  signed = -2
    ulngT2 = &HFFFFFFFE     'i.e.unsigned = 4294967294, signed = -2
    Debug.Print "Comparing two unsigned longs"
    result = Int32Bitwise.CompareUnsigned(ulngT1, ulngT2)
    Debug.Print "Result: " & result & " " & "values:  " & "4294967294" & ", " & "4294967294"
    Debug.Print
End Sub


