VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Int32Bitwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rubberduck annotations
'@Folder("VBACorLib.DataTypes.Bitwise")
'@PredeclaredId
'@Singleton

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.03 February 15, 2023
'@LastModified  February 15, 2023

''
'@Remarks
'   Bitwise operations for the Int32 data type.
'
'   Bit shifting involves moving bits one or more steps in either the left or right direction.
'   When the bits are shifted one step the bit that is located furthest in the shift direction will
'   fall away and a new bit will be added at the opposite end. The value of the new bit depends on
'   what type of shift operation is used.

' Logical shift
'   With logical shifts the new bits that are shifted in always get the value zero. It's commonly
'   used with unsigned integers (i.e. integers that can't be negative) or when more interested in the
'   actual bits than the value that they represent.

'   For positive integers, a step with logical left shift is the same as multiply by two, and a step
'   with logical right shift is the same as integer division by two, so by doing multiple steps it is
'   possible to multiply and divide by 2^n, where n is the number of steps, as long as the result fits
'   in the number of bits that is being used. If a multiplication or a division can be replaced with a
'   shift operation it is often slightly faster for the computer to calculate.

' Arithmetic shift
'   Arithmetic shifts are suitable for signed integers (i.e. integers that can be both positive and
'   negative) that uses two's complement representation for negative numbers.
'
'   Arithmetic left shift is identical to logical left shift and can be used in the same way to
'   multiply, both positive and negative values, by two.
'
'   With arithmetic right shift new bits get the same value as the sign bit (the leftmost bit).
'   This ensures that the sign (+/-) remains the same before and after. One step with arithmetic
'   right shift is almost the same as integer division by two. The difference is that the result is
'   always rounded down (towards minus infinity) instead of towards zero.
'
' Circular shift
'   Circular shifts, also called rotations, use the bit that got shifted out at one end and inserts
'   it back as the new bit value at the other end. Circular shifts are often used for cryptographic
'   applications and are suitable when it is desirable to not lose any bit values.
'
' Rotate through carry
'   The value of the last bit that got shifted out is normally stored in a carry flag. A special
'   type of circular shift, called rotate through carry, uses the old value of this flag for the
'   bit that is shifted in.
'
'   Rotate through carry can be used to shift larger values than the computer can normally handle.
'   For example, if a computer can only perform shifts on 32 bits at a time, but we want to perform
'   an arithmetic right shift on a 64-bit value, we can do the calculations in two steps. First we
'   perform an arithmetic right shift on the half containing the most significant bits. The bit that
'   got shifted out will be stored in the carry flag. To finish the calculation we then perform a
'   rotate through carry operation on the second half.
'
'@References
' https://home.adelphi.edu/~siegfried/cs174/174l7.pdf
' https://onlinetoolz.net/bitshift
' https://stackoverflow.com/questions/25981365/converting-from-a-double-to-an-integer-for-bitwise-operations
' https://gist.github.com/fernandofig/938cad4e0235a1f8e249
' https://graphics.stanford.edu/~seander/bithacks.html
' https://www.geeksforgeeks.org/detect-if-two-integers-have-opposite-signs/
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Numerics/BitOperations.cs
' http://something4skiddie.blogspot.com/2012/11/bit-shift.html

Option Explicit

'------------------------------------------------------------------
'Constants
'------------------------------------------------------------------
Private Const CHAR_BIT As Long = 8
Private Const SIZEOF_INT32 As Long = 4
Private Const SIZEOF_INT32_BITS As Long = SIZEOF_INT32 * CHAR_BIT

Private Const MIN_BITSHIFT As Long = 0
Private Const MAX_BITSHIFT As Long = 32

Private Const MASK As Long = &HFFFFFFFF
Private Const MOST_SIGNIFICANT_BIT_MASK As Long = &H80000000

Private Const LEAST_SIGNIFICANT_BIT As Long = 1
Private Const MOST_SIGNIFICANT_BIT As Long = 31
Private Const UINT32_MAXVALUE As Long = &HFFFFFFFF

'------------------------------------------------------------------
'Types
'------------------------------------------------------------------
Private Type TInt32Bitwise
    Powers() As Long    'Used for bitwise operations
    Masks() As Long     'Used for bitwise operations
    RevMasks() As Long  'Used for bitwise operations
End Type

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As TInt32Bitwise

'------------------------------------------------------------------
'Constructors and destructors
'------------------------------------------------------------------
Private Sub Class_Initialize()
    'Initialise bitwise arrays
    InitPowers
    InitMasks
    InitRevMasks
End Sub

'------------------------------------------------------------------
'Public Methods
'------------------------------------------------------------------
Public Function Negate(ByVal val As Long) As Long
    Negate = Not val
End Function

''
'@Static
'@Description "Bitwise logical left-shift operation, by the given number of bits."
'@Parameters
'   val: Long value
'   offset: The number of bits to left-shift by.
'@Returns Long
'   Returns a Int32 of the bitwise logical left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
Public Function ShiftLeft(ByVal val As Long, ByVal offset As Long) As Long
    If offset > MIN_BITSHIFT And offset < SIZEOF_INT32_BITS Then
        ShiftLeft = (val And this.Masks(31 - offset)) * _
                                this.Powers(offset)
        If val And this.Powers(31 - offset) Then
            ShiftLeft = ShiftLeft Or MOST_SIGNIFICANT_BIT_MASK ' set bit 31
        End If
    ElseIf offset <= 0 Then
        ShiftLeft = val
    ElseIf offset >= MAX_BITSHIFT Then
        ShiftLeft = &H0
    End If
End Function

''
'@Static
'@Description "Bitwise logical right-shift operation, by the given number of bits."
'@Parameters
'   val: Long value
'   offset: The number of bits to right-shift by.
'@Returns Long
'   Returns a Int32 of the bitwise logical right-shift operation, by the given number of bits.
'@Remarks
'   Right shifting is equal to dividing value by 2^Shift.
Public Function ShiftRight(ByVal val As Long, ByVal offset As Long) As Long
    If offset = 31 Then
        If val And &H80000000 Then
            ShiftRight = 1
        Else
            ShiftRight = 0
        End If
        Exit Function
    End If

    '@TODO Test to check for Div by 0 errors possible that divisor is incorrect  i.e. divide then apply mask?
    If offset > MIN_BITSHIFT And offset < SIZEOF_INT32_BITS Then
        ShiftRight = (val And this.RevMasks(offset)) \ _
                                 this.Powers(offset) And this.Masks(MAX_BITSHIFT - offset)
    ElseIf offset = MIN_BITSHIFT Then
        ShiftRight = val
    ElseIf offset >= MAX_BITSHIFT Then
        ShiftRight = &H0
    End If
End Function

''
'@Static
'@Description "Bitwise arithmetic left-shift operation, by the given number of bits."
'@Parameters
'   val: Long value
'   offset: The number of bits to left-shift by.
'@Returns Long
'   Returns a Int32 of the bitwise arithmetic left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
''
Public Function ShiftLeftArithmetic(ByVal val As Long, ByVal offset As Long) As Long
    Const MAX_SHIFTLEFT As Long = SIZEOF_INT32_BITS - 1

    If offset > MIN_BITSHIFT And offset < MAX_SHIFTLEFT Then
        If val And MOST_SIGNIFICANT_BIT_MASK Then
            val = val Or this.RevMasks(MAX_SHIFTLEFT - offset)
        Else
            val = val And this.Masks(MAX_SHIFTLEFT - offset)
        End If
        ShiftLeftArithmetic = val * this.Powers(offset)
    ElseIf offset = 0 Then
        ShiftLeftArithmetic = val
    ElseIf offset >= MAX_SHIFTLEFT Then
        'To avoid an overflow error
        If val And MOST_SIGNIFICANT_BIT_MASK Then
            ShiftLeftArithmetic = MOST_SIGNIFICANT_BIT_MASK ' set bit 31
        Else
            ShiftLeftArithmetic = &H0&
        End If
    End If
End Function

''
'@Static
'@Description "Bitwise arithmetic right-shift operation, by the given number of bits."
'@Parameters
'   val: Long value
'   offset: The number of bits to right-shift by.
'@Returns Long
'   Returns a Int32 of the bitwise arithmetic right-shift operation, by the given number of bits.
'@Return: Long value
Public Function ShiftRightArithmetic(ByVal val As Long, ByVal offset As Long) As Long
    Const MAX_SHIFTRIGHT As Long = SIZEOF_INT32_BITS - 1
    
    If offset > MIN_BITSHIFT And offset < MAX_SHIFTRIGHT Then
        ShiftRightArithmetic = (val And this.RevMasks(offset)) \ _
                                    this.Powers(offset)
    ElseIf offset = MIN_BITSHIFT Then
        ShiftRightArithmetic = val
    ElseIf offset >= MAX_SHIFTRIGHT Then
        If val And MOST_SIGNIFICANT_BIT_MASK Then
            ShiftRightArithmetic = &HFFFFFFFF
        Else
            ShiftRightArithmetic = &H0
        End If
    End If
End Function

''
'@Static
'@Description "Rotates the specified value left by the specified number of bits."
'"
'@Parameters
'   val: Long value
'   offset: The number of bits to rotate-left by.
'@Returns Long
'   Returns a Int32 of the bitwise rotate-left, by the given number of bits.
'
'@Remarks
'   Bitwise operatation shifts each bit to the left, with the highest bit copied in the Carry flag
'   and into the lowest bit.
'
'@References
'   https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Numerics/BitOperations.cs
''
Public Function RotateLeft(ByVal val As Long, ByVal offset As Long) As Long
    RotateLeft = ShiftLeft(val, offset) Or ShiftRight(val, (SIZEOF_INT32_BITS - offset))
End Function

''
'@Static
'@Description "Rotates the specified value right by the specified number of bits."
'@Parameters
'   val: Long value
'   offset: The number of bits to rotate-right by.
'@Returns Long
'   Returns a Int32 of the bitwise rotate-right, by the given number of bits.
'@Remarks
'   Bitwise operatation shifts each bit to the right, with the lowest bit copied in the Carry flag
'   and into the highest bit.
''
Public Function RotateRight(ByVal val As Long, ByVal offset As Long) As Long
    RotateRight = ShiftRight(val, offset) Or ShiftLeft(val, (SIZEOF_INT32_BITS - offset))
End Function

''
'@TODO Implement
'@Static
'@Description "Bitwise operatation shifts each bit to the right, copies the Carry flag to the most significant bit and copies the least significant bit into the Carry flag."
'@Parameters
'   val: Long value
'   offset: The number of bits to rotate carry right by.
'@Returns Long
'   Returns a Int32 of the bitwise rotate carray right, by the given number of bits.
'@Remarks
Public Function RotateCarryLeft(ByVal val As Long, ByVal offset As Long) As Long
End Function

''
'@TODO Implement
'@Static
'@Description "Bitwise operatation shifts each bit to the right, copies the Carry flag to the most significant bit and copies the least significant bit into the Carry flag."
'@Parameters
'   val: Long value
'   offset: The number of bits to rotate carry right by.
'@Returns Long
'   Returns a Int32 of the bitwise rotate carray right, by the given number of bits.
'@Remarks
''
Public Function RotateCarryRight(ByVal val As Long, ByVal offset As Long) As Long
End Function

''
'@Static
'@Description "Returns 1 if a 32-bit integer is negative and 0 otherwise."
'@Parameters
'   val: Long
'       value to determine if sign is negative
'
'@Returns Long
'   Returns 1 if a 32-bit integer is negative and 0 otherwise.
'
'@Remarks
'   sign = v >> (sizeof(int) * CHAR_BIT - 1);
'   Evaluates to sign = v >> 31 for 32-bit integers.
'   This trick works because when signed integers are shifted right, the value of the far left bit
'   is copied to the other bits. The far left bit is 1 when the value is negative and 0 otherwise.
'
'@References
'   http://graphics.stanford.edu/~seander/bithacks.html#CopyIntegerSign
''
Public Function SignFlag(ByVal val As Long) As Long
    If val And &H80000000 Then
         SignFlag = 1
    End If
End Function

''
'@Static
'@Description "Compares two Int32/Long for opposite signs. "
'@Parameters
'   t1 Long
'       value to compare signs
'   t2 Long
'       value to compare signs against
'
'@Returns Boolean
'   True if signs are opposite to each other
'
'@Remarks
'   The sign bit is 1 in negative numbers, and 0 in positive numbers. The XOR of x and y will have
'   the sign bit as 1 if they have opposite sign. In other words, XOR of x and y will be negative
'   number if x and y have opposite signs.
'
'@References
'   https://www.geeksforgeeks.org/detect-if-two-integers-have-opposite-signs/
'   https://graphics.stanford.edu/~seander/bithacks.html#DetectOppositeSigns
''
Public Function SignsOpposite(ByVal t1 As Long, ByVal t2 As Long) As Boolean
    SignsOpposite = (t1 Xor t2) < 0
End Function

''
'@Static
'@Description "Converts an Int32 to an binary string"
'@Parameters
'   val: Long value
'   byteGrouping: Boolean
'
'@Returns String
'   Returns a string representing an Int32 as bits.
'@Remarks
Public Function ToBinary(ByVal val As Long, Optional ByVal byteGrouping As Boolean = False) As String
    Const DELIMITER As String = " "
    Dim i
    For i = 0 To 31
        'if digit grouping, every 8 bits insert delimiter/space
        If byteGrouping Then
            If (i Mod 8 = 0) And i <> 0 Then
                ToBinary = ToBinary & DELIMITER
            End If
        End If
        
        If this.Powers(31 - i) And val Then
            ToBinary = ToBinary & "1"
        Else
            ToBinary = ToBinary & "0"
        End If
    Next
End Function

''
'@Static
'@Description "Twos Complement of an Int32."
'@References https://stackoverflow.com/questions/20250456/get-twos-complement-for-a-byte
''
Public Function TwosComplement(ByVal val As Long) As Long
    If val = 0 Then
        TwosComplement = 0
    Else
        TwosComplement = (val Xor UINT32_MAXVALUE) + 1
    End If
End Function

''
'@Static
'@Description "Compares two unsigned longs/UInt32 which their raw values are stored in a Long/Int32."
'@Parameters
'   ulngT1: Long
'       unsigned long value to compare against
'   ulngT2: Long
'       unsigned long value to compare to ulngT1
'
'@Returns Long
'   Returns If greater than returns 1, equals 0, less than -1
'@Remarks
''
Public Function CompareUnsigned(ByRef ulngT1 As Long, ByRef ulngT2 As Long) As Long
    'If the most significant bit are different
    If (ulngT1 Xor ulngT2) < 0 Then
        'If t1 most significant bit set it is the larger unsigned value
        If ulngT1 And &H80000000 Then
             CompareUnsigned = 1
             Exit Function
        Else
            CompareUnsigned = -1
            Exit Function
        End If
    End If

    'If "same sign" compare raw values
    If ulngT1 > ulngT2 Then
        CompareUnsigned = 1
    ElseIf ulngT1 < ulngT2 Then
        CompareUnsigned = -1
    End If
End Function


'------------------------------------------------------------------
'Private Methods
'------------------------------------------------------------------

'2^i
Private Sub InitPowers()
    ReDim this.Powers(SIZEOF_INT32_BITS)
    
    this.Powers(0) = &H1&         ' 00000000000000000000000000000001    1
    this.Powers(1) = &H2&         ' 00000000000000000000000000000010    2
    this.Powers(2) = &H4&         ' 00000000000000000000000000000100    4
    this.Powers(3) = &H8&         ' 00000000000000000000000000001000    8
    
    this.Powers(4) = &H10&        ' 00000000000000000000000000010000    16
    this.Powers(5) = &H20&        ' 00000000000000000000000000100000    32
    this.Powers(6) = &H40&        ' 00000000000000000000000001000000    64
    this.Powers(7) = &H80&        ' 00000000000000000000000010000000    128

    this.Powers(8) = &H100&       ' 00000000000000000000000100000000    256
    this.Powers(9) = &H200&       ' 00000000000000000000001000000000    512
    this.Powers(10) = &H400&      ' 00000000000000000000010000000000    1024
    this.Powers(11) = &H800&      ' 00000000000000000000100000000000    2048
    
    this.Powers(12) = &H1000&     ' 00000000000000000001000000000000    4096
    this.Powers(13) = &H2000&     ' 00000000000000000010000000000000    8192
    this.Powers(14) = &H4000&     ' 00000000000000000100000000000000    16384
    this.Powers(15) = &H8000&     ' 00000000000000001000000000000000    32768
    
    this.Powers(16) = &H10000     ' 00000000000000010000000000000000    65536
    this.Powers(17) = &H20000     ' 00000000000000100000000000000000    131072
    this.Powers(18) = &H40000     ' 00000000000001000000000000000000    262144
    this.Powers(19) = &H80000     ' 00000000000010000000000000000000    524288
    
    this.Powers(20) = &H100000    ' 00000000000100000000000000000000    1048576
    this.Powers(21) = &H200000    ' 00000000001000000000000000000000    2097152
    this.Powers(22) = &H400000    ' 00000000010000000000000000000000    4194304
    this.Powers(23) = &H800000    ' 00000000100000000000000000000000    8388608
    
    this.Powers(24) = &H1000000   ' 00000001000000000000000000000000    16777216
    this.Powers(25) = &H2000000   ' 00000010000000000000000000000000    33554432
    this.Powers(26) = &H4000000   ' 00000100000000000000000000000000    67108864
    this.Powers(27) = &H8000000   ' 00001000000000000000000000000000    134217728
    
    this.Powers(28) = &H10000000  ' 00010000000000000000000000000000    268435456
    this.Powers(29) = &H20000000  ' 00100000000000000000000000000000    536870912
    this.Powers(30) = &H40000000  ' 01000000000000000000000000000000    1073741824
    this.Powers(31) = &H80000000  ' 10000000000000000000000000000000    -2147483648
    this.Powers(32) = &H0&        ' 00000000000000000000000000000000    0
    
End Sub

'2^i-1
Private Sub InitMasks()
    ReDim this.Masks(SIZEOF_INT32_BITS)

    this.Masks(0) = &H0&        ' 00000000000000000000000000000000  0
    this.Masks(1) = &H1&        ' 00000000000000000000000000000001  1
    this.Masks(2) = &H3&        ' 00000000000000000000000000000011  3
    this.Masks(3) = &H7&        ' 00000000000000000000000000000111  7

    this.Masks(4) = &HF&        ' 00000000000000000000000000001111  15
    this.Masks(5) = &H1F&       ' 00000000000000000000000000011111  31
    this.Masks(6) = &H3F&       ' 00000000000000000000000000111111  63
    this.Masks(7) = &H7F&       ' 00000000000000000000000001111111  127

    this.Masks(8) = &HFF&       ' 00000000000000000000000011111111  255
    this.Masks(9) = &H1FF&      ' 00000000000000000000000111111111  511
    this.Masks(10) = &H3FF&     ' 00000000000000000000001111111111  1023
    this.Masks(11) = &H7FF&     ' 00000000000000000000011111111111  2047

    this.Masks(12) = &HFFF&     ' 00000000000000000000111111111111  4095
    this.Masks(13) = &H1FFF&    ' 00000000000000000001111111111111  8191
    this.Masks(14) = &H3FFF&    ' 00000000000000000011111111111111  16383
    this.Masks(15) = &H7FFF&    ' 00000000000000000111111111111111  32767

    this.Masks(16) = &HFFFF&    ' 00000000000000001111111111111111  65535
    this.Masks(17) = &H1FFFF    ' 00000000000000011111111111111111  131071
    this.Masks(18) = &H3FFFF    ' 00000000000000111111111111111111  262143
    this.Masks(19) = &H7FFFF    ' 00000000000001111111111111111111  524287

    this.Masks(20) = &HFFFFF    ' 00000000000011111111111111111111  1048575
    this.Masks(21) = &H1FFFFF   ' 00000000000111111111111111111111  2097151
    this.Masks(22) = &H3FFFFF   ' 00000000001111111111111111111111  4194303
    this.Masks(23) = &H7FFFFF   ' 00000000011111111111111111111111  8388607

    this.Masks(24) = &HFFFFFF   ' 00000000111111111111111111111111  16777215
    this.Masks(25) = &H1FFFFFF  ' 00000001111111111111111111111111  33554431
    this.Masks(26) = &H3FFFFFF  ' 00000011111111111111111111111111  67108863
    this.Masks(27) = &H7FFFFFF  ' 00000111111111111111111111111111  134217727

    this.Masks(28) = &HFFFFFFF  ' 00001111111111111111111111111111  268435455
    this.Masks(29) = &H1FFFFFFF ' 00011111111111111111111111111111  536870911
    this.Masks(30) = &H3FFFFFFF ' 00111111111111111111111111111111  1073741823
    this.Masks(31) = &H7FFFFFFF ' 01111111111111111111111111111111  2147483647

    this.Masks(32) = &HFFFFFFFF ' 11111111111111111111111111111111  -1

End Sub

'-2^i
Private Sub InitRevMasks()
    ReDim this.RevMasks(SIZEOF_INT32_BITS)

    this.RevMasks(0) = &HFFFFFFFF   ' 11111111111111111111111111111111    -1
    this.RevMasks(1) = &HFFFFFFFE   ' 11111111111111111111111111111110    -2
    this.RevMasks(2) = &HFFFFFFFC   ' 11111111111111111111111111111100    -4
    this.RevMasks(3) = &HFFFFFFF8   ' 11111111111111111111111111111000    -8

    this.RevMasks(4) = &HFFFFFFF0   ' 11111111111111111111111111110000    -16
    this.RevMasks(5) = &HFFFFFFE0   ' 11111111111111111111111111100000    -32
    this.RevMasks(6) = &HFFFFFFC0   ' 11111111111111111111111111000000    -64
    this.RevMasks(7) = &HFFFFFF80   ' 11111111111111111111111110000000    -128

    this.RevMasks(8) = &HFFFFFF00   ' 11111111111111111111111100000000    -256
    this.RevMasks(9) = &HFFFFFE00   ' 11111111111111111111111000000000    -512
    this.RevMasks(10) = &HFFFFFC00  ' 11111111111111111111110000000000    -1024
    this.RevMasks(11) = &HFFFFF800  ' 11111111111111111111100000000000    -2048

    this.RevMasks(12) = &HFFFFF000  ' 11111111111111111111000000000000    -4096
    this.RevMasks(13) = &HFFFFE000  ' 11111111111111111110000000000000    -8192
    this.RevMasks(14) = &HFFFFC000  ' 11111111111111111100000000000000    -16384
    this.RevMasks(15) = &HFFFF8000  ' 11111111111111111000000000000000    -32768

    this.RevMasks(16) = &HFFFF0000  ' 11111111111111110000000000000000    -65536
    this.RevMasks(17) = &HFFFE0000  ' 11111111111111100000000000000000    -131072
    this.RevMasks(18) = &HFFFC0000  ' 11111111111111000000000000000000    -262144
    this.RevMasks(19) = &HFFF80000  ' 11111111111110000000000000000000    -524288

    this.RevMasks(20) = &HFFF00000  ' 11111111111100000000000000000000    -1048576
    this.RevMasks(21) = &HFFE00000  ' 11111111111000000000000000000000    -2097152
    this.RevMasks(22) = &HFFC00000  ' 11111111110000000000000000000000    -4194304
    this.RevMasks(23) = &HFF800000  ' 11111111100000000000000000000000    -8388608

    this.RevMasks(24) = &HFF000000  ' 11111111000000000000000000000000    -16777216
    this.RevMasks(25) = &HFE000000  ' 11111110000000000000000000000000    -33554432
    this.RevMasks(26) = &HFC000000  ' 11111100000000000000000000000000    -67108864
    this.RevMasks(27) = &HF8000000  ' 11111000000000000000000000000000    -134217728

    this.RevMasks(28) = &HF0000000  ' 11110000000000000000000000000000    -268435456
    this.RevMasks(29) = &HE0000000  ' 11100000000000000000000000000000    -536870912
    this.RevMasks(30) = &HC0000000  ' 11000000000000000000000000000000    -1073741824
    this.RevMasks(31) = &H80000000  ' 10000000000000000000000000000000    -2147483648
    
    this.RevMasks(32) = &H0&        ' 00000000000000000000000000000000      0
End Sub
