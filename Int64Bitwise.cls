VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Int64Bitwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBACorLib.DataTypes")
'@PredeclaredId

'@Singleton Class

'@Version v1.0 (Date January 30, 2023)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/

'@Author markjohnstone@hotmail.com
'@LastModified January 30,2023
'
'@Remarks
' Bitwise operations for the Int64 data type.
'
' Bit shifting involves moving bits one or more steps in either the left or right direction.
' When the bits are shifted one step the bit that is located furthest in the shift direction will
' fall away and a new bit will be added at the opposite end. The value of the new bit depends on
' what type of shift operation is used.

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
'https://home.adelphi.edu/~siegfried/cs174/174l7.pdf
'https://onlinetoolz.net/bitshift

Option Explicit

'------------------------------------------------------------------
'Constants
'------------------------------------------------------------------
Private Const SIZEOF_INT64 As Long = 64
Private Const MIN_BITSHIFT As Byte = 0

'------------------------------------------------------------------
'Types
'------------------------------------------------------------------
Private Type TBitwiseInt64
    Powers() As LongLong    'Used for bitwise operations
    Masks() As LongLong     'Used for bitwise operations
    RevMasks() As LongLong  'Used for bitwise operations
End Type

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As TBitwiseInt64

'------------------------------------------------------------------
'Constructors and destructors
'------------------------------------------------------------------
Private Sub Class_Initialize()
    InitBitwiseArrays
End Sub

'------------------------------------------------------------------
'Public Methods
'------------------------------------------------------------------

'@Static
'@Description "Bitwise logical right-shift operation, by the given number of bits."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to right-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise logical right-shift operation, by the given number of bits.
'@Remarks
'   Right shifting is equal to dividing value by 2^Shift.
Public Function ShiftRight(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
    Const MAX_BITSHIFT As Long = 64

    If bitsShifted > MIN_BITSHIFT And bitsShifted < MAX_BITSHIFT Then
        ShiftRight = (val And this.RevMasks(bitsShifted)) \ _
                                 this.Powers(bitsShifted) And this.Masks(MAX_BITSHIFT - bitsShifted)
    ElseIf bitsShifted <= MIN_BITSHIFT Then
        ShiftRight = val
    ElseIf bitsShifted >= MAX_BITSHIFT Then
        ShiftRight = &H0
    End If
End Function

'@Static
'@Description "Bitwise arithmetic right-shift operation, by the given number of bits."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to right-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise arithmetic right-shift operation, by the given number of bits.
'@Return: LongLong value
Public Function ShiftRightArithmetic(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
    Const MAX_BITSHIFT As Long = SIZEOF_INT64 - 1
    Const LAST_BIT_MASK As LongLong = "&H8000000000000000"

    If bitsShifted > 0 And bitsShifted < MAX_BITSHIFT Then
        ShiftRightArithmetic = (val And this.RevMasks(bitsShifted)) \ _
                                    this.Powers(bitsShifted)
'    ElseIf bitsShifted <= 0 Then
'        ShiftRightArithmetic = val
    ElseIf bitsShifted >= MAX_BITSHIFT Then
        If val And LAST_BIT_MASK Then
            ShiftRightArithmetic = "&HFFFFFFFFFFFFFFFF"
        Else
            ShiftRightArithmetic = &H0
        End If
    End If
End Function

'@Static
'@Description "Bitwise logical left-shift operation, by the given number of bits."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to left-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise logical left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
Public Function ShiftLeft(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
    Const LAST_BIT_MASK As LongLong = "&H8000000000000000"
    If bitsShifted > MIN_BITSHIFT And bitsShifted < SIZEOF_INT64 Then
        ShiftLeft = (val And this.Masks(63 - bitsShifted)) * _
                                this.Powers(bitsShifted)
        If val And this.Powers(63 - bitsShifted) Then
            ShiftLeft = ShiftLeft Or LAST_BIT_MASK ' set bit 63
        End If
    ElseIf bitsShifted <= 0 Then
        ShiftLeft = val
    ElseIf bitsShifted >= SIZEOF_INT64 Then
        ShiftLeft = &H0&
    End If
End Function

'@Static
'@Description "Bitwise arithmetic left-shift operation, by the given number of bits."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to left-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise arithmetic left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
Public Function ShiftLeftArithmetic(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
    Const MAX_BITSHIFT As Long = SIZEOF_INT64 - 1
    Const LAST_BIT_MASK As LongLong = "&H8000000000000000"

    If bitsShifted > MIN_BITSHIFT And bitsShifted < MAX_BITSHIFT Then
        If val And LAST_BIT_MASK Then
            val = val Or this.RevMasks(MAX_BITSHIFT - bitsShifted)
        Else
            val = val And this.Masks(MAX_BITSHIFT - bitsShifted)
        End If
        ShiftLeftArithmetic = val * this.Powers(bitsShifted)
    ElseIf bitsShifted = 0 Then
        ShiftLeftArithmetic = val
    ElseIf bitsShifted >= MAX_BITSHIFT Then
        'To avoid an overflow error
        If val And LAST_BIT_MASK Then
            ShiftLeftArithmetic = LAST_BIT_MASK ' set bit 63
        Else
            ShiftLeftArithmetic = &H0
        End If
    End If
    
End Function


'@TODO Implement
'@Static
'@Description "Bitwise operatation shifts each bit to the left, with the highest bit copied in the Carry flag and into the lowest bit."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to rotate-left by.
'@Returns LongLong
'   Returns a Int64 of the bitwise rotate-left, by the given number of bits.
'@Remarks
Public Function RotateLeft(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
End Function

'@TODO Implement
'@Static
'@Description "Bitwise operatation shifts each bit to the right, with the lowest bit copied in the Carry flag and into the highest bit."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to rotate-right by.
'@Returns LongLong
'   Returns a Int64 of the bitwise rotate-right, by the given number of bits.
'@Remarks
Public Function RotateRight(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
End Function

'@TODO Implement
'@Static
'@Description "Bitwise operatation shifts each bit to the left, copies the Carry flag to the least significant bit and copies the most significant bit into the Carry flag."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to rotate carry left by.
'@Returns LongLong
'   Returns a Int64 of the bitwise rotate carray left, by the given number of bits.
'@Remarks
Public Function RotateCarryLeft(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
End Function

'@TODO Implement
'@Static
'@Description "Bitwise operatation shifts each bit to the right, copies the Carry flag to the most significant bit and copies the least significant bit into the Carry flag."
'@Parameters
'   val: LongLong value
'   bitsShifted: The number of bits to rotate carry right by.
'@Returns LongLong
'   Returns a Int64 of the bitwise rotate carray right, by the given number of bits.
'@Remarks
Public Function RotateCarryRight(ByVal val As LongLong, ByVal bitsShifted As Byte) As LongLong
End Function

'@Static
'@Description "Converts an Int64 to an binary string"
'@Parameters
'   val: LongLong value
'   byteGrouping: Boolean
'
'@Returns String
'   Returns a string representing an Int64 as bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
Public Function ToBinary(ByVal val As LongLong, Optional ByVal byteGrouping As Boolean = False) As String
    Const DELIMITER As String = " "
    Dim i
    For i = 0 To 63
        'if digit grouping, every 8 bits insert delimiter/space
        If byteGrouping Then
            If (i Mod 8 = 0) And i <> 0 Then
                ToBinary = ToBinary & DELIMITER
            End If
        End If
        
        If this.Powers(63 - i) And val Then
            ToBinary = ToBinary & "1"
        Else
            ToBinary = ToBinary & "0"
        End If
    Next
End Function

'------------------------------------------------------------------
'Private Methods
'------------------------------------------------------------------

'@Description "Initialises bitwise arrays"
Private Sub InitBitwiseArrays()
    this.Powers = InitPowers
    this.Masks = InitMasks
    this.RevMasks = InitRevMasks
End Sub

'2^i
Private Function InitPowers() As LongLong()
    Dim bitPosition As Long
    Dim pvtPowers(64) As LongLong
    
    For bitPosition = 0 To 62
        pvtPowers(bitPosition) = 2 ^ bitPosition
    Next bitPosition
    
    pvtPowers(63) = "&H8000000000000000" 'To avoid overflow calculation
    pvtPowers(64) = &H0
    
    InitPowers = pvtPowers
End Function

'2^i-1
Private Function InitMasks() As LongLong()
    Dim bitPosition As Long
    Dim pvtMasks(64) As LongLong
    
    For bitPosition = 0 To 63
        pvtMasks(bitPosition) = (2 ^ bitPosition) - 1
    Next bitPosition
    pvtMasks(64) = "&HFFFFFFFFFFFFFFFF"
    
    InitMasks = pvtMasks
End Function

'-2^i
Private Function InitRevMasks() As LongLong()
    Dim bitPosition As Long
    Dim pvtRevMasks(64) As LongLong
    
    For bitPosition = 0 To 63
        pvtRevMasks(bitPosition) = -2 ^ bitPosition
    Next bitPosition
    pvtRevMasks(64) = &H0
    
    InitRevMasks = pvtRevMasks
End Function

