VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitwiseInt64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBACorLib.DataTypes")
'@PredeclaredId

'@Singleton Class

'@Version v1.0 (Date January 20, 2023)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/

'@Author markjohnstone@hotmail.com
'@LastModified January 30,2023
'
'@Remarks
' Bitwise operations for the Int64 data type.

'http://something4skiddie.blogspot.com/2012/11/bit-shift.html
'https://www.wireshark.org/docs/wsdg_html_chunked/lua_module_Int64.html
'https://www.excely.com/excel-vba/bit-shifting-function.shtml

Option Explicit

'------------------------------------------------------------------
'Constants
'------------------------------------------------------------------

'------------------------------------------------------------------
'Types
'------------------------------------------------------------------
Private Type TInt64Static
    Powers() As LongLong    'Used for bitwise operations
    Masks() As LongLong     'Used for bitwise operations
    RevMasks() As LongLong  'Used for bitwise operations
End Type

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As TInt64Static

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
'   numbits: The number of bits to right-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise logical right-shift operation, by the given number of bits.
'@Remarks
'   Right shifting is equal to dividing value by 2^Shift.
Public Function ShiftRight(ByVal val As LongLong, ByVal numbits As Byte) As LongLong
    Const MIN_BITSHIFT As Long = 0
    Const MAX_BITSHIFT As Long = 64

    If numbits > MIN_BITSHIFT And numbits < MAX_BITSHIFT Then
        ShiftRight = (val And this.RevMasks(numbits)) \ _
                                 this.Powers(numbits) And this.Masks(MAX_BITSHIFT - numbits)
    ElseIf numbits <= MIN_BITSHIFT Then
        ShiftRight = val
    ElseIf numbits >= MAX_BITSHIFT Then
        ShiftRight = &H0&
    End If
End Function

'@Static
'@Description "Bitwise arithmetic right-shift operation, by the given number of bits."
'@Parameters
'   val: LongLong value
'   numbits: The number of bits to right-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise arithmetic right-shift operation, by the given number of bits.
'@Return: LongLong value
Public Function ShiftRightArithmetic(ByVal val As LongLong, ByVal numbits As Byte) As LongLong
    Const MIN_BITSHIFT As Long = 0
    Const MAX_BITSHIFT As Long = 64 - 1
    Const LAST_BIT_MASK As LongLong = "&H8000000000000000"

    If numbits > 0 And numbits < MAX_BITSHIFT Then
        ShiftRightArithmetic = (val And this.RevMasks(numbits)) \ _
                                    this.Powers(numbits)
'    ElseIf numbits <= 0 Then
'        ShiftRightArithmetic = val
    ElseIf numbits >= MAX_BITSHIFT Then
        If val And LAST_BIT_MASK Then
            ShiftRightArithmetic = "&HFFFFFFFFFFFFFFFF"
        Else
            ShiftRightArithmetic = &H0&
        End If
    End If
End Function

'@Static
'@Description "Bitwise logical left-shift operation, by the given number of bits."
'@Parameters
'   val: LongLong value
'   numbits: The number of bits to left-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise logical left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
Public Function ShiftLeft(ByVal val As LongLong, ByVal numbits As Byte) As LongLong
    Const MIN_BITSHIFT As Byte = 0
    Const MAX_BITSHIFT As Byte = 64
    Const LAST_BIT_MASK As LongLong = "&H8000000000000000"

    If numbits > MIN_BITSHIFT And numbits < MAX_BITSHIFT Then
        ShiftLeft = (val And this.Masks(63 - numbits)) * _
                                this.Powers(numbits)
        If val And this.Powers(63 - numbits) Then
            ShiftLeft = ShiftLeft Or LAST_BIT_MASK ' set bit 63
        End If
    ElseIf numbits <= 0 Then
        ShiftLeft = val
    ElseIf numbits >= MAX_BITSHIFT Then
        ShiftLeft = &H0&
    End If
End Function


'@Static
'@Description "Bitwise arithmetic left-shift operation, by the given number of bits."
'@Parameters
'   val: LongLong value
'   numbits: The number of bits to left-shift by.
'@Returns LongLong
'   Returns a Int64 of the bitwise arithmetic left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
Public Function ShiftLeftArithmetic(ByVal val As LongLong, ByVal numbits As Byte) As LongLong
    Const MIN_BITSHIFT As Long = 0
    Const MAX_BITSHIFT As Long = 64 - 1
    Const LAST_BIT_MASK As LongLong = "&H8000000000000000"

    If numbits > MIN_BITSHIFT And numbits < MAX_BITSHIFT Then
        If val And LAST_BIT_MASK Then
            val = val Or this.RevMasks(MAX_BITSHIFT - numbits)
        Else
            val = val And this.Masks(MAX_BITSHIFT - numbits)
        End If
        ShiftLeftArithmetic = val * this.Powers(numbits)
    ElseIf numbits = 0 Then
        ShiftLeftArithmetic = val
    ElseIf numbits >= MAX_BITSHIFT Then
        'To avoid an overflow error
        If val And LAST_BIT_MASK Then
            ShiftLeftArithmetic = LAST_BIT_MASK ' set bit 63
        Else
            ShiftLeftArithmetic = &H0&
        End If
    End If
    
End Function

'@Static
'@Description "Converts an Int64 to an binary string"
'@Parameters
'   val: LongLong value
'   digitGrouping: Boolean
'
'@Returns LongLong
'   Returns a Int64 of the bitwise arithmetic left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
Public Function ToBinary(ByVal val As LongLong, Optional ByVal digitGrouping As Boolean = False) As String
    Const DELIMITER As String = " "
    Dim i
    For i = 0 To 63
        'if digit grouping, every 8 bits insert delimiter/space
        If digitGrouping Then
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
    Dim RevMasks(64) As LongLong
    
    For bitPosition = 0 To 63
        RevMasks(bitPosition) = -2 ^ bitPosition
    Next bitPosition
    RevMasks(64) = &H0
    
    InitRevMasks = RevMasks
End Function

