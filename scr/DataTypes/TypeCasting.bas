Attribute VB_Name = "TypeCasting"
Attribute VB_Description = "Converts a value to the required type."
'@ModuleDescription "Converts a value to the required type."
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
''MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

''
'@Static
'@Description "Converts the byte value into an unsigned 32-bit integer."
'@Parameters
'   val Variant
'       The value to be converted to an unsigned 32-bit integer.
'       The byte value is copied to an unsigned 32-bit integer where negative values are converted
'       into larger unsigned values.
'@Returns ULong
'   The unsigned 32-bit integer created from the byte value.
'
'@Exceptions
'   ArgumentException
'       Raised for when not a valid type to cast of a Byte, Integer, Long, LongLong or Currency
'
'@Remarks
'   Bytes values to be converted maybe of types Byte, Long, Integers or Currency.
'   Negative values are converted into a larger unsigned 32-bit integers.
'
'   For the Byte type it is copied to the lower byte of the ULong DWORD.
'   eg. If byte value is 255 Hex FF is converted to the ULong value of 255, Hex 000000FF
'
'   For the Integer type its WORD value is copied to the ULong DWORD lower WORD.
'   eg. If Integer value is -1 Hex FFFF is converted to ULong value of 65535 Hex 0000FFFF
'
'   For the Long type its DWORD value is copied to the ULong DWORD
'   eg. If Long value is -1 Hex FFFFFFFF is converted to ULong value of 4294967295 Hex FFFFFFFF
'
'   For the Currency type the low DWORD of a currency value is copied to the DWORD of the ULong.
'   Eg Currency value of 0.0001 Hex 00000000 00000001 converts to ULong 00000001 i.e. of value 1.
'
'   For the LongLong type the low DWORD of a LongLong value is copied to the DWORD of the ULong.
'   eg. If LongLong value is 42949672958 i.e. Hex 00000009 FFFFFFFE is converted to ULong value Hex FEFFFFFF
'   i.e. value of 4294967294
''
Public Function CBytesUInt32(ByVal val As Variant) As ULong
    CBytesUInt32 = UInt32Static.CBytesUInt32(val)
End Function

''
'@Static
'@Description "Converts a value to unsigned 32-bit integer."
'@Parameters
'   val: Variant
'       value to be converted to an unsigned 32-bit value
'@Returns ULong
'   value converted to an unsigned 32 bit value
'
'@Exceptions
'   OverflowException
'       Raised when a value is less then 0 or exceeds the max unsigned 32-bit value of 4294967295
'   ArgumentException
'       Raised for an invalid value which is not numeric.
'@Remarks
'   Decimal values are truncated
'   Negative values return an overflow exception
''
Public Function CUInt32(ByRef val As Variant) As ULong
    CUInt32 = UInt32Static.CUInt32(val)
End Function
