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
'   destination ULong
'       The destination which the source value is copied to.
'@Returns ULong
'   The unsigned 32-bit integer created from the value byte values.
'
'@Exceptions
'   OverflowException
'       Raised when a hex string or LongLong value exceeds the max unsigned 32-bit value of 4294967295
'   ArgumentException
'       Raised for when not a valid type to cast of a Byte, Integer, LongLong or string
'       containing a hex value i.e. "&H" preceeding the hex value within the string.
'
'@Remarks
'   Bytes values to be converted maybe of types Byte, Long, Integers or string containing a hex value.
'   Negative values are converted into a large unsigned 32-bit integer.
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
