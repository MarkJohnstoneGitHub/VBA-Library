Attribute VB_Name = "Convert"
Attribute VB_Description = "Converts a value to the required type."
'@ModuleDescription "Converts a value to the required type."
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
''MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 6, 2023
'@LastModified March 9, 2023

'@References
' Casts are unchecked by default
' https://github.com/dotnet/runtime/issues/30580
'
'

Option Explicit

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
'   Casts are unchecked by default at runtime
'   CreateTruncating method is used when converting
''
Public Function CULong32(ByRef val As Variant) As ULong
    CULong32 = ULong32.CreateTruncating(val)
End Function
