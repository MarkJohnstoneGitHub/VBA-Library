Attribute VB_Name = "Convert"
Attribute VB_Description = "Converts a value to the required type."
'@ModuleDescription "Converts a value to the required type."
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
''MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

'@References
' Casts are unchecked by default
' https://github.com/dotnet/runtime/issues/30580
''

Option Explicit

''
'@Static
'@Description "Converts a value to unsigned 32-bit integer."
'@Parameters
'   val Variant
'       Value to be converted to an unsigned 32-bit value
'@Returns ULong
'   Value converted to an unsigned 32 bit value
'
'@Exceptions
'   ArgumentException
'       Raised for an invalid value which is not numeric.
'@Remarks
'   Casts are unchecked by default at runtime
'   CreateTruncating method is used when converting
''
Public Function CULong32(ByRef val As Variant) As ULong
    CULong32 = ULong32.CreateTruncating(val)
End Function
