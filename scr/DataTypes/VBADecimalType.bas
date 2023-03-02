Attribute VB_Name = "VBADecimalType"
Attribute VB_Description = "VBA Decimal type structure within a variant."
'@Folder("VBACorLib.DataTypes")
'@ModuleDescription "VBA Decimal type structure within a variant."

'The DECIMAL structure specifies a sign and scale for a number. Decimal variables are represented
'as 96-bit unsigned integers that are scaled by a variable power of 10.

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 25, 2023
'@LastModified March 2, 2023

'@References
' https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/b5493025-e447-4109-93a8-ac29c48d018d
' https://www.vbforums.com/showthread.php?834827-The-Decimal-Data-Type
' https://newtonexcelbach.com/2015/10/26/the-vba-decimal-data-type/
' https://stackoverflow.com/questions/59899919/how-does-the-caller-know-when-theres-a-decimal-inside-a-variant

Option Explicit

Public Type DecimalType     ' (when sitting in a Variant)
    vt           As Integer ' Reserved, to act as the variable Type when sitting in a 16-Byte-Variant.  Equals vbDecimal(14) when it's a Decimal type.
    Scale        As Byte    ' Base 10 exponent (0 to 28), moving decimal to right (smaller numbers) as this value goes higher.  Top three bits are never used.
    Sign         As Byte    ' Sign bit only (high bit).  Other bits aren't used.
    Hi32         As DWORD   ' Mantissa.
    Lo32         As DWORD   ' Mantissa.
    Mid32        As DWORD   ' Mantissa.
End Type

'@References
' https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oaut/b5493025-e447-4109-93a8-ac29c48d018d
Public Type tagDEC         ' (when sitting in a Variant)
    vt      As Integer     ' Reserved, to act as the variable Type when sitting in a 16-Byte-Variant.  Equals vbDecimal(14) when it's a Decimal type.
    Scale   As Byte        ' Base 10 exponent (0 to 28), moving decimal to right (smaller numbers) as this value goes higher.  Top three bits are never used.
    Sign    As Byte        ' Sign bit only (high bit).  Other bits aren't used.
    Hi32    As ULong       ' Mantissa.
    Lo64    As ULongLong   ' Mantissa.
End Type


