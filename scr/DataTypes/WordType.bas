Attribute VB_Name = "WordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Public Type WORD
    Value    As Integer
End Type

Public Type WORDLoHi
    LowPart    As Byte     ' the ordering is important to remain consistant with memory layout.
    HighPart   As Byte
End Type


