Attribute VB_Name = "WordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.1 March 4, 2023
'@LastModified March 4, 2023

Option Explicit

Public Type WORD
    Value    As Integer
End Type

Public Type WORDLoHi
    LowPart    As Byte     ' the ordering is important to remain consistant with memory layout.
    HighPart   As Byte
End Type


