Attribute VB_Name = "WordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 26, 2023
'@LastModified February 26, 2023

Option Explicit

Public Type WORD
    LowByte     As Byte     ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
    HighByte    As Byte
End Type
