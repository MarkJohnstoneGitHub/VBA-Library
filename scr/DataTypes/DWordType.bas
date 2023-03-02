Attribute VB_Name = "DWordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 26, 2023
'@LastModified March 2, 2023

Option Explicit

Public Type DWORD
    wLow     As WORD     ' the ordering is important to remain consistant with memory layout.
    wHigh    As WORD
End Type

Public Type DWORD_VALUE
    Value As Long
End Type
