Attribute VB_Name = "DWordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.1 March 4, 2023
'@LastModified March 4, 2023

Option Explicit

Public Type DWORD
    Value   As Long
End Type

Public Type DWORDwLoHi
    wLow    As WORDLoHi         ' the ordering is important to remain consistant with memory layout.
    wHigh   As WORDLoHi
End Type

Public Type DWORDLoHi
    wLow     As WORD     ' the ordering is important to remain consistant with memory layout.
    wHigh    As WORD
End Type
