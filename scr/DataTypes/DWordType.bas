Attribute VB_Name = "DWordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

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
