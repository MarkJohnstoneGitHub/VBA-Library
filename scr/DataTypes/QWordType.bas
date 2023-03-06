Attribute VB_Name = "QWordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

#If Win64 Then
    Public Type QWORD
        Value As LongLong
    End Type
#Else
    Public Type QWORD
        LowPart     As Long     ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
        HighPart    As Long
    End Type
#End If

Public Type QWORDLoHi
    dwLow   As DWORDwLoHi    ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
    dwHigh  As DWORDwLoHi
End Type

'@TODO rename WIP for UInt64
Public Type QWORDV2
    LowPart     As Long     ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
    HighPart    As Long
End Type


