Attribute VB_Name = "QWordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 27, 2023
'@LastModified March 2, 2023

Option Explicit

#If Win64 Then
    Public Type QWORD_VALUE
        Value As LongLong
    End Type
#Else
    Public Type QWORD_VALUE
        LowPart     As Long     ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
        HighPart    As Long
    End Type
#End If

Public Type QWORD
    dwLow   As DWORD    ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
    dwHigh  As DWORD
End Type

