Attribute VB_Name = "QWordType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 27, 2023
'@LastModified February 27, 2023

Option Explicit

#If VBA7 Then
    Public Type QWORD
        Value As LongLong
    End Type
#Else
    Public Type QWORD
        LowPart     As Long     ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
        HighPart    As Long
    End Type
#End If
