Attribute VB_Name = "ULongLongType"
'@Folder("VBACorLib.DataTypes")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 March 2, 2023
'@LastModified March 2, 2023

Option Explicit

#If Win64 Then
    Public Type ULongLong
        Value As LongLong
    End Type
#Else
    Public Type ULongLong
        LowPart   As Long
        HighPart  As Long
    End Type
#End If

Public Type ULongLongLowHigh
    LowPart   As Long
    HighPart  As Long
End Type



