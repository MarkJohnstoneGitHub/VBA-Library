Attribute VB_Name = "Testing_UInt32Static_ToString"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 6, 2023

Option Explicit

Private Sub TestingUInt32ToString()
    Dim lngVal  As Long
    Dim ulngResult As ULong
    
    lngVal = &HF6F2F1F0
    ulngResult = CBytesUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
End Sub

Private Sub TestingUInt32ToStringErrorOverflow()
    Dim lngVal  As Long
    Dim ulngResult As ULong
    
    lngVal = &HF6F2F1F0
    ulngResult = CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
End Sub
