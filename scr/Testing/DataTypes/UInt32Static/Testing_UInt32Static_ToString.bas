Attribute VB_Name = "Testing_UInt32Static_ToString"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 March 4, 2023
'@LastModified March 4, 2023

Option Explicit

Private Sub TestingUInt32ToString()
    Dim lngVal  As Long
    Dim ulngResult As ULong
    
    lngVal = &HF6F2F1F0
    ulngResult = UInt32Static.CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
End Sub
