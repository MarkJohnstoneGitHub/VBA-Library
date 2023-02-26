Attribute VB_Name = "Testing_UInt32Static_CUInt32"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 26, 2023
'@LastModified February 26, 2023

Option Explicit

Private Sub TestingUInt32StaticCUInt32()
    Dim lngVal  As Long
    Dim ulngResult As ULong
    
    lngVal = &HF6F2F1F0
    ulngResult = UInt32Static.CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)

    'Note -ve values are converted into large UInt32 values
    lngVal = -1
    ulngResult = UInt32Static.CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    lngVal = -342345
    ulngResult = UInt32Static.CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    lngVal = 342345
    ulngResult = UInt32Static.CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    lngVal = &HFFFFFFFF
    ulngResult = UInt32Static.CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim strVal As String
    strVal = "4294967295"
    ulngResult = UInt32Static.CUInt32(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim lnglngVal As LongLong
    lnglngVal = 4294967295#
    ulngResult = UInt32Static.CUInt32(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim byteVal As Byte
    byteVal = 255
    ulngResult = UInt32Static.CUInt32(byteVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim intVal As Integer
    intVal = 255
    ulngResult = UInt32Static.CUInt32(intVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    'Note -ve values are converted into large UInt32 values
    intVal = -23766
    ulngResult = UInt32Static.CUInt32(intVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    intVal = 23766
    ulngResult = UInt32Static.CUInt32(intVal)
    Debug.Print UInt32Static.ToString(ulngResult)
End Sub
