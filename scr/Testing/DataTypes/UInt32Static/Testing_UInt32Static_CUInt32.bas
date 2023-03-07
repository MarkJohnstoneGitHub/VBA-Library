Attribute VB_Name = "Testing_UInt32Static_CUInt32"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 6, 2023
'@LastModified March 7, 2023

Option Explicit

Private Sub TestingUInt32StaticCUInt32()
    Dim lngVal  As Long
    Dim ulngResult As ULong
    
    ulngResult = CUInt32("&HF6F2F1F0")
    Debug.Print UInt32Static.ToString(ulngResult)

    lngVal = 1
    ulngResult = CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    lngVal = 342345
    ulngResult = CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    lngVal = 34
    ulngResult = CUInt32(lngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    ulngResult = CUInt32("&HFFFFFFFF")
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim strVal As String
    strVal = "4294967295"
    ulngResult = CUInt32(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "&HFFFFFFFF"
    ulngResult = CUInt32(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    #If Win64 Then
    Dim lnglngVal As LongLong
    lnglngVal = 4294967295#
    ulngResult = UInt32Static.CUInt32(lnglngVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    #End If
    
    Dim byteVal As Byte
    byteVal = 255
    ulngResult = UInt32Static.CUInt32(byteVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim intVal As Integer
    intVal = 255
    ulngResult = CUInt32(intVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    'Note -ve values are converted into large UInt32 values
    intVal = 23766
    ulngResult = CUInt32(intVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    intVal = 237
    ulngResult = CUInt32(intVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    ulngResult = CUInt32(0)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    ulngResult = CUInt32(0&)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    ulngResult = CUInt32(0@)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim curVal As Currency
    
    'Rounded down
    curVal = 1245.43@
    ulngResult = CUInt32(curVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    'Rounded up
    curVal = 1245.51@
    ulngResult = CUInt32(curVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    Dim dbVal As Double
    
    'Rounded down
    dbVal = 34325.5
    ulngResult = CUInt32(dbVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    'Rounded up
    dbVal = 34325.56
    ulngResult = CUInt32(dbVal)
    Debug.Print UInt32Static.ToString(ulngResult)
End Sub
