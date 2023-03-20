Attribute VB_Name = "Test_ULong32_CreateChecked"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingULong32CreateChecked()
    Dim lngVal  As Long
    Dim ulngResult As ULong
    
    lngVal = 1
    ulngResult = ULong32.CreateChecked(lngVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    lngVal = 342345
    ulngResult = ULong32.CreateChecked(lngVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    lngVal = 34
    ulngResult = ULong32.CreateChecked(lngVal)
    Debug.Print ULong32.ToString(ulngResult)
        
    #If Win64 Then
    Dim lnglngVal As LongLong
    lnglngVal = 4294967295^
    ulngResult = ULong32.CreateChecked(lnglngVal)
    Debug.Print ULong32.ToString(ulngResult)
    #End If
    
    Dim byteVal As Byte
    byteVal = 255
    ulngResult = ULong32.CreateChecked(byteVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    Dim intVal As Integer
    intVal = 255
    ulngResult = ULong32.CreateChecked(intVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    'Note -ve values are converted into large UInt32 values
    intVal = 23766
    ulngResult = ULong32.CreateChecked(intVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    intVal = 237
    ulngResult = ULong32.CreateChecked(intVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    ulngResult = ULong32.CreateChecked(0)
    Debug.Print ULong32.ToString(ulngResult)
    
    ulngResult = ULong32.CreateChecked(0&)
    Debug.Print ULong32.ToString(ulngResult)
    
    ulngResult = ULong32.CreateChecked(0@)
    Debug.Print ULong32.ToString(ulngResult)
    
    Dim curVal As Currency
    
    'Rounded down
    curVal = 1245.43@
    ulngResult = ULong32.CreateChecked(curVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    
    'Rounded up
    curVal = 1245.51@
    ulngResult = ULong32.CreateChecked(curVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    Dim dbVal As Double
    'Rounded down
    dbVal = 34325.49
    ulngResult = ULong32.CreateChecked(dbVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    'Rounded down
    dbVal = 34325.5
    ulngResult = ULong32.CreateChecked(dbVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    'Rounded up
    dbVal = 34325.56
    ulngResult = ULong32.CreateChecked(dbVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    Dim sngVal As Single
    'Rounded down
    sngVal = 34325.49
    ulngResult = ULong32.CreateChecked(sngVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    'Rounded down
    sngVal = 34325.5
    ulngResult = ULong32.CreateChecked(sngVal)
    Debug.Print ULong32.ToString(ulngResult)

    'Rounded up
    sngVal = 34325.56
    ulngResult = ULong32.CreateChecked(sngVal)
    Debug.Print ULong32.ToString(ulngResult)
End Sub


Private Sub TestingULong32CreateCheckedErrorOverflow()
    Dim lngVal  As Long
    Dim ulngResult As ULong
    
    lngVal = &HF6F2F1F0
    On Error GoTo ErrorHandler
    ulngResult = ULong32.CreateChecked(lngVal)
    Debug.Print ULong32.ToString(ulngResult)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number, Err.source, Err.Description
End Sub

Private Sub TestingULong32CreateCheckedErrorArgumentException()
    Dim ulngResult As ULong
    On Error GoTo ErrorHandler
    ulngResult = ULong32.CreateChecked("&HF6F2F1F0")
    Debug.Print ULong32.ToString(ulngResult)
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number, Err.source, Err.Description
End Sub
