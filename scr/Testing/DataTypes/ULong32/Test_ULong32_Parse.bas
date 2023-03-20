Attribute VB_Name = "Test_ULong32_Parse"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingULong32Parse()
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "0"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    
    strVal = "4294967295"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
        
    strVal = "    4294967295"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "    4294967295  "
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "&H0"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "&HFF"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "&HFFFFFFFE"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "&HFFFFFFFF"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "&O37777777777"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
End Sub

Private Sub TestingULong32ParseInvalid()
On Error Resume Next
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "    4294967295.95  "
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "-1.21"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)

    strVal = "abc"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = ""
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "  "
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = VBA.vbNullString
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
End Sub

Private Sub TestingULong32ParseArgumentNullException()
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = VBA.vbNullString
    On Error GoTo ErrorHandler
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
ErrorHandler:
    Debug.Print Err.Number, Err.source, Err.Description
End Sub

Private Sub TestingULong32ParseHexStringOverflow()
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "&HFFFFFFFFF"
    On Error GoTo ErrorHandler
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
ErrorHandler:
    Debug.Print Err.Number, Err.source, Err.Description
End Sub

Private Sub TestingULong32ParseErrorInvalid()
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "&FFFFFFFFF"
    On Error GoTo ErrorHandler
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
ErrorHandler:
    Debug.Print Err.Number, Err.source, Err.Description
End Sub

Private Sub TestingULong32ParseErrorArgumentException()
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "10.45"
    On Error GoTo ErrorHandler
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
ErrorHandler:
    Debug.Print Err.Number, Err.source, Err.Description
End Sub


