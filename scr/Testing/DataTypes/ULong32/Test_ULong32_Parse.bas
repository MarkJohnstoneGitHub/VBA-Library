Attribute VB_Name = "Test_ULong32_Parse"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.3 March 9, 2023
'@LastModified March 9, 2023

Option Explicit

Private Sub TestingULong32Parse()
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "0"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "0.45"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "0.5"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "0.51"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "4294967295"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "4294967294.95"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "    4294967295"
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "    4294967295  "
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "    4294967294.95  "
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
    
    strVal = "&HFFFFFFFE"
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
    ulngResult = ULong32.Parse(strVal)
    Debug.Print ULong32.ToString(ulngResult)
End Sub

