Attribute VB_Name = "Testing_UInt32Static_Parse"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 28, 2023
'@LastModified February 28, 2023

Option Explicit

Private Sub TestingUInt32StaticParse()
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "0"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "0.45"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "0.5"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "0.51"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "4294967295"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "4294967294.95"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "    4294967295"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "    4294967295  "
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "    4294967294.95  "
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
End Sub

Private Sub TestingUInt32StaticParseInvalid()
On Error Resume Next
    Dim ulngResult As ULong
    Dim strVal As String
    
    strVal = "    4294967295.95  "
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "-1.21"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)

    strVal = "abc"
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = ""
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = "  "
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
    strVal = VBA.vbNullString
    ulngResult = UInt32Static.Parse(strVal)
    Debug.Print UInt32Static.ToString(ulngResult)
    
End Sub
