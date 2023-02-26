Attribute VB_Name = "Testing_UInt32Static_Addition"
'@Folder("Testing.VBACorLib.DataTypes.UInt32Static")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 26, 2023
'@LastModified February 26, 2023

Option Explicit

Private Sub TestingUInt32Add()
    Dim result As ULong

    Dim t1  As ULong
    Dim t2 As ULong

    t1.Value = &HF6F2F1F0
    t2.Value = &H1F3&

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)

    t1.Value = &HFF2F1F
    t2.Value = &H1F364

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)
    
    t1.Value = &HFF2F1F
    t2.Value = &H0&

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)
    
    t1.Value = &H0
    t2.Value = &HFF2F1F

    result = UInt32Static.Add(t1, t2)
    Debug.Print UInt32Static.ToString(t1) & " + " & UInt32Static.ToString(t2) & " = " & UInt32Static.ToString(result)
End Sub

