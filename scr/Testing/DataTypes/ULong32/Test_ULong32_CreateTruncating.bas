Attribute VB_Name = "Test_ULong32_CreateTruncating"
'@Folder "Testing.VBACorLib.DataTypes.ULong32"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 21, 2023

Option Explicit

Private Sub TestingULong32CreateTruncating_Currency()
    Dim val As Currency
    Dim ulngResult As ULong
    val = 0.0001@                                   'Hex 00000000 00000001
    ulngResult = ULong32.CreateTruncating(val)      'Hex 00000001
    Debug.Print ULong32.ToString(ulngResult)        '1
End Sub

#If Win64 Then
Private Sub TestingULong32CreateTruncating_LongLong()
    Dim val As LongLong
    Dim ulngResult As ULong
    val = 42949672958#                              'Hex 00000009 FFFFFFFE
    ulngResult = ULong32.CreateTruncating(val)      'Hex FEFFFFFF
    Debug.Print ULong32.ToString(ulngResult)        '4294967294
End Sub
#End If

Private Sub TestingULong32CreateTruncating_Long()
    Dim val As Long
    Dim ulngResult As ULong
    val = &HFFFFFFFF                                'Hex FFFFFFFF  Value - 1
    ulngResult = ULong32.CreateTruncating(val)      'Hex FFFFFFFF
    Debug.Print ULong32.ToString(ulngResult)        '4294967295
End Sub

Private Sub TestingULong32CreateTruncating_Integer()
    Dim val As Integer
    Dim ulngResult As ULong
    val = -1                                        'Hex FFFFF  Value - 1
    ulngResult = ULong32.CreateTruncating(val)      'Hex 0000FFFFF
    Debug.Print ULong32.ToString(ulngResult)        '65535
End Sub

Private Sub TestingULong32CreateTruncating_Byte()
    Dim val As Byte
    Dim ulngResult As ULong
    val = 255                                       'Hex FF  Value 255
    ulngResult = ULong32.CreateTruncating(val)      'Hex 000000FF
    Debug.Print ULong32.ToString(ulngResult)        '255
End Sub

