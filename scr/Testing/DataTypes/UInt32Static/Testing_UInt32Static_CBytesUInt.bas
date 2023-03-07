Attribute VB_Name = "Testing_UInt32Static_CBytesUInt"
'@Folder "Testing.VBACorLib.DataTypes.UInt32Static"

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.2 March 7, 2023
'@LastModified March 7, 2023

Private Sub TestingInt32StaticCBytesUInt_Currency()
    Dim val As Currency
    Dim ulngResult As ULong
    val = 0.0001@                                   'Hex 00000000 00000001
    ulngResult = UInt32Static.CBytesUInt32(val)     'Hex 00000001
    Debug.Print UInt32Static.ToString(ulngResult)   '1
End Sub

#If Win64 Then
Private Sub TestingInt32StaticCBytesUInt_LongLong()
    Dim val As LongLong
    Dim ulngResult As ULong
    val = 42949672958#                              'Hex 00000009 FFFFFFFE
    ulngResult = UInt32Static.CBytesUInt32(val)     'Hex FEFFFFFF
    Debug.Print UInt32Static.ToString(ulngResult)   '4294967294
End Sub
#End If

Private Sub TestingInt32StaticCBytesUInt_Long()
    Dim val As Long
    Dim ulngResult As ULong
    val = &HFFFFFFFF                                'Hex FFFFFFFF  Value - 1
    ulngResult = UInt32Static.CBytesUInt32(val)     'Hex FFFFFFFF
    Debug.Print UInt32Static.ToString(ulngResult)   '4294967295
End Sub

Private Sub TestingInt32StaticCBytesUInt_Integer()
    Dim val As Integer
    Dim ulngResult As ULong
    val = -1                                        'Hex FFFFF  Value - 1
    ulngResult = UInt32Static.CBytesUInt32(val)     'Hex 0000FFFFF
    Debug.Print UInt32Static.ToString(ulngResult)   '65535
End Sub

Private Sub TestingInt32StaticCBytesUInt_Byte()
    Dim val As Byte
    Dim ulngResult As ULong
    val = 255                                       'Hex FF  Value 255
    ulngResult = UInt32Static.CBytesUInt32(val)     'Hex 000000FF
    Debug.Print UInt32Static.ToString(ulngResult)   '255
End Sub

