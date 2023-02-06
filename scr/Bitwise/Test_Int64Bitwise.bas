Attribute VB_Name = "Test_Int64Bitwise"
'@Folder("VBACorLib.DataTypes")

'@Version v1.0 (Date January 20, 2023)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/

'@Author markjohnstone@hotmail.com
'@LastModified January 30,2023
'
'@Remarks
' Bitwise operations for the Int64 data type.
Option Explicit

Private Sub TestBitshift()
    Dim int64Value As LongLong
    Dim binary As String
    Dim result As LongLong

    Dim numbits As Byte
    
    numbits = 1
    int64Value = "4353245634767778845"
    
    binary = Int64Bitwise.ToBinary(int64Value, True)
    Debug.Print "Input           :  "; int64Value, binary
    
    result = Int64Bitwise.ShiftRight(int64Value, numbits)
    binary = Int64Bitwise.ToBinary(result, True)
    Debug.Print "Shift right bits: " & numbits; result, binary
    
    result = Int64Bitwise.ShiftLeft(result, numbits)
    binary = Int64Bitwise.ToBinary(result, True)
    Debug.Print "Shift left bits : " & numbits; result, binary
    Debug.Print
    
    
    int64Value = 43534534789#
    
    
    binary = Int64Bitwise.ToBinary(int64Value, True)
    Debug.Print "Input           :  "; int64Value, binary
    
    result = Int64Bitwise.ShiftLeft(int64Value, numbits)
    binary = Int64Bitwise.ToBinary(result, True)
    Debug.Print "Shift right bits: " & numbits; result, binary
    
    result = Int64Bitwise.ShiftRight(result, numbits)
    binary = Int64Bitwise.ToBinary(result, True)
    Debug.Print "Shift left bits : " & numbits; result, binary
    Debug.Print
End Sub
