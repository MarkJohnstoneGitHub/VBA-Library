Attribute VB_Name = "Test_BitShift"
'@Folder("VBALib.DataTypes")

'@Version v1.0 (Date January 20, 2023)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/

'@Author markjohnstone@hotmail.com
'@LastModified January 30,2023
'
'@Remarks
' Bitwise operations for the Int64 data type.
Option Explicit

Private Sub TestBitshift()
    Dim inputInt64 As LongLong
    Dim output As String
    Dim result As LongLong

    Dim numbits As Byte
    
    numbits = 1
    inputInt64 = "4353245634767778845"
    
    output = Int64Static.ToBinary(inputInt64, True)
    Debug.Print "Input           :  "; inputInt64, output
    
    result = Int64Static.ShiftRight(inputInt64, numbits)
    output = Int64Static.ToBinary(result, True)
    Debug.Print "Shift right bits: " & numbits; result, output
    
    result = Int64Static.ShiftLeft(result, numbits)
    output = Int64Static.ToBinary(result, True)
    Debug.Print "Shift left bits : " & numbits; result, output
    Debug.Print
    
    
    inputInt64 = 43534534789#
    
    
    output = Int64Static.ToBinary(inputInt64, True)
    Debug.Print "Input           :  "; inputInt64, output
    
    result = Int64Static.ShiftLeft(inputInt64, numbits)
    output = Int64Static.ToBinary(result, True)
    Debug.Print "Shift right bits: " & numbits; result, output
    
    result = Int64Static.ShiftRight(result, numbits)
    output = Int64Static.ToBinary(result, True)
    Debug.Print "Shift left bits : " & numbits; result, output
    Debug.Print
End Sub
