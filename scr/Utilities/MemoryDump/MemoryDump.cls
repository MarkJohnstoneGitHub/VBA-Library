VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MemoryDump"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleDescription "Memory dump utility"
'@Folder("Utilities.MemoryDump")
'@PredeclaredId

'@Singleton

''
'@Version  v1.0 MemoryDump (Date February 07, 2023)
'(c) Mark Johnstone - https://github.com/MarkJohnstoneGitHub/
'@Author markjohnstone@hotmail.com
'@LastModified February 07, 2023
''

'@TODO Error handling for parameters

Option Explicit

#If Mac Then
  #If Win64 Then
    Private Declare PtrSafe Function CopyMemoryByPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As LongLong) _
             As LongPtr
  #Else
    Private Declare Function CopyMemoryByPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As Long, _
             ByVal src As Long, _
             ByVal size As Long) _
             As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Private Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As LongLong)
  #Else
    Private Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As Long)
  #End If
#Else
  Private Declare Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
          (ByVal destination As Long, _
           ByVal src As Long, _
           ByVal size As Long)
#End If

'------------------------------------------------------------------
'Types
'------------------------------------------------------------------
Public Enum MemoryDumpFormat
    [_First] = 1
    Hexadecimal = 1
    Dec = 2
    ASCII = 3
    Binary = 4
    [_Last] = 4
End Enum

Private Type TMemoryDump
    Powers() As Byte    'Used for bitwise operations
End Type

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As TMemoryDump

'------------------------------------------------------------------
'Constructors and destructors
'------------------------------------------------------------------
Private Sub Class_Initialize()
    InitPowers
End Sub


'------------------------------------------------------------------
'Public Methods
'------------------------------------------------------------------

''
'@Static
'@Description "Returns a byte array for the given starting address and byte length."
'@Parameters
'   value Byte
'       Byte value to be converted into a ASCII character
'@Returns String
'   Returns a string representing a ASCII character.
''
#If VBA7 Then
Public Function ToBytes(ByVal pStartAddress As LongPtr, ByVal byteLength As Long) As Byte()
#Else
Public Function ToBytes(ByVal pStartAddress As Long, ByVal byteLength As Long) As Byte()
#End If
    Dim output() As Byte
    ReDim output(byteLength - 1)
     
    Dim i As Long
    For i = 0 To byteLength - 1
        CopyMemoryByPtr ByVal VarPtr(output(i)), ByVal (pStartAddress + i), 1
    Next i
    ToBytes = output
End Function

''
'@Static
'@Description "Converts an byte to an ASCII character."
'@Parameters
'   value Byte
'       Byte value to be converted into a ASCII character
'@Returns String
'   Returns a string representing a ASCII character.
''
Public Function ToASCII(ByVal value As Byte) As String
    If (value >= 32) And (value <= 255) Then
        ToASCII = Chr(value)
    Else
        ToASCII = "."
    End If
End Function

'@Static
'@Description "Converts an Byte to an binary string"
'@Parameters
'   value Byte
'       Byte value to be converted into a binary string
'@Returns String
'   Returns a string representing a Byte as bits.
Public Function ToBinary(ByVal value As Byte) As String
    Dim i As Long
    For i = 7 To 0 Step -1
        If this.Powers(i) And value Then
            ToBinary = ToBinary & "1"
        Else
            ToBinary = ToBinary & "0"
        End If
    Next
End Function

''
'@Static
'@Description "Converts an Byte to an decimal string with space padding"
'@Parameters
'   value Byte
'       Byte value to be converted into a decimal string
'   padding Boolean
'       Optional space padding for three character decimaal string.
'@Returns String
'   Returns a string representing a Byte as three character decimal string
''
Public Function ToDecimal(ByVal value As Byte, Optional ByVal padding As Boolean = True) As String
    If padding Then
        ToDecimal = VBA.Right$("   " & CStr(value), 3)
    Else
        ToDecimal = CStr(value)
    End If
End Function

''
'@Static
'@Description "Converts an Byte to an hex string with zero padding"
'@Parameters
'   value Byte
'       Byte value to be converted into a hex string
'   padding Boolean
'       Optional zero padding for two character hex string.
'@Returns String
'   Returns a string representing a Byte as two character hex string
''
Public Function ToHexidecimal(ByVal value As Byte, Optional ByVal padding As Boolean = True) As String
    If padding Then
        ToHexidecimal = VBA.Right$("00" & VBA.Hex(value), 2)
    Else
        ToHexidecimal = VBA.Hex(value)
    End If
End Function

''
'@Static
'@Description "Returns a string memory dump for the selected format from a starting address to byte length."
'@Parameters
'   pStartAddress LongPtr
'       Memory Address to begin memory dump
'   byteLength Long
'       Number of bytes for memory dump
'   outputFormat MemoryDumpFormat
'       Format to display i.e.  Decimal, hexadecimal,  ASCII or binary
'   bytesPerLine Long
'       Number of bytes per line displayed. Default is 8 bytes.
'@Returns String
'   Returns a string memory dump for the provided starting address and byte length for the required format.
''
#If VBA7 Then
Public Function ToString( _
                     ByVal pStartAddress As LongPtr, _
                     ByVal byteLength As Long, _
                     ByVal outputFormat As MemoryDumpFormat, _
            Optional ByVal bytesPerLine As Long = 8) _
                     As String
#Else
Public Function ToString( _
                     ByVal pStartAddress As Long, _
                     ByVal byteLength As Long, _
                     ByVal outputFormat As MemoryDumpFormat, _
            Optional ByVal bytesPerLine As Long = 8) _
                     As String

#End If
    'Validate Arguments (Positive byteLength and bytesPerLine, valid  outputFormat)
    'Raise error or exit. Preferred behaviour catch errors and exit function?
    
    Dim mem() As Byte
    mem = Me.ToBytes(pStartAddress, byteLength)
    
    Dim output As String
    'Table column headings
    output = output & "Address      "
    Select Case outputFormat
        Case MemoryDumpFormat.Binary
            output = output & vbTab & "Binary"
        Case MemoryDumpFormat.Dec
            output = output & vbTab & "Decimal"
        Case MemoryDumpFormat.Hexadecimal
            output = output & vbTab & "Hexadecimal"
        Case MemoryDumpFormat.ASCII
            output = output & vbTab & "ASCII"
        Case Else
            'Invalid argument outputType
    End Select
    
    #If VBA7 Then
    Dim memoryAddress As LongPtr
    #Else
    Dim memoryAddress As Long
    #End If
    
    memoryAddress = pStartAddress
    Dim bytesPerLineCounter As Long
    bytesPerLineCounter = 0
    Dim i As Long
    i = 0
    'output = output & mem(i, MEMORY_ADDRESS)
    For i = 0 To UBound(mem)
        If bytesPerLineCounter = 0 Or bytesPerLineCounter = bytesPerLine Then
            bytesPerLineCounter = 0
            output = output & vbNewLine & memoryAddress + i & vbTab
        End If
        
        Select Case outputFormat
            Case MemoryDumpFormat.Binary
                output = output & Me.ToBinary(mem(i)) & " "
            Case MemoryDumpFormat.Dec
                output = output & Me.ToDecimal(mem(i)) & " "
            Case MemoryDumpFormat.Hexadecimal
                output = output & Me.ToHexidecimal(mem(i)) & " "
            Case MemoryDumpFormat.ASCII
                output = output & ToASCII(mem(i))
            Case Else
                'Invalid argument outputType
        End Select
        bytesPerLineCounter = bytesPerLineCounter + 1
    Next i
    ToString = output
End Function

''
'@Static
'@Description "Returns a string memory dump for all formats from a starting address to byte length. "
'@Parameters
'   pStartAddress LongPtr
'       Memory Address to begin memory dump
'   byteLength Long
'       Number of bytes for memory dump
'   bytesPerLine Long
'       Number of bytes per line displayed. Default is 8 bytes.
'@Returns String
'   Returns a string memory dump for the provided starting address and byte length.
''
#If VBA7 Then
Public Function ToStringAll( _
                     ByVal pStartAddress As LongPtr, _
                     ByVal byteLength As Long, _
            Optional ByVal bytesPerLine As Long = 8) _
                     As String
#Else
Public Function ToStringAll( _
                     ByVal pStartAddress As Long, _
                     ByVal byteLength As Long, _
            Optional ByVal bytesPerLine As Long = 8) _
                     As String
    
#End If
    '@TODO Validate Arguments (Positive byteLength and bytesPerLine, valid  outputFormat)
    'Raise error or exit. Preferred behaviour catch errors and exit function?
    
    Dim output As String
    Dim enumVal As MemoryDumpFormat
    For enumVal = MemoryDumpFormat.[_First] To MemoryDumpFormat.[_Last]
        output = output & Me.ToString(pStartAddress, byteLength, enumVal, bytesPerLine) & vbNewLine & vbNewLine
    Next enumVal
    
    ToStringAll = output
End Function

''
'@Static
'@Description "Returns a string memory dump for all formats from a starting address to byte length. "
'@Parameters
'   pStartAddress LongPtr
'       Memory Address to begin memory dump
'   byteLength Long
'       Number of bytes for memory dump
'@Returns String
'   Returns a string memory dump for the provided starting address and byte length.
''
#If VBA7 Then
Public Function ToStringTable(ByVal pStartAddress As LongPtr, ByVal byteLength As Long) As String
#Else
Public Function ToStringTable(ByVal pStartAddress As Long, ByVal byteLength As Long) As String
#End If
    '@TODO Validate Arguments (Positive byteLength and bytesPerLine)
    'Raise error or exit. Preferred behaviour catch errors and exit function
    Dim mem() As Byte
    mem = Me.ToBytes(pStartAddress, byteLength)
    
#If VBA7 Then
    Dim memoryAddress As LongPtr
#Else
    Dim memoryAddress As Long
#End If

    memoryAddress = pStartAddress

    Dim output As String
    
    'Table column headings
    output = output & "Address      " & vbTab & "Hex" & vbTab & "Dec " & "ASCII " & "Binary" & vbNewLine

    Dim i As Long
    For i = 0 To byteLength - 1
        output = output & memoryAddress + i & vbTab & _
                Me.ToHexidecimal(mem(i)) & vbTab & _
                Me.ToDecimal(mem(i)) & vbTab & _
                "  " & Me.ToASCII(mem(i)) & vbTab & _
                "  " & Me.ToBinary(mem(i)) & vbNewLine
    Next i
    output = output & vbNewLine
    ToStringTable = output
End Function

'------------------------------------------------------------------
'Private Methods
'------------------------------------------------------------------

'2^i
Private Sub InitPowers()
    this.Powers = ArrayByte(1, 2, 4, 8, 16, 32, 64, 128)
End Sub

Private Function ArrayByte(ParamArray values() As Variant) As Byte()
    Dim result() As Byte
    ReDim result(LBound(values) To UBound(values))
    Dim i As Long
    
    For i = LBound(values) To UBound(values)
        result(i) = values(i)
    Next i
    ArrayByte = result
End Function


