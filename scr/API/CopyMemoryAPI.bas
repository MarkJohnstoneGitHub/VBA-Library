Attribute VB_Name = "CopyMemoryAPI"
Attribute VB_Description = "API declarations for copy memory by pointer for Windows and Mac, with VBA6 and VBA7 compatibility."
'@Folder "VBACorLib.API"
'@ModuleDescription "API declarations for copy memory by pointer for Windows and Mac, with VBA6 and VBA7 compatibility."

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 February 25, 2023
'@LastModified February 25, 2023

''
'@API_Declaration
'@References
' https://stackoverflow.com/questions/45756170/how-to-read-write-memory-on-mac-os-x-with-vba
'
'@Remarks
'
''
Option Explicit

''
'@Description "Copy memory using pointers of source and destination variables. Compatibility for Windows and Mac, with VBA6 and VBA7."
''
#If Mac Then
  #If Win64 Then
    Public Declare PtrSafe Function CopyMemoryByPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As Long) _
             As LongPtr
  #Else
    Public Declare Function CopyMemoryByPtr Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As Long, _
             ByVal src As Long, _
             ByVal size As Long) _
             As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Public Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As LongLong)
  #Else
    Public Declare PtrSafe Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByVal source As LongPtr, _
             ByVal size As Long)
  #End If
#Else
  Public Declare Sub CopyMemoryByPtr Lib "kernel32" Alias "RtlMoveMemory" _
          (ByVal destination As Long, _
           ByVal src As Long, _
           ByVal size As Long)
#End If

''
'@Description "Copy any memory using pointer of source and variable for destination. Compatibility for Windows and Mac, with VBA6 and VBA7."
''
#If Mac Then
  #If Win64 Then
    Public Declare PtrSafe Function CopyAnyToMemory Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As LongPtr, _
             ByRef source As Any, _
             ByVal size As Long) _
             As LongPtr
  #Else
    Public Declare Function CopyAnyToMemory Lib "libc.dylib" Alias "memmove" _
            (ByVal destination As Long, _
             ByVal source As Any, _
             ByVal size As Long) _
             As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Public Declare PtrSafe Sub CopyAnyToMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
            (ByVal destination As LongPtr, _
             ByRef source As Any, _
             ByVal size As LongLong)
  #Else
    Public Declare PtrSafe Sub CopyAnyToMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
        (ByVal destination As LongPtr, _
         ByRef source As Any, _
         ByVal size As Long)
  #End If
#Else
    Public Declare Sub CopyAnyToMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
        (ByVal destination As Long, _
         ByRef source As Any, _
         ByVal size As Long)
#End If

''
'@Description "CopyMemory using variable of source and destination. Compatibility for Windows and Mac, with VBA6 and VBA7."
''
#If Mac Then
  #If Win64 Then
    Public Declare PtrSafe Function CopyMemory Lib "libc.dylib" Alias "memmove" _
            (ByRef destination As Any, _
             ByRef source As Any, _
             ByVal size As Long) _
             As LongPtr
  #Else
    Public Declare Function CopyMemory Lib "libc.dylib" Alias "memmove" _
            (ByRef destination As Any, _
             ByRef source As Any, _
             ByVal size As Long) _
             As Long
  #End If
#ElseIf VBA7 Then
  #If Win64 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
            (ByRef destination As Any, _
             ByRef source As Any, _
             ByVal size As LongLong)
  #Else
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
        (ByRef destination As Any, _
         ByRef source As Any, _
         ByVal size As Long)
  #End If
#Else
    Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" _
        (ByRef destination As Any, _
         ByRef source As Any, _
         ByVal size As Long)
#End If
