Attribute VB_Name = "Testing_MemoryDump"
'@Folder("Utilities.MemoryDump")
Option Explicit

Public Sub TestMemoryDumpToString()
    Dim intTemp As Long
    intTemp = 12823456

    Debug.Print MemoryDump.ToString(VarPtr(intTemp), LenB(intTemp), MemoryDumpFormat.Hexadecimal)
    Debug.Print vbNewLine
    Debug.Print MemoryDump.ToString(VarPtr(intTemp), LenB(intTemp), MemoryDumpFormat.Binary, 1)
    Debug.Print vbNewLine
    Debug.Print MemoryDump.ToString(VarPtr(intTemp), LenB(intTemp), MemoryDumpFormat.Dec, 2)
End Sub

Public Sub TestMemoryDumpToStringAll()
    Dim intTemp As LongLong
    intTemp = 4532789
    
    Dim result As String
    
    result = MemoryDump.ToStringAll(VarPtr(intTemp), LenB(intTemp), 1)
    Debug.Print result
End Sub

Public Sub TestMemoryDumpToStringTable()
    Dim intTemp As LongLong
    intTemp = 4532789
    
    Dim result As String
    
    result = MemoryDump.ToStringTable(VarPtr(intTemp), LenB(intTemp))
    Debug.Print result
End Sub
