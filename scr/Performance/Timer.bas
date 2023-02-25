Attribute VB_Name = "Timer"
Attribute VB_Description = "Micro timer to measure performance."
'@ModuleDescription "Micro timer returning number of seconds."
'@Folder("Utilities.Performance")

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library is licensed under the MIT License
'@Version v1.0 February 26, 2023
'@LastModified February 26, 2023

'@References
'https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff700515(v=office.14)#Office2007excelPerf_MakingWorkbooksCalculateFaster
'https://codereview.stackexchange.com/questions/67596/a-lightning-fast-stringbuilder

Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function getFrequency Lib "kernel32" _
        Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" _
        Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" _
        Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" _
        Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If


''
'@Description"MicroTimer returning the number of seconds."
'@Returns Double
'   Returns The number of seconds
'@Usage
'    Dim dTime As Double
' Initialize
'   dTime = MicroTimer
' Calculate duration.
'   dTime = MicroTimer - dTime
'@Reference
' https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff700515(v=office.14)?redirectedfrom=MSDN#Office2007excelPerf_MakingWorkbooksCalculateFaster
''
Public Function MicroTimer() As Double

' Returns seconds.
    Dim cyTicks1 As Currency
    Static cyFrequency As Currency
    '
    MicroTimer = 0

    ' Get frequency.
    If cyFrequency = 0 Then getFrequency cyFrequency

    ' Get ticks.
    getTickCount cyTicks1

    ' Seconds
    If cyFrequency Then MicroTimer = cyTicks1 / cyFrequency
End Function
