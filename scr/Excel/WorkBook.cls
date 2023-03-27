VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Workbook functions"
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'Rubberduck annotations
'@ModuleDescription "Workbook functions"
'@Folder("Excel.WorkBook")
'@PredeclaredId

'MarkJohnstoneGitHub/VBA-Library
'@Version v1.0 March 28, 2023
'@LastModified March 28, 2023

Option Explicit

''
'@Static
'@Description "Copies an 2D array to a specified Worksheet at a starting row and  starting column"
'@Parameters
'   data Variant
'       Array to copy to the specified Worksheet, starting row and starting column
'   sheetName String
'       Worksheet name
'   startRow Long
'       Starting row to copy array to.
'   startCol Long
'       Starting column to copy array to.

'@Remarks
' Assumes data is an valid array
' SheetName exists
' startRow and startColumn are valid
'
'@References
' https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
'
'@TODO
' Add row and column headings?
''
Sub From2DArray(ByRef data As Variant, sheetName As String, ByVal startRow As Integer, ByVal startColumn As Integer)
    Dim Rng As Range
    With Sheets(sheetName)
        Set Rng = .Range(.Cells(startRow, startColumn), _
            .Cells(UBound(data, 1) - LBound(data, 1) + startRow, _
            UBound(data, 2) - LBound(data, 2) + startColumn))
    End With
    Rng.Value2 = data
End Sub

''
'@Static
'@Description "Copies an 2D array to a specified Worksheet at a starting row and  starting column"
'@Parameters
'   data Variant
'       Array to copy to the specified Worksheet, starting row and starting column
'   sheetName String
'       Worksheet name
'   startRow Long
'       Starting row to copy array to.
'   startCol Long
'       Starting column to copy array to.

'@Remarks
' Assumes data is an valid array
' SheetName exists
' startRow and startColumn are valid
'
'@References
' https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
'
'@TODO
' Add row and column headings?
''
Sub From1DArray(ByRef data As Variant, sheetName As String, ByVal startRow As Integer, ByVal startColumn As Integer)
    From2DArray Application.WorksheetFunction.Transpose(data), sheetName, startRow, startColumn
End Sub

''
'@Static
'@Description "Turns off Excel Calculations, ScreenUpdating and Events"
'@Remarks
'   Turns off Excel Calculations, ScreenUpdating and Events for performance improvements when
'   running macros.
'@References
' Excel Macro Mastery
' https://www.youtube.com/watch?v=GCSF5tq7pZ0
''
Sub TurnOff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
End Sub

''
'@Static
'@Description "Turns on Excel Calculations, ScreenUpdating and Events"
'@References
' Excel Macro Mastery
' https://www.youtube.com/watch?v=GCSF5tq7pZ0
''
Sub TurnOn()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
