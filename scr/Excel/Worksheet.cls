VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Worksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Workbook functions"
'Rubberduck annotations
'@ModuleDescription "Worksheet functions"
'@Folder("Excel.WorkSheet")
'@PredeclaredId

'MarkJohnstoneGitHub/VBA-Library
'@Version v1.0 March 28, 2023
'@LastModified March 28, 2023

'@References
' https://www.ExcelMacroMastery.com
' https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
' https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
' https://usefulgyaan.wordpress.com/2013/06/12/vba-trick-of-the-week-slicing-an-array-without-loop-application-index/
' https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
' https://www.exceldemy.com/excel-vba-get-row-and-column-number-from-cell-address/

Option Explicit
Private Type TWorksheet
    MaxRows     As Long
    MaxColumns  As Long
End Type

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As TWorksheet

'------------------------------------------------------------------
'Constructors and destructors
'------------------------------------------------------------------
Private Sub Class_Initialize()
    With Application.Worksheets.Application
        this.MaxRows = .Rows.Count
        this.MaxColumns = .Columns.Count
    End With
End Sub

''
'@Static
'@Description "Returns the maximum allowable rows in a WorkSheet."
'@Field Value Long
'   Value = 1048576 for Ms-Excel 2016
'@Remarks
'   Returns a version independent maximum allowable rows in a Worksheet.
''
Public Property Get MaxRows() As Long
    MaxRows = this.MaxRows
End Property

''
'@Static
'@Description "Returns the maximum allowable columns in a WorkSheet."
'@Field Value Long
'   Value = 16384 for Ms-Excel 2016
'@Remarks
'   Returns a version independent maximum allowable columns in a Worksheet.
''
Public Property Get MaxColumns() As Long
    MaxColumns = this.MaxColumns
End Property

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
Public Sub From2DArray(ByRef data As Variant, sheetName As String, ByVal startRow As Integer, ByVal startColumn As Integer)
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
' SheetName exists in active workbook
' startRow and startColumn are valid
'
'@References
' https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
'
'@TODO
' Add row and column headings?
''
Public Sub From1DArray(ByRef data As Variant, ByVal sheetName As String, ByVal startRow As Integer, ByVal startColumn As Integer)
    From2DArray Application.WorksheetFunction.Transpose(data), sheetName, startRow, startColumn
End Sub

''
'@Static
'@Description "Copies to one-dimensional variant array for a column of a specified Workseet."
'@Parameters
'   sheetName String
'       Array to copy to the specified Worksheet, starting row and starting column
'   column Long
'       Column indexto copy to array
'   startRow Long
'       Starting row index of column to copy to array.
'   endRow Long
'       Last row index of column to copy to array.
'   zeroBaseArray Boolean
'       Array returned is be zero based,  default is one based.
'@TODO
' Validation of parameters
'   sheetName exists
'   column doesnt exceed MaxColumn
'   startRow
'       doesn't exceed MaxRow
'       is less than or equal to EndRow
'   endRow
'       doesn't exceed MaxRow
''
Public Function ColumnToArray(ByVal sheetName As String, ByVal columnNumber As Long, ByVal startRow As Long, ByVal endRow As Long, Optional ByVal zeroBaseArray As Boolean = False) As Variant
    Dim Rng As Range
    With Sheets(sheetName)
        Set Rng = .Range(.Cells(startRow, columnNumber), .Cells(endRow, columnNumber))
    End With
    Dim rangeArray As Variant
    rangeArray = Rng.Value2
    
    Dim arrayBaseOffset As Long
    If zeroBaseArray Then
        arrayBaseOffset = 1
    End If
    Dim output As Variant
    ReDim output(LBound(rangeArray) - arrayBaseOffset To UBound(rangeArray) - arrayBaseOffset)
    Dim i As Long
    For i = LBound(rangeArray) To UBound(rangeArray)
        output(i - arrayBaseOffset) = rangeArray(i, 1)
    Next i
    ColumnToArray = output
End Function

''
'@Static
'@Description "Copies to one-dimensional variant array for a row of a specified Workseet."
'@Parameters
'   sheetName String
'       Array to copy to the specified Worksheet, starting row and starting column
'   row Long
'       Row to copy to array
'   startColumn Long
'       Starting column index of row to copy to array.
'   endColumn Long
'       Last column index of row to copy to array.
'   zeroBaseArray Boolean
'       Array returned is be zero based,  default is one based.
'@TODO
' Validation of parameters
'   sheetName exists
'   row doesnt exceed MaxRow
'   startColumn
'       doesn't exceed MaxColumn
'       is less than or equal to endColumn
'   endColumn
'       doesn't exceed MaxColumn
''
Public Function RowToArray(ByVal sheetName As String, ByVal rowNumber As Long, ByVal startColumn As Long, ByVal endColumn As Long, Optional ByVal zeroBaseArray As Boolean = False) As Variant
    Dim Rng As Range
    With Sheets(sheetName)
        Set Rng = .Range(.Cells(rowNumber, startColumn), .Cells(rowNumber, endColumn))
    End With
    Dim rangeArray As Variant
    rangeArray = Rng.Value2
    
    Dim arrayBaseOffset As Long
    If zeroBaseArray Then
        arrayBaseOffset = 1
    End If
    Dim output As Variant
    ReDim output(LBound(rangeArray, 2) - arrayBaseOffset To UBound(rangeArray, 2) - arrayBaseOffset)
    Dim i As Long
    For i = LBound(rangeArray, 2) To UBound(rangeArray, 2)
        output(i - arrayBaseOffset) = rangeArray(1, i)
    Next i
    RowToArray = output
End Function

''
'@Static
'@Description "Obtains a row number from a cell address."
'@Parameters
'   cellAddress String
'@Returns Long
'   Row number from address
Public Function GetRowFromAddress(ByVal cellAddress As String) As Long
    GetRowFromAddress = Range(cellAddress).Row
End Function

''
'@Static
'@Description "Obtains a column number from a cell address."
'@Parameters
'   cellAddress String
'@Returns Long
'   Column number from address
Public Function GetColumnFromAddress(ByVal cellAddress As String) As Long
    GetColumnFromAddress = Range(cellAddress).Column
End Function

''
'@Static
'@Description "Turn off automatic calculations, events and screen updating."
'@References
' https://www.ExcelMacroMastery.com
''
Public Sub TurnOffFunctionality()
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

''
'@Static
'@Description "Turn on automatic calculations, events and screen updating"
'@References
' https://www.ExcelMacroMastery.com
''
Public Sub TurnOnFunctionality()
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
