VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Worksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Worksheet functions for the ActiveWorkbook"
'Rubberduck annotations
'@ModuleDescription "Worksheet functions for the ActiveWorkbook"
'@Folder("Excel.WorkSheet")
'@PredeclaredId


'MarkJohnstoneGitHub/VBA-Library
'@Version v1.0 March 28, 2023
'@LastModified April 02, 2023

'@References
' https://www.ExcelMacroMastery.com
' https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
' https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
' https://usefulgyaan.wordpress.com/2013/06/12/vba-trick-of-the-week-slicing-an-array-without-loop-application-index/
' https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
' https://www.exceldemy.com/excel-vba-get-row-and-column-number-from-cell-address/
' https://github.com/cristianbuse/VBA-FileTools

'@Remarks
' Work in progress
' @TODO Add flexibility selecting a Wookbook and remove coupling for sheetname
' Would require a class accepting, Workbook name and location and sheetname
' Function required such as opening a Workbook from a path see https://github.com/cristianbuse/VBA-FileTools
''

Option Explicit
'------------------------------------------------------------------
'Types
'------------------------------------------------------------------
Public Enum RangePosition
    First = 0
    Last = 1
End Enum

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
    With Application.ActiveWorkbook.Worksheets.Application
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
'@Description "Sets a cell value2 property for a specified Worksheet, row index and column index."
'@Parameters
'   sheetName String
'       Name of Worksheet to obtain the cell value2 property.
'   rowNumber Long
'       Row index of the cell to obtain its value2 property.
'   columnNumber Long
'       Column index of the cell to obtain its value2 property.
''
Public Property Let Value2(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnNumber As Long, ByVal val As Variant)
    Dim rng As Range
    With ActiveWorkbook.Sheets(sheetName)
        Set rng = .Range(.Cells(rowNumber, columnNumber), .Cells(rowNumber, columnNumber))
    End With
    rng.Value2 = val
End Property

''
'@Static
'@Description "Gets a cell value2 property for a specified Worksheet, row index and column index."
'@Parameters
'   sheetName String
'       Name of Worksheet to obtain the cell value2 property.
'   rowNumber Long
'       Row index of the cell to obtain its value2 property.
'   columnNumber Long
'       Column index of the cell to obtain its value2 property.
'@Returns Variant
'   Cell Value2 property for the specified Worksheet, row index and column index.
''
Public Property Get Value2(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnNumber As Long) As Variant
    Dim rng As Range
    With ActiveWorkbook.Sheets(sheetName)
        Set rng = .Range(.Cells(rowNumber, columnNumber), .Cells(rowNumber, columnNumber))
    End With
    Value2 = rng.Value2
End Property

''
'@Static
'@Description "Copies a two-dimensional array to a specified Worksheet at a starting row and starting column."
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
Public Sub FromArray2D(ByRef data As Variant, ByVal sheetName As String, ByVal startRow As Long, ByVal startColumn As Long)
    Dim rng As Range
    With ActiveWorkbook.Sheets(sheetName)
        Set rng = .Range(.Cells(startRow, startColumn), _
            .Cells(UBound(data, 1) - LBound(data, 1) + startRow, _
            UBound(data, 2) - LBound(data, 2) + startColumn))
    End With
    rng.Value2 = data
End Sub

'@TODO Require to check that dimensions of 2D array fit range provided?
'Used by RowFromArray1D and ColumnFromArray1D
'Currently private until decide if feature is requied
'Overloads FromArray2D signature with endRow and endColumn
Private Sub FromArrayTwoDimensional(ByRef data As Variant, ByVal sheetName As String, ByVal startRow As Long, ByVal startColumn As Long, ByVal endRow As Long, ByVal endColumn As Long)
    '@TODO Validate parameters?
    Dim rng As Range
    With ActiveWorkbook.Sheets(sheetName)
        Set rng = .Range(.Cells(startRow, startColumn), .Cells(endRow, endColumn))
    End With
    rng.Value2 = data
End Sub

''
'@Static
'@Description "Copies an one-dimensional array to row for a specified Worksheet, row and starting column."
'@Parameters
'   data Variant
'       One-dimensional array to copy to the specified Worksheet, starting row and starting column
'   sheetName String
'       Worksheet name
'   rowNumer Long
'       Row index to copy array to.
'   startCol Long
'       Starting column index to copy array to.
'@Remarks
' Assumes data is a valid one-dimensional array
' SheetName exists in active workbook
' startRow and startColumn are valid
'
'@References
' https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
'
'@TODO Add Error Handling validating parameters
'   Add row heading optional?
''
Public Sub RowFromArray1D(ByRef data As Variant, ByVal sheetName As String, ByVal rowNumber As Long, ByVal startColumn As Long)
    FromArrayTwoDimensional Application.WorksheetFunction.Transpose(data), sheetName, rowNumber, startColumn, rowNumber, UBound(data) - LBound(data) + startColumn
End Sub

''
'@Static
'@Description "Copies an one-dimensional array to column for a specified Worksheet, from a starting row and column."
'@Parameters
'   data Variant
'       One-dimensional array to copy to the specified Worksheet, from a starting row and column
'   sheetName String
'       Worksheet name
'   startRow Long
'       Row starting index to copy array to.
'   columnNumber Long
'       Column index to copy array to.
'@Remarks
' Assumes data is a valid one-dimensional array
' SheetName exists in active workbook
' startRow and columnNumber are valid ranges
'
'@References
' https://stackoverflow.com/questions/6063672/excel-vba-function-to-print-an-array-to-the-workbook
'
'@TODO Add Error Handling validating parameters
'   Add column heading optional?
''
Public Sub ColumnFromArray1D(ByRef data As Variant, ByVal sheetName As String, ByVal startRow As Long, ByVal columnNumber As Long)
    FromArrayTwoDimensional Application.WorksheetFunction.Transpose(data), sheetName, startRow, columnNumber, UBound(data) - LBound(data) + startRow, columnNumber
End Sub

''
'@Static
'@Description "Copies to one-dimensional variant array for a column of a specified Workseet."
'@Parameters
'   sheetName String
'       Array to copy to the specified Worksheet, starting row and starting column
'   columnNumber Long
'       Column index to copy to array
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
    Dim rng As Range
    With ActiveWorkbook.Sheets(sheetName)
        Set rng = .Range(.Cells(startRow, columnNumber), .Cells(endRow, columnNumber))
    End With
    Dim rangeArray As Variant
    rangeArray = rng.Value2
    
    Dim arrayBaseOffset As Long
    If zeroBaseArray Then
        arrayBaseOffset = 1
    End If
    Dim arrOutput As Variant
    ReDim arrOutput(LBound(rangeArray) - arrayBaseOffset To UBound(rangeArray) - arrayBaseOffset)
    Dim i As Long
    For i = LBound(rangeArray) To UBound(rangeArray)
        arrOutput(i - arrayBaseOffset) = rangeArray(i, 1)
    Next i
    ColumnToArray = arrOutput
End Function

''
'@Static
'@Description "Copies to one-dimensional variant array for a row of a specified Workseet."
'@Parameters
'   sheetName String
'       Array to copy to the specified Worksheet, starting row and starting column
'   rowNumber Long
'       Row index to copy to array
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
    Dim rng As Range
    With ActiveWorkbook.Sheets(sheetName)
        Set rng = .Range(.Cells(rowNumber, startColumn), .Cells(rowNumber, endColumn))
    End With
    Dim rangeArray As Variant
    rangeArray = rng.Value2
    
    Dim arrayBaseOffset As Long
    If zeroBaseArray Then
        arrayBaseOffset = 1
    End If
    Dim arrOutput As Variant
    ReDim arrOutput(LBound(rangeArray, 2) - arrayBaseOffset To UBound(rangeArray, 2) - arrayBaseOffset)
    Dim i As Long
    For i = LBound(rangeArray, 2) To UBound(rangeArray, 2)
        arrOutput(i - arrayBaseOffset) = rangeArray(1, i)
    Next i
    RowToArray = arrOutput
End Function

''
'@Static
'@Description "Test if a Worksheet or Range exists on a Worksheet"
'@Parameters
'   sheetName String
'       Name of the Worksheet to check
'   rangeAddress
'       Range address Optional, Default = "A1"
'@Remarks
'   Leave range blank to test if sheet exists
'@Reference
' https://www.automateexcel.com/vba/rename-sheet#check-if-sheet-name-exists
''
Public Function Exists(ByVal sheetName As String, Optional ByVal rangeAddress As String = "A1") As Boolean
    Dim test As Range
    On Error Resume Next
    Set test = ActiveWorkbook.Sheets(sheetName).Range(rangeAddress)
    Exists = Err.Number = 0
    On Error GoTo 0
End Function

''
'@Static
'@Description "Converts a given row and column index into its' corresponding cell address reference."
'@Parameters
'   rowNumber Long
'       Row index to convert to a cell address
'   columnNumber Long
'       Column index to convert to a cell address
'@TODO Option to return absolute reference or relative reference
'@TODO Option to return a range address?
''
Public Function CellAddress(ByVal rowNumber As Long, ByVal columnNumber As Long) As String
    CellAddress = Cells(rowNumber, columnNumber).Address
End Function

''
'@Static
'@Description "Obtains the row index from a range address."
'@Parameters
'   rangeAddress String
'   position RangePosition
'       First or Last row index to return for the range address
'@Returns Long
'   Row number from range address and range position i.e. first or last
''
Public Function Row(ByVal rangeAddress As String, Optional ByVal position As RangePosition = RangePosition.First) As Long
    '@TODO Test if rangeAddress is valid i.e. Exists, Not VBA.vbNullString etc
    With Range(rangeAddress)
        Select Case position
            Case RangePosition.First
                Row = .Row
            Case RangePosition.Last
                Row = .Row + .Rows.Count - 1
            Case Else
                '@TODO Raise Error
        End Select
    End With
End Function

''
'@Static
'@Description "Obtains the column index from a range address."
'@Parameters
'   rangeAddress String
'   position RangePosition
'       First or Last column index to return for the range address
'@Returns Long
'   Column number from range address and range position i.e. first or last
''
Public Function Column(ByVal rangeAddress As String, Optional ByVal position As RangePosition = RangePosition.First) As Long
    '@TODO Test if rangeAddress is valid i.e. Exists, Not VBA.vbNullString etc
    
    With Range(rangeAddress)
        Select Case position
            Case RangePosition.First
                Column = .Column
            Case RangePosition.Last
                Column = .Column + .Columns.Count - 1
            Case Else
                '@TODO Raise Error
        End Select
    End With
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
