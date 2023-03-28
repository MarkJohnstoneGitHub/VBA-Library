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

'@TODO
' Add convert to and from bijective base-26

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
Public Function ColumnToArray(ByVal sheetName As String, ByVal column As Long, ByVal startRow As Long, ByVal endRow As Long, Optional ByVal zeroBaseArray As Boolean = False) As Variant
    Dim Rng As Range
    With Sheets(sheetName)
        Set Rng = .Range(.Cells(startRow, column), .Cells(endRow, column))
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
Public Function RowToArray(ByVal sheetName As String, ByVal row As Long, ByVal startColumn As Long, ByVal endColumn As Long, Optional ByVal zeroBaseArray As Boolean = False) As Variant
    Dim Rng As Range
    With Sheets(sheetName)
        Set Rng = .Range(.Cells(row, startColumn), .Cells(row, endColumn))
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
    RowToArray = output
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

'https://stackoverflow.com/questions/181596/how-to-convert-a-column-number-e-g-127-into-an-excel-column-e-g-aa
'Although there are already a bunch of valid answers1, none get into the theory behind it.
'
'Excel column names are bijective base-26 representations of their number. This is quite different than an ordinary base 26 (there is no leading zero), and I really recommend reading the Wikipedia entry to grasp the differences. For example, the decimal value 702 (decomposed in 26*26 + 26) is represented in "ordinary" base 26 by 110 (i.e. 1x26^2 + 1x26^1 + 0x26^0) and in bijective base-26 by ZZ (i.e. 26x26^1 + 26x26^0).
'
'Differences aside, bijective numeration is a positional notation, and as such we can perform conversions using an iterative (or recursive) algorithm which on each iteration finds the digit of the next position (similarly to an ordinary base conversion algorithm).
'
'The general formula to get the digit at the last position (the one indexed 0) of the bijective base-k representation of a decimal number m is (f being the ceiling function minus 1):
'
'm -(f(m / k) * k)
'The digit at the next position (i.e. the one indexed 1) is found by applying the same formula to the result of f(m / k). We know that for the last digit (i.e. the one with the highest index) f(m / k) is 0.
'
'This forms the basis for an iteration that finds each successive digit in bijective base-k of a decimal number. In pseudo-code it would look like this (digit() maps a decimal integer to its representation in the bijective base -- e.g. digit(1) would return A in bijective base-26):
'
'fun conv(m)
'    q = f(m / k)
'    a = m - (q * k)
'    if (q == 0)
'        return digit(a)
'    Else
'        return conv(q) + digit(a);
'So we can translate this to C#2 to get a generic3 "conversion to bijective base-k" ToBijective() routine:
'
'class BijectiveNumeration {
'    private int baseK;
'    private Func<int, char> getDigit;
'    public BijectiveNumeration(int baseK, Func<int, char> getDigit) {
'        this.baseK = baseK;
'        this.getDigit = getDigit;
'    }
'
'    public string ToBijective(double decimalValue) {
'        double q = f(decimalValue / baseK);
'        double a = decimalValue - (q * baseK);
'        return ((q > 0) ? ToBijective(q) : "") + getDigit((int)a);
'    }
'
'    private static double f(double i) {
'        return (Math.Ceiling(i) - 1);
'    }
'}
'Now for conversion to bijective base-26 (our "Excel column name" use case):
'
'static void Main(string[] args)
'{
'    BijectiveNumeration bijBase26 = new BijectiveNumeration(
'        26,
'        (value) => Convert.ToChar('A' + (value - 1))
'    );
'
'    Console.WriteLine(bijBase26.ToBijective(1));     // prints "A"
'    Console.WriteLine(bijBase26.ToBijective(26));    // prints "Z"
'    Console.WriteLine(bijBase26.ToBijective(27));    // prints "AA"
'    Console.WriteLine(bijBase26.ToBijective(702));   // prints "ZZ"
'    Console.WriteLine(bijBase26.ToBijective(16384)); // prints "XFD"
'}
'Excel 's maximum column index is 16384 / XFD, but this code will convert any positive number.
