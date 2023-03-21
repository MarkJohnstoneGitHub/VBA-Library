VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ULong32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a 32-bit unsigned integer."
'Rubberduck annotations
'@Folder("VBACorLib.DataTypes")
'@ModuleDescription "Represents a 32-bit unsigned integer."
'@PredeclaredId
'@IgnoreModule ImplicitlyTypedConst

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.4 March 21, 2023
'@LastModified March 22, 2023

'@Dependencies
'   ULongType.bas

'@Remarks
'   Version v1.4
'   Performance improvements to math functions, Add, Subtract, Divide, Multiply, DivRem
'   CreateSaturating
'       Changed to truncate decimal values for consistency with .Net 7.0 behaviour.
'   CreateChecked
'       Removed String parsing, only numeric types allowed  for consistency with .Net 7.0 behaviour.
'   CreateTruncating
'       Added suppport for Decimal, Double, Single data types.
'       Note for these types the behaviour is CreateSaturating as per .Net 7.0 behaviour.
'   Parse
'       Added Hex and Octal parsing
'       Raise error ArgumentException for decimal point for consistency with Net 7.0 behaviour.
'   Log2 Added function
'
'   Added Friend properties for bitwise arrays for reuse in ULongLong64.
'   Possibly move bitwise arrays to separate BitOperator class.

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.uint32?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/UInt32.cs
' https://newtonexcelbach.com/2015/10/26/the-vba-decimal-data-type/
' https://www.vbforums.com/showthread.php?834827-The-Decimal-Data-Type
' https://stackoverflow.com/questions/52099277/what-is-the-largest-possible-decimal-value-in-vba
' https://wiki.sei.cmu.edu/confluence/display/c/INT30-C.+Ensure+that+unsigned+integer+operations+do+not+wrap
' https://nolongerset.com/return-multiple-values-udt/
' https://cdn.ttgtmedia.com/searchEnterpriseLinux/downloads/WGC_Chapter_3.pdf
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Math.cs,bc38e67ca558e060
' https://onlinetoolz.net/bitshift

'@TODO
'   Implement functions
'       PopCount
'       LeadingZeroCount
'       TryFormat
'       TryParse
'
'   Testing for Bitwise functions
'   Testing for Clamp
'   Testing for Win32 compiler constraints
'   Unit testing

Option Explicit

'------------------------------------------------------------------
'Constants
'------------------------------------------------------------------
#If Win64 Then
Private Const MINUInt32 As LongLong = 0
Private Const MAXUInt32 As LongLong = 4294967295#
Private Const MAXInt32 As LongLong = 2147483647
Private Const NegativeInt32 As LongLong = 4294967296# ' For converting negative values to positive ULong values
#Else
Private Const MINUInt32 As Currency = 0
Private Const MAXUInt32 As Currency = 4294967295@
Private Const MAXInt32 As Currency = 2147483647@
Private Const NegativeInt32 As Currency = 4294967296@ ' For converting negative values to positive ULong values
#End If

Private Const MIN_ULONG As Long = 0
Private Const MAX_ULONG As Long = &HFFFFFFFF
Private Const MAXVALUE_Int32 As Long = 2147483647

'Bit shifting constants
Private Const CHAR_BIT As Long = 8
Private Const SIZEOF_UINT32 As Long = 4
Private Const SIZEOF_UINT32_BITS As Long = SIZEOF_UINT32 * CHAR_BIT
Private Const MIN_BITSHIFT As Long = 0
Private Const MAX_BITSHIFT As Long = 32
Private Const MOST_SIGNIFICANT_BIT_MASK As Long = &H80000000

'Error number constants
Private Const OverflowException As Long = &H6&
Private Const ArgumentException As Long = &H5&  'Invalid procedure call or argument.
Private Const DivideByZeroException As Long = &H11&
Private Const ArgumentNullException As Long = &H80004003
Private Const MinMaxException As Long = &H80004004
Private Const NotImplementedException As Long = 9999

'------------------------------------------------------------------
'Types
'------------------------------------------------------------------
#If Win64 Then
Private Type QWORD
    Value As LongLong
End Type
#Else
Private Type QWORD
    LowPart     As Long  ' the ordering is important to remain consistant with memory layout of a 64-bit integer.
    HighPart    As Long
End Type
#End If

Private Type QWORD_Currency
    Value As Currency
End Type

Private Type WORD
    Value    As Integer
End Type

Private Type TULong32
    decMaxUInt32 As Variant 'contains a decimal type
    maxVal  As ULong
    minVal  As ULong
    ulngOne As ULong    'For caculation in IsPow2
    Powers() As Long    'Used for bitwise operations
    Masks() As Long     'Used for bitwise operations
    RevMasks() As Long  'Used for bitwise operations
End Type

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As TULong32

'------------------------------------------------------------------
'Constructors and destructors
'------------------------------------------------------------------
Private Sub Class_Initialize()
    this.decMaxUInt32 = CDec(MAXUInt32)
    this.ulngOne.Value = 1 'constant for caculation in IsPow2
    this.maxVal.Value = MAX_ULONG
    this.minVal.Value = MIN_ULONG
    
    'Initialise bitwise arrays
    InitPowers
    InitMasks
    InitRevMasks
End Sub

'------------------------------------------------------------------
'Properties Internal
'------------------------------------------------------------------
Friend Property Get Zero() As ULong
    Zero.Value = 0
End Property

Friend Property Get One() As ULong
    One.Value = 1
End Property

'2^i
Friend Property Get Powers() As Long()
    Powers = this.Powers
End Property

'2^i-1
Friend Property Get Masks() As Long()
    Masks = this.Masks
End Property

'-2^i
Friend Property Get RevMasks() As Long()
    RevMasks = this.RevMasks
End Property

'------------------------------------------------------------------
'Properties Public
'------------------------------------------------------------------
''
'@Description"Represents the largest possible value of UInt32. This field is constant."
'@Field Value ULong
'   Value = 4294967295#
'@Remarks
'   The value of this constant is 4,294,967,295; that is, hexadecimal 0xFFFFFFFF.
''
Public Property Get MaxValue() As ULong
    MaxValue = this.maxVal
End Property

''
'@Description"Represents the smallest possible value of UInt32. This field is constant."
'@Field Value UInt32
'   Value = 0
'@Remarks
'   The value of this constant is 0.
''
Public Property Get MinValue() As ULong
    MinValue = this.minVal
End Property

'------------------------------------------------------------------
'Public Methods
'------------------------------------------------------------------

''
'@Static
'@Description "Returns a new unsigned32-bit value whose value is the absolute value"
'@Parameters
'   val: ULong
'        The value to return the absolute value of.
'@Returns ULong
'   ULong
'       The absolute value of the required value.
''
Public Function Absolute(ByRef val As ULong) As ULong
    Absolute.Value = val.Value
End Function

''
'@Static
'@Description "Copies the source unsigned 32 bit value to the destination."
'@Parameters
'   destination ULong
'       The destination which the source value is copied to.
'   source  ULong
'       The value to be copied to the destination.
''
Public Sub Assign(ByRef destination As ULong, ByRef source As ULong)
    destination.Value = source.Value
End Sub

''
'@Static
'@Description "Clamps a value to an inclusive minimum and maximum value."
'@Parameters
'   value ULong
'       The value to clamp.
'   min ULong
'       The inclusive minimum to which value should clamp.
'   max ULong
'       The inclusive maximum to which value should clamp.
'@Returns ULong
'   The result of clamping value to the inclusive range of min and max.
'
'@Reference
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Math.cs,bc38e67ca558e060
''
Public Function Clamp(ByRef val As ULong, ByRef minVal As ULong, ByRef maxVal As ULong) As ULong
    If Me.GreaterThan(minVal, maxVal) Then
        Err.Raise MinMaxException, "ULong32.Clamp", "Argument min value cannot be greater than max value."
    End If
    If Me.LessThan(val, minVal) Then
        Clamp.Value = minVal.Value
        Exit Function
    ElseIf Me.GreaterThan(val, maxVal) Then
        Clamp.Value = maxVal.Value
        Exit Function
    End If
    Clamp.Value = val.Value
End Function

''
'@Static
'@Description "Creates an unsigned 32-bit integer from a value, truncating any values that fall outside the representable range of the current type."
'@Parameters
'   val Variant
'       The value to be converted to an unsigned 32-bit integer.
'       The byte value is copied to an unsigned 32-bit integer where negative values are converted
'       into larger unsigned values.
'@Returns ULong
'   The unsigned 32-bit integer created from the byte value.
'
'@Exceptions
'   ArgumentException
'       Raised for when not a valid type to cast of a Byte, Integer, Long, LongLong or Currency
'
'@Remarks
'   Bytes values to be converted maybe of types Byte, Long, Integers or Currency.
'   Negative values are converted into a larger unsigned 32-bit integers.
'
'   For the Byte type it is copied to the lower byte of the ULong DWORD.
'   eg. If byte value is 255 Hex FF is converted to the ULong value of 255, Hex 000000FF
'
'   For the Integer type its WORD value is copied to the ULong DWORD lower WORD.
'   eg. If Integer value is -1 Hex FFFF is converted to ULong value of 65535 Hex 0000FFFF
'
'   For the Long type its DWORD value is copied to the ULong DWORD
'   eg. If Long value is -1 Hex FFFFFFFF is converted to ULong value of 4294967295 Hex FFFFFFFF
'
'   For the Currency type the low DWORD of a currency value is copied to the DWORD of the ULong.
'   Eg Currency value of 0.0001 Hex 00000000 00000001 converts to ULong 00000001 i.e. of value 1.
'
'   For the LongLong type the low DWORD of a LongLong value is copied to the DWORD of the ULong.
'   eg. If LongLong value is 42949672958 i.e. Hex 00000009 FFFFFFFE is converted to ULong value Hex FFFFFFFE
'   i.e. value of 4294967294
'
'   The Create* methods also have some special considerations for IEEE 754 floating-point types,
'   like float and double, as they have the special values PositiveInfinity, NegativeInfinity, and
'   NaN. All three Create* APIs behave as CreateSaturating. Also, while MinValue and MaxValue
'   represent the largest negative/positive "normal" number, the actual minimum and maximum values
'   are NegativeInfinity and PositiveInfinity, so they clamp to these values instead.

'@References
' https://github.com/dotnet/docs/blob/main/docs/standard/generics/math.md
' https://learn.microsoft.com/en-us/dotnet/standard/generics/math
''
Public Function CreateTruncating(ByVal val As Variant) As ULong
    Select Case VarType(val)
        Case VBA.vbLong
            CreateTruncating.Value = val
        Case VBA.vbByte
            CreateTruncating.Value = val
        Case VBA.vbInteger
            'copy Integer type WORD to lower WORD of ULong
            Dim wVal As WORD
            wVal.Value = val
            LSet CreateTruncating = wVal
        #If Win64 Then
        Case VBA.vbLongLong
            'copy low DWORD of LongLong type to ULong
            Dim qwVal As QWORD
            qwVal.Value = val
            LSet CreateTruncating = qwVal
        #End If
        Case VBA.vbCurrency
            'copy low DWORD of Currency type to ULong
            Dim qwValCurrency As QWORD_Currency
            qwValCurrency.Value = val
            LSet CreateTruncating = qwValCurrency
        'vbDecimal, vbDouble, vbSingle behave as CreateSaturating
        Case VBA.vbDecimal
            If val < MIN_ULONG Then
                CreateTruncating.Value = this.minVal.Value
            ElseIf val > this.decMaxUInt32 Then
                CreateTruncating.Value = this.maxVal.Value
            Else
                CreateTruncating = DecimalToULong(val)
            End If
        Case VBA.vbDouble
            If val < MIN_ULONG Then
                CreateTruncating.Value = this.minVal.Value
            ElseIf val > MAXUInt32 Then
                CreateTruncating.Value = this.maxVal.Value
            Else
                CreateTruncating.Value = DoubleToULong(val).Value
            End If
        Case VBA.vbSingle
            If val < MIN_ULONG Then
                CreateTruncating.Value = this.minVal.Value
            ElseIf val > MAXUInt32 Then
                CreateTruncating.Value = this.maxVal.Value
            Else
                CreateTruncating.Value = SingleToULong(val).Value
            End If
        Case Else
            Err.Raise ArgumentException, "ULong32.CreateTruncating", "Invalid argument"
    End Select
End Function

''
'@Static
'@Description "Converts a value to unsigned 32-bit integer."
'@Parameters
'   val Variant
'       value to be converted to an unsigned 32-bit value
'@Returns ULong
'   value converted to an unsigned 32 bit value
'
'@Exceptions
'   OverflowException
'       Raised when a value is less then 0 or exceeds the max unsigned 32-bit value of 4294967295
'   ArgumentException
'       Raised for an invalid value which is not numeric.
'@Remarks
'   Valid types to convert are Byte, Integer, Long, LongLong, Currency, Double, Single, Strings
'   containing numeric or hex strings.
'   Negative values return an overflow exception.
'
'   Be careful using CreateChecked with decimal places.
'   Decimal places are truncated.
''
Public Function CreateChecked(ByVal val As Variant) As ULong
    Select Case VarType(val)
        Case VBA.vbLong
            If val < MINUInt32 Then
                Err.Raise OverflowException, "ULong32.CreateChecked"
            Else
                CreateChecked.Value = val
            End If
        Case VBA.vbByte
            CreateChecked.Value = val
        Case VBA.vbInteger
            If val < MINUInt32 Then
                Err.Raise OverflowException, "ULong32.CreateChecked"
            Else
                CreateChecked.Value = val
            End If
        Case VBA.vbCurrency
            If val < MINUInt32 Or val > MAXUInt32 Then
                Err.Raise OverflowException, "ULong32.CreateChecked"
            Else
                CreateChecked = CurrencyToULong(val)
            End If
        Case VBA.vbDecimal
            If val < MINUInt32 Or val > this.decMaxUInt32 Then
                Err.Raise OverflowException, "ULong32.CreateChecked"
            Else
                CreateChecked = DecimalToULong(val)
            End If
        #If Win64 Then
        Case VBA.vbLongLong
            If val < MINUInt32 Or val > MAXUInt32 Then
                Err.Raise OverflowException, "ULong32.CreateChecked"
            Else
                CreateChecked.Value = LongLongToULong(val).Value
            End If
        #End If
        Case VBA.vbDouble
            If val < MINUInt32 Or val > MAXUInt32 Then
                Err.Raise OverflowException, "ULong32.CreateChecked"
            Else
                CreateChecked.Value = DoubleToULong(val).Value
            End If
        Case VBA.vbSingle
            If val < MINUInt32 Or val > MAXUInt32 Then
                Err.Raise OverflowException, "ULong32.CreateChecked"
            Else
                CreateChecked.Value = SingleToULong(val).Value
            End If
        Case Else
            Err.Raise ArgumentException, "ULong32.CreateChecked", "Invalid argument"
    End Select
End Function

''
'@Static
'@Description "Creates an instance of the current type from a value, saturating any values that fall outside the representable range of the current type."
'   val Variant
'       The value which is used to create the instance of an unsigned 32-bit integer.
'@Remarks
'   Creates a value, clamping to T.MinValue or T.MaxValue if the input can't fit.
''
Public Function CreateSaturating(ByVal val As Variant) As ULong
    Select Case VarType(val)
        Case VBA.vbLong
            If val < MIN_ULONG Then
                CreateSaturating.Value = this.minVal.Value
            Else
                CreateSaturating.Value = val
            End If
        Case VBA.vbByte
            CreateSaturating.Value = val
        Case VBA.vbInteger
            If val < MIN_ULONG Then
                CreateSaturating.Value = this.minVal.Value
            Else
                CreateSaturating.Value = val
            End If
        Case VBA.vbCurrency
            If val < MIN_ULONG Then
                CreateSaturating.Value = this.minVal.Value
            ElseIf val > MAXUInt32 Then
                CreateSaturating.Value = this.maxVal.Value
            Else
                CreateSaturating = CurrencyToULong(val)
            End If
        Case VBA.vbDecimal
            If val < MIN_ULONG Then
                CreateSaturating.Value = this.minVal.Value
            ElseIf val > this.decMaxUInt32 Then
                CreateSaturating.Value = this.maxVal.Value
            Else
                CreateSaturating = DecimalToULong(val)
            End If
        #If Win64 Then
        Case VBA.vbLongLong
            If val < MIN_ULONG Then
                CreateSaturating.Value = this.minVal.Value
            ElseIf val > MAXUInt32 Then
                CreateSaturating.Value = this.maxVal.Value
            Else
                CreateSaturating.Value = LongLongToULong(val).Value
            End If
        #End If
        Case VBA.vbDouble
            If val < MIN_ULONG Then
                CreateSaturating.Value = this.minVal.Value
            ElseIf val > MAXUInt32 Then
                CreateSaturating.Value = this.maxVal.Value
            Else
                CreateSaturating.Value = DoubleToULong(val).Value
            End If
        Case VBA.vbSingle
            If val < MIN_ULONG Then
                CreateSaturating.Value = this.minVal.Value
            ElseIf val > MAXUInt32 Then
                CreateSaturating.Value = this.maxVal.Value
            Else
                CreateSaturating.Value = SingleToULong(val).Value
            End If
        Case Else
            Err.Raise ArgumentException, "ULong32.CreateSaturating", "Invalid argument"
    End Select
End Function

''
'@Static
'@Description "Compares two 32-bit unsigned integers and returns an indication of their relative values."
'@Parameters
'   lhs: ULong
'       unsigned long value to compare to rhs
'   rhs: ULong
'       unsigned long value to compare to lhs
'
'@Returns Long
'   Returns If greater than is 1, equals is 0, less than is -1
'@Remarks
''
Public Function CompareTo(ByRef lhs As ULong, ByRef rhs As ULong) As Long
    'If the most significant bit is different
    If (lhs.Value Xor rhs.Value) < 0 Then
        'If left most significant bit set it is the larger unsigned value
        If lhs.Value And MOST_SIGNIFICANT_BIT_MASK Then
             CompareTo = 1
             Exit Function
        Else
            CompareTo = -1
            Exit Function
        End If
    End If
    'If "same sign" i.e. the most significant bit are the same compare raw Int32 values
    If lhs.Value > rhs.Value Then
        CompareTo = 1
    ElseIf lhs.Value < rhs.Value Then
        CompareTo = -1
    End If
End Function

''
'@Static
'@Description "Returns the quotient and the remainder of two signed 64-bit numbers."
'@Parameters
'   dividend ULong
'       The value to be divided.
'   divisor ULong
'       The value to divide by.
'   outRemainder ULong
'       returns the remainder of the dividend and divisor
'@Returns ULong
'   The quotient of the of the dividend and divisor.
'
'@Exceptions
'   DivideByZeroException
'       Raised when a divisor is 0, runtime error 11
''
#If Win64 Then
Public Function DivRem(ByRef dividend As ULong, ByRef divisor As ULong, ByRef outRemainder As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "ULong32.DivRem"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        outRemainder.Value = 0
        DivRem.Value = 0
        Exit Function
    ElseIf divisor.Value = 1 Then
        outRemainder.Value = 0
        DivRem.Value = dividend.Value
        Exit Function
    End If

    'Convert dividend and divisor to LongLong
    Dim lnglngDividend As LongLong
    If dividend.Value < 0 Then
        lnglngDividend = dividend.Value + NegativeInt32
    Else
        lnglngDividend = dividend.Value
    End If
    Dim lnglngDivisor As LongLong
    If divisor.Value < 0 Then
        lnglngDivisor = divisor.Value + NegativeInt32
    Else
        lnglngDivisor = divisor.Value
    End If
    
    Dim lnglngQuotient As LongLong
    lnglngQuotient = lnglngDividend \ lnglngDivisor 'Using integer division which truncates decimal places
    Dim lnglngRemainder As LongLong
    lnglngRemainder = lnglngDividend - (lnglngDivisor * lnglngQuotient)
    
    'Convert lnglngQuotient to ULong
    If lnglngQuotient > MAXInt32 Then
        DivRem.Value = CLng(lnglngQuotient - NegativeInt32)
    Else
        DivRem.Value = CLng(lnglngQuotient)
    End If
    'Convert lnglngRemainder to outRemainder ULong
    If lnglngRemainder > MAXInt32 Then
        outRemainder.Value = CLng(lnglngRemainder - NegativeInt32)
    Else
        outRemainder.Value = CLng(lnglngRemainder)
    End If
End Function
#Else
Public Function DivRem(ByRef dividend As ULong, ByRef divisor As ULong, ByRef outRemainder As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "ULong32.DivRem"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        outRemainder.Value = 0
        DivRem.Value = 0
        Exit Function
    ElseIf divisor.Value = 1 Then
        outRemainder.Value = 0
        DivRem.Value = dividend.Value
        Exit Function
    End If

    'Convert dividend and divisor to Currency
    Dim curDividend As Currency
    If dividend.Value < 0 Then
        curDividend = dividend.Value + NegativeInt32
    Else
        curDividend = dividend.Value
    End If
    Dim curDivisor As Currency
    If divisor.Value < 0 Then
        curDivisor = divisor.Value + NegativeInt32
    Else
        curDivisor = divisor.Value
    End If
    
    Dim curQuotient As Currency
    curQuotient = Fix(curDividend / curDivisor) 'Using integer division which truncates decimal places
    Dim curRemainder As Currency
    curRemainder = curDividend - (curDivisor * curQuotient)
    
    'Convert curQuotient to ULong
    If curQuotient > MAXInt32 Then
        DivRem.Value = CLng(curQuotient - NegativeInt32)
    Else
        DivRem.Value = CLng(curQuotient)
    End If
    'Convert curRemainder to outRemainder ULong
    If curRemainder > MAXInt32 Then
        outRemainder.Value = CLng(curRemainder - NegativeInt32)
    Else
        outRemainder.Value = CLng(curRemainder)
    End If
End Function
#End If

''
'@Static
'@Description("Determines if a value represents an even integral number.")
'@Parameters
'   Val ULong
'       The value to be checked.
'@Returns Boolean
'   true if value is an even integer; otherwise, false.
''
Public Function IsEvenInteger(ByRef val As ULong) As Boolean
    If (val.Value And 1) = 0 Then
        IsEvenInteger = True
    End If
End Function

''
'@Static
'@Description("Determines if a value represents an odd integral number.")
'@Parameters
'   Val ULong
'       The value to be checked.
'@Returns Boolean
'   true if value is an odd integer; otherwise, false.
''
Public Function IsOddInteger(ByRef val As ULong) As Boolean
    If (val.Value And 1) <> 0 Then
        IsOddInteger = True
    End If
End Function

''
'@Static
'@Description "Determines if a value is a power of two."
'@Parameters
'    val ULong
'       The value to be checked.
'
'@Returns Boolean
'   true if value is a power of two; otherwise, false.
''
Public Function IsPow2(ByRef val As ULong) As Boolean
    If val.Value <> 0 Then 'to avoid overflow with subtraction of 0
        If (val.Value And (Me.Subtract(val, this.ulngOne).Value)) = 0 And val.Value <> 0 Then
            IsPow2 = True
        End If
    End If
End Function

''
'@Static
'@Description "Computes the number of leading zeros in a value."
'@Parameters
'   val ULong
'       The value whose leading zero bits are to be counted.
'
'@Returns ULong
'   The number of leading zeros in value.
'
'@References
' https://docs.oracle.com/javase/1.5.0/docs/api/java/lang/Long.html#numberOfLeadingZeros(long)
''
Public Function LeadingZeroCount(ByRef val As ULong) As ULong
    Const LOG_2 As Double = 0.693147180559945
    Const NegativeInt32 As Double = 4294967296#

    'Convert val to double
    Dim dbVal As Double
    If val.Value < MIN_ULONG Then
        dbVal = val.Value + NegativeInt32
    Else
        dbVal = val.Value
    End If
    
    LeadingZeroCount = 31 - CLng(Fix(Log(dbVal) / LOG_2))
End Function


''
'@Static
'@Description "Computes the log2 of a value."
'@Parameters
'    val ULong
'       The value whose log2 is to be computed.
'
'@Returns ULong
'   The log2 of value.
'@References
' https://www.oreilly.com/library/view/vb-vba/1565923588/1565923588_ch07-1509-fm2xml.html
'
'@Remarks
'  Decimal result is truncated as per .Net 7.0 behaviour
''
Public Function Log2(ByRef val As ULong) As ULong
    Const LOG_2 As Double = 0.693147180559945
    Const NegativeInt32 As Double = 4294967296#

    'Convert val to double
    Dim dbVal As Double
    If val.Value < MIN_ULONG Then
        dbVal = val.Value + NegativeInt32
    Else
        dbVal = val.Value
    End If
    
    Log2.Value = CLng(Fix(Log(dbVal) / LOG_2))
End Function


''
'@Static
'@Description "Calculates the remainder of two 32-bit unsigned integers."
'@Parameters
'   dividend ULong
'       The value to be divided.
'   divisor ULong
'       The value to divide by.
'@Returns ULong
'   Returns the remainder portion of the division operation.
'
'@Exceptions
'   DivideByZeroException
'       Raised when a divisor is 0, runtime error 11
'
'@Remarks
'   A modulus operation is used to calculate the remainder of a division operation.
''
Public Function Modulus(ByRef dividend As ULong, ByRef divisor As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "ULong32.Modulus"
    End If
    DivRem dividend, divisor, Modulus
End Function

''
'@Static
'@Description "Compares two values to compute which is greater."
'@Parameters
'   lhs ULong
'       The value to compare with rhs.
'   rhs ULong
'       The value to compare with lhs.
'
'@Returns ULong
'   lhs if it is greater than rhs; otherwise, rhs.
'
'@Remarks
''
Public Function Max(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    If CompareTo(lhs, rhs) > 0 Then
        Max.Value = lhs.Value
    Else
        Max.Value = rhs.Value
    End If
End Function

''
'@Static
'@Description "Compares two values to compute which is lesser."
'@Parameters
'   lhs ULong
'       The value to compare with rhs.
'   rhs ULong
'       The value to compare with lhs.
'
'@Returns ULong
'   lhs if it is less than rhs; otherwise, rhs.
''
Public Function Min(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    If CompareTo(lhs, rhs) < 0 Then
        Min.Value = lhs.Value
    Else
        Min.Value = rhs.Value
    End If
End Function

''
'@Static
'@Description "Converts the string representation of a number to its 32-bit unsigned integer equivalent."
'@Parameters
'   s String
'       A string representing the number to convert.
'
'@Returns ULong
'   A 32-bit unsigned integer equivalent to the number contained in s.
'
'@Exceptions
'   ArgumentNullException
'       The s parameter is null.
'
'   ArgumentException
'       The s parameter is not of the correct format.
'
'   OverflowException
'       The s parameter represents a number that is less than UInt32.MinValue or greater than UInt32.MaxValue.
'@Remarks
'   .Net 7.0 behaviour raises error ArgumentException for decimal places in string
'   Added support for hexidecimal and octal literal strings.
''
#If Win64 Then
Public Function Parse(ByRef s As String) As ULong
    Const HEX_LITERAL As String = "&H"
    Const OCT_LITERAL As String = "&O"
    Const DECIMALPOINT As String = "."
    
    Dim strInput As String
    strInput = Trim(s)
    If (strInput = VBA.vbNullString) Then
        Err.Raise ArgumentNullException, "ULong32.Parse", "Argument Null Exception"
    End If
    '.Net raises ArgumentException for decimal point
    If InStr(1, strInput, DECIMALPOINT) <> 0 Then
        Err.Raise ArgumentException, "ULong32.Parse", "Invalid argument"
    End If
    'Convert Hex or Octal literal to ULong
    If Left$(strInput, 2) = HEX_LITERAL Or Left$(strInput, 2) = OCT_LITERAL Then
        On Error GoTo ErrorHandler
        Parse.Value = strInput
        On Error GoTo 0         'Disable ErrorHandler
        Exit Function
    End If
    'Convert decimal input to ULong
    On Error GoTo ErrorHandler
    Dim val As LongLong
    val = strInput
    On Error GoTo 0         'Disable ErrorHandler
    If val < MINUInt32 Or val > MAXUInt32 Then
        Err.Raise OverflowException, "ULong32.Parse"
    End If
    'convert LongLong to ULong
    If val > MAXInt32 Then
        Parse.Value = CLng(val - NegativeInt32)
    Else
        Parse.Value = CLng(val)
    End If
    Exit Function
ErrorHandler:
    If Err.Number = OverflowException Then
        Err.Raise Err.Number, "ULong32.Parse", Err.Description
    Else
        Err.Raise ArgumentException, "ULong32.Parse", "Invalid argument"
    End If
End Function

#Else
Public Function Parse(ByRef s As String) As ULong
    Const HEX_LITERAL As String = "&H"
    Const OCT_LITERAL As String = "&O"
    Const DECIMALPOINT As String = "."
    
    Dim strInput As String
    strInput = Trim(s)
    If (strInput = VBA.vbNullString) Then
        Err.Raise ArgumentNullException, "ULong32.Parse", "Argument Null Exception"
    End If
    '.Net raises ArgumentException for decimal point
    If InStr(1, strInput, DECIMALPOINT) <> 0 Then
        Err.Raise ArgumentException, "ULong32.Parse", "Invalid argument"
    End If
    'Convert Hex or Octal literal to ULong
    If Left$(strInput, 2) = HEX_LITERAL Or Left$(strInput, 2) = OCT_LITERAL Then
        On Error GoTo ErrorHandler
        Parse.Value = strInput
        On Error GoTo 0         'Disable ErrorHandler
        Exit Function
    End If
    'Convert decimal input to ULong
    On Error GoTo ErrorHandler
    Dim val As Currency
    val = CCur(strInput)
    On Error GoTo 0             'Disable ErrorHandler
    If val < MINUInt32 Or val > MAXUInt32 Then
        Err.Raise OverflowException
    End If
    'convert currency to ULong
    If val > MAXInt32 Then
        Parse.Value = val - NegativeInt32
    Else
        Parse.Value = val
    End If
    Exit Function
ErrorHandler:
    If Err.Number = OverflowException Then
        Err.Raise Err.Number, "ULong32.Parse", Err.Description
    Else
        Err.Raise ArgumentException, "ULong32.Parse", "Invalid argument"
    End If
End Function
#End If

''
'@Static
'@Description "Converts the numeric value of this instance to its equivalent string representation."
'@Parameters
'   val ULong
'       value to be converted into a string representation of its value.
'
'@Returns String
'   The string representation of the value of this instance, consisting of a sequence of digits
'   ranging from 0 to 9, without a sign or leading zeroes.
'
'@Remarks
'   Win64 converted ULong to LongLong type.
'   Win32 converted ULong to Currency type.
''
#If Win64 Then
Public Function ToString(ByRef val As ULong) As String
    Dim outputVal As LongLong
    If val.Value < MIN_ULONG Then
        outputVal = val.Value + NegativeInt32
    Else
        outputVal = val.Value
    End If
    ToString = CStr(outputVal)
End Function
#Else
Public Function ToString(ByRef val As ULong) As String
    Dim outputVal As Currency
    If val.Value < MIN_ULONG Then
        outputVal = val.Value + NegativeInt32
    Else
        outputVal = val.Value
    End If
    ToString = CStr(outputVal)
End Function
#End If

''
'@Static
'@Description "Returns a value indicating whether two unsigned 32-bit values are equal."
'@Parameters
'   lhs: ULong
'       unsigned  32-bit value to compare rhs.
'   rhs: ULong
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if values are equal.
''
Public Function Equals(ByRef lhs As ULong, ByRef rhs As ULong) As Boolean
    If (lhs.Value = rhs.Value) Then
        Equals = True
    End If
End Function

''
'@Description "Computes the number of bits that are set in a value."
'@Parameters
'   val ULong
'       The value whose set bits are to be counted.
'@Returns ULong
'   The number of set bits in value.
'
'@TODO Implement
''
Public Function PopCount(ByRef val As ULong) As ULong
    Err.Raise NotImplementedException, "ULong32.PopCount"
End Function

'------------------------------------------------------------------
'Equality Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Returns a boolean value indicating whether two unsigned 32-bit values are not equal."
'@Parameters
'   lhs: ULong
'       unsigned  32-bit value to compare rhs.
'   rhs: ULong
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if values are not equal.
''
Public Function NotEquals(ByRef lhs As ULong, ByRef rhs As ULong) As Boolean
    If (lhs.Value <> rhs.Value) Then
        NotEquals = True
    End If
End Function

'------------------------------------------------------------------
'Comparison Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is less than another."
'@Parameters
'   lhs: ULong
'       unsigned  32-bit value to compare rhs.
'   rhs: ULong
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is less than rhs value.
''
Public Function LessThan(ByRef lhs As ULong, ByRef rhs As ULong) As Boolean
    'If the most significant bit is different
    If (lhs.Value Xor rhs.Value) < 0 Then
        'If left most significant bit set it is the larger unsigned value
        If lhs.Value And MOST_SIGNIFICANT_BIT_MASK Then
             LessThan = False
             Exit Function
        Else
            LessThan = True
            Exit Function
        End If
    End If
    'If "same sign" i.e. the most significant bit are the same compare raw Int32 values
    If lhs.Value > rhs.Value Then
        LessThan = False
    ElseIf lhs.Value < rhs.Value Then
        LessThan = True
    End If
End Function

''
'@Static
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is less than or equal to another."
'@Parameters
'   lhs: ULong
'       unsigned  32-bit value to compare rhs.
'   rhs: ULong
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is less than or equal to rhs value.
''
Public Function LessThanOrEqual(ByRef lhs As ULong, ByRef rhs As ULong) As Boolean
    'If the most significant bit is different
    If (lhs.Value Xor rhs.Value) < 0 Then
        'If left most significant bit set it is the larger unsigned value
        If lhs.Value And MOST_SIGNIFICANT_BIT_MASK Then
             LessThanOrEqual = False
             Exit Function
        Else
            LessThanOrEqual = True
            Exit Function
        End If
    End If
    'If "same sign" i.e. the most significant bit are the same compare raw Int32 values
    If lhs.Value > rhs.Value Then
        LessThanOrEqual = False
    ElseIf lhs.Value < rhs.Value Then
        LessThanOrEqual = True
    Else
        LessThanOrEqual = True
    End If
End Function

''
'@Static
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is greater than another."
'@Parameters
'   lhs: ULong
'       unsigned  32-bit value to compare rhs.
'   rhs: ULong
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is greater than rhs value.
''
Public Function GreaterThan(ByRef lhs As ULong, ByRef rhs As ULong) As Boolean
    'If the most significant bit is different
    If (lhs.Value Xor rhs.Value) < 0 Then
        'If left most significant bit set it is the larger unsigned value
        If lhs.Value And MOST_SIGNIFICANT_BIT_MASK Then
             GreaterThan = True
             Exit Function
        Else
            GreaterThan = False
            Exit Function
        End If
    End If
    'If "same sign" i.e. the most significant bit are the same compare raw Int32 values
    If lhs.Value > rhs.Value Then
        GreaterThan = True
    ElseIf lhs.Value < rhs.Value Then
        GreaterThan = False
    End If
End Function

''
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is greater than or equal to another."
'@Parameters
'   lhs: ULong
'       unsigned  32-bit value to compare rhs.
'   rhs: ULong
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is greater than or equal to rhs value.
''
Public Function GreaterThanOrEqual(ByRef lhs As ULong, ByRef rhs As ULong) As Boolean
    'If the most significant bit is different
    If (lhs.Value Xor rhs.Value) < 0 Then
        'If left most significant bit set it is the larger unsigned value
        If lhs.Value And MOST_SIGNIFICANT_BIT_MASK Then
             GreaterThanOrEqual = True
             Exit Function
        Else
            GreaterThanOrEqual = False
            Exit Function
        End If
    End If
    'If "same sign" i.e. the most significant bit are the same compare raw Int32 values
    If lhs.Value > rhs.Value Then
        GreaterThanOrEqual = True
    ElseIf lhs.Value < rhs.Value Then
        GreaterThanOrEqual = False
    Else
        GreaterThanOrEqual = True
    End If
End Function

'------------------------------------------------------------------
'Math Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Adds two unsigned 32-bit values."
'@Parameters
'   lhs ULong
'       value to be added to rhs
'   rhs ULong
'       value to be addded  to lhs
'
'@Returns ULong
'   The result of the addition of two unsigned 32 bit values
'
'@Exceptions
'   OverflowException
'       Raised when a result exceeds the max UInt32 value of 4294967295
'
'@Remarks
'   For Win64 the addition calculation uses the LongLong data type.
'   For Win32 the addition calculation uses the Currency data type.
''
#If Win64 Then
Public Function Add(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' Perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Add.Value = rhs.Value
        Exit Function
    ElseIf rhs.Value = 0 Then
        Add.Value = lhs.Value
        Exit Function
    End If
    
    'Convert lhs and rhs to LongLong
    Dim lnglngLhs As LongLong
    If lhs.Value < MIN_ULONG Then
        lnglngLhs = lhs.Value + NegativeInt32
    Else
        lnglngLhs = lhs.Value
    End If
    Dim lnglngRhs As LongLong
    If rhs.Value < MIN_ULONG Then
        lnglngRhs = rhs.Value + NegativeInt32
    Else
        lnglngRhs = rhs.Value
    End If
    
    Dim lnglngResult As LongLong
    lnglngResult = lnglngLhs + lnglngRhs
    If lnglngResult > MAXUInt32 Then
       Err.Raise OverflowException, "ULong32.Add"
    End If
    
    'Convert result to ULong
    If lnglngResult > MAXInt32 Then
        Add.Value = CLng(lnglngResult - NegativeInt32)
    Else
        Add.Value = CLng(lnglngResult)
    End If
End Function
#Else
Public Function Add(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' Perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Add.Value = rhs.Value
        Exit Function
    ElseIf rhs.Value = 0 Then
        Add.Value = lhs.Value
        Exit Function
    End If
    
    'Convert ULong lhs and ULong rhs to Currency
    Dim curLhs As Currency
    If lhs.Value < MIN_ULONG Then
        curLhs = lhs.Value + NegativeInt32
    Else
        curLhs = lhs.Value
    End If
    Dim curRhs As Currency
    If rhs.Value < MIN_ULONG Then
        curRhs = rhs.Value + NegativeInt32
    Else
        curRhs = rhs.Value
    End If
    
    Dim curResult As Currency
    curResult = curLhs + curRhs
    If curResult > MAXUInt32 Then
       Err.Raise OverflowException, "ULong32.Add"
    End If
    
    'Convert result to ULong
    If curResult > MAXInt32 Then
        Add.Value = CLng(curResult - NegativeInt32)
    Else
        Add.Value = CLng(curResult)
    End If
End Function
#End If

''
'@Static
'@Description "Divides two unsigned 32-bit values."
'@Parameters
'   dividend UInt32
'       The dividend to be divided by the divisor
'   divisor UInt32
'       The divisor to divide the dividend
'
'@Returns ULong
'   The quotient of the of the dividend and divisor.
'
'@Exceptions
'   DivideByZeroException
'       Raised when a divisor is 0, runtime error 11
'
'@Remarks
'   Decimal places are rounded.
'   If the fraction is less than or equal to .5, the result will round down.
'   If the fraction is greater than .5, the result will round up.
'
'   For Win64 the divide calculation uses the LongLong data type.
'   For Win32 the divide calculation uses the Currency data type.
''
#If Win64 Then
Public Function Divide(ByRef dividend As ULong, ByRef divisor As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "ULong32.Divide"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        Exit Function
    ElseIf divisor.Value = 1 Then
        Divide.Value = dividend.Value
        Exit Function
    End If
    
    'Convert lhs and rhs to LongLong
    Dim lnglngDividend As LongLong
    If dividend.Value < MIN_ULONG Then
        lnglngDividend = dividend.Value + NegativeInt32
    Else
        lnglngDividend = dividend.Value
    End If
    
    Dim lnglngDivisor As LongLong
    If divisor.Value < MIN_ULONG Then
        lnglngDivisor = divisor.Value + NegativeInt32
    Else
        lnglngDivisor = divisor.Value
    End If
    
    Dim lnglngResult As LongLong
    lnglngResult = lnglngDividend / lnglngDivisor     'performing floating point division, decimal places are rounded
    'Convert result to ULong
    If lnglngResult > MAXInt32 Then
        Divide.Value = CLng(lnglngResult - NegativeInt32)
    Else
        Divide.Value = CLng(lnglngResult)
    End If
End Function
#Else
Public Function Divide(ByRef dividend As ULong, ByRef divisor As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "ULong32.Divide"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        Exit Function
    ElseIf divisor.Value = 1 Then
        Divide.Value = dividend.Value
        Exit Function
    End If
    
    'Convert lhs and rhs to LongLong
    Dim curDividend As Currency
    If dividend.Value < MIN_ULONG Then
        curDividend = dividend.Value + NegativeInt32
    Else
        curDividend = dividend.Value
    End If
    
    Dim curDivisor As Currency
    If divisor.Value < MIN_ULONG Then
        curDivisor = divisor.Value + NegativeInt32
    Else
        curDivisor = divisor.Value
    End If
    
    Dim curResult As Currency
    curResult = curDividend / curDivisor     'performing floating point division, decimal places are rounded
    'Convert result to ULong
    If curResult > MAXInt32 Then
        Divide.Value = CLng(curResult - NegativeInt32)
    Else
        Divide.Value = CLng(curResult)
    End If
End Function
#End If

''
'@Static
'@Description "Multiplies two unsigned 32-bit values."
'@Parameters
'   lhs ULong
'   rhs ULong
'
'@Returns ULong
'
'@Exceptions
'   OverflowException
'       Raised when the result exceeds the max unsigned 32-bit value of 4294967295
'@Remarks
'   For Win64 the multiply calculation uses the LongLong data type.
'   For Win32 the multiply calculation uses the Currency data type.
''
#If Win64 Then
Public Function Multiply(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' Perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Exit Function
    ElseIf rhs.Value = 0 Then
        Exit Function
    End If
    
    'Convert lhs and rhs to LongLong
    Dim lnglngLhs As LongLong
    If lhs.Value < MIN_ULONG Then
        lnglngLhs = lhs.Value + NegativeInt32
    Else
        lnglngLhs = lhs.Value
    End If
    Dim lnglngRhs As LongLong
    If rhs.Value < MIN_ULONG Then
        lnglngRhs = rhs.Value + NegativeInt32
    Else
        lnglngRhs = rhs.Value
    End If
    
    Dim lnglngResult As LongLong
    On Error GoTo ExceptionHandler
    lnglngResult = lnglngLhs * lnglngRhs
    On Error GoTo 0                                     'Disable ExceptionHandler handler.
    If lnglngResult > MAXUInt32 Then
       Err.Raise OverflowException, "ULong32.Multiply"
    End If
    
    'Convert result to ULong
    If lnglngResult > MAXInt32 Then
        Multiply.Value = CLng(lnglngResult - NegativeInt32)
    Else
        Multiply.Value = CLng(lnglngResult)
    End If
Exit Function
ExceptionHandler:
    Err.Raise OverflowException, "ULong32.Multiply"
End Function

#Else
Public Function Multiply(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' Perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Exit Function
    ElseIf rhs.Value = 0 Then
        Exit Function
    End If
    
    'Convert lhs and rhs to LongLong
    Dim curLhs As Currency
    If lhs.Value < MIN_ULONG Then
        curLhs = lhs.Value + NegativeInt32
    Else
        curLhs = lhs.Value
    End If
    Dim curRhs As Currency
    If rhs.Value < MIN_ULONG Then
        curRhs = rhs.Value + NegativeInt32
    Else
        curRhs = rhs.Value
    End If
    
    Dim curResult As Currency
    On Error GoTo ExceptionHandler
    curResult = curLhs * curRhs
    On Error GoTo 0             'Disable ExceptionHandler
    If curResult > MAXUInt32 Then
       Err.Raise OverflowException, "ULong32.Multiply"
    End If
    
    'Convert result to ULong
    If curResult > MAXInt32 Then
        Multiply.Value = CLng(curResult - NegativeInt32)
    Else
        Multiply.Value = CLng(curResult)
    End If
Exit Function
ExceptionHandler:
    Err.Raise OverflowException, "ULong32.Multiply"
End Function
#End If

''
'@Static
'@Description "Subtracts two unsigned 32 bit values."
'@Parameters
'   lhs ULong
'       Value to subtracted by rhs.
'   rhs ULong
'       Value to subtract from lhs.
'@Returns ULong
'
'@Exceptions
'   OverflowException
'       Raised when the result is less the minimum UInt32 value of 0.
'
'@Remarks
'   For Win64 the subtraction calculation uses the LongLong data type.
'   For Win32 the subtraction calculation uses the Currency data type.
''
#If Win64 Then
Public Function Subtract(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' Perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Subtract.Value = rhs.Value
        Exit Function
    ElseIf rhs.Value = 0 Then
        Subtract.Value = lhs.Value
        Exit Function
    End If
    
    'Convert lhs and rhs to LongLong
    Dim lnglngLhs As LongLong
    If lhs.Value < MIN_ULONG Then
        lnglngLhs = lhs.Value + NegativeInt32
    Else
        lnglngLhs = lhs.Value
    End If
    Dim lnglngRhs As LongLong
    If rhs.Value < MIN_ULONG Then
        lnglngRhs = rhs.Value + NegativeInt32
    Else
        lnglngRhs = rhs.Value
    End If
    
    Dim lnglngResult As LongLong
    lnglngResult = lnglngLhs - lnglngRhs
    If lnglngResult < MINUInt32 Then
       Err.Raise OverflowException, "ULong32.Subtract"
    End If
    
    'Convert result to ULong
    If lnglngResult > MAXInt32 Then
        Subtract.Value = CLng(lnglngResult - NegativeInt32)
    Else
        Subtract.Value = CLng(lnglngResult)
    End If
End Function

#Else
Public Function Subtract(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' Perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Subtract.Value = rhs.Value
        Exit Function
    ElseIf rhs.Value = 0 Then
        Subtract.Value = lhs.Value
        Exit Function
    End If
    
    'Convert ULong lhs and ULong rhs to Currency
    Dim curLhs As Currency
    If lhs.Value < MIN_ULONG Then
        curLhs = lhs.Value + NegativeInt32
    Else
        curLhs = lhs.Value
    End If
    Dim curRhs As Currency
    If rhs.Value < MIN_ULONG Then
        curRhs = rhs.Value + NegativeInt32
    Else
        curRhs = rhs.Value
    End If
    
    Dim curResult As Currency
    curResult = curLhs - curRhs
    If curResult < MINUInt32 Then
       Err.Raise OverflowException, "ULong32.Subtract"
    End If
    
    'Convert result to ULong
    If curResult > MAXInt32 Then
        Subtract.Value = CLng(curResult - NegativeInt32)
    Else
        Subtract.Value = CLng(curResult)
    End If
End Function
#End If

'------------------------------------------------------------------
'Bitwise Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Bitwise AND of two unsigned 32-bit value."
'@Parameters
'   lhs ULong
'   rhs ULong
'
'@Returns ULong
''
Public Function BitwiseAnd(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    BitwiseAnd.Value = lhs.Value And rhs.Value
End Function

''
'@Static
'@Description "Bitwise NOT for an unsigned 32-bit value."
'@Parameters
'   val ULong
'
'@Returns ULong
''
Public Function BitwiseNot(ByRef val As ULong) As ULong
    BitwiseNot.Value = Not val.Value
End Function

''
'@Static
'@Description "Bitwise OR for two unsigned 32-bit value."
'@Parameters
'   lhs ULong
'   rhs ULong
'
'@Returns ULong
'
'@Remarks
''
Public Function BitwiseOr(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    BitwiseOr.Value = lhs.Value Or rhs.Value
End Function

''
'@Static
'@Description "Bitwise XOR for two unsigned 32-bit values."
'@Parameters
'   lhs ULong
'   rhs ULong
'
'@Returns ULong
'
'@Remarks
''
Public Function BitwiseXor(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    BitwiseXor.Value = lhs.Value Xor rhs.Value
End Function

'------------------------------------------------------------------
'Shift Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Bitwise logical left-shift operation, by the given number of bits."
'@Parameters
'   val: ULong value
'   shiftAmount: The number of bits to left-shift by.
'@Returns ULong
'   Returns a Int32 of the bitwise logical left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
'   If shiftAmount <= 0 the value is returned
'   If the shiftAmount >= the maximum bitshift i.e. 32 then 0 is returned
''
Public Function ShiftLeft(ByRef val As ULong, ByVal shiftAmount As Long) As ULong
    If shiftAmount > MIN_BITSHIFT And shiftAmount < SIZEOF_UINT32_BITS Then
        ShiftLeft.Value = (val.Value And this.Masks(31 - shiftAmount)) * this.Powers(shiftAmount)
        If val.Value And this.Powers(31 - shiftAmount) Then
            ShiftLeft.Value = ShiftLeft.Value Or MOST_SIGNIFICANT_BIT_MASK  ' set bit 31
        End If
    ElseIf shiftAmount <= 0 Then
        ShiftLeft.Value = val.Value
    ElseIf shiftAmount >= MAX_BITSHIFT Then
        ShiftLeft.Value = &H0
    End If
End Function

''
'@Static
'@Description "Bitwise logical right-shift operation, by the given number of bits."
'@Parameters
'   val: ULong value
'   shiftAmount: The number of bits to right-shift by.
'@Returns ULong
'   Returns a Int32 of the bitwise logical right-shift operation, by the given number of bits.
'@Remarks
'   Right shifting is equal to dividing value by 2^Shift.
''
Public Function ShiftRight(ByRef val As ULong, ByVal shiftAmount As Long) As ULong
    If shiftAmount > MIN_BITSHIFT And shiftAmount < SIZEOF_UINT32_BITS Then
        ShiftRight.Value = ((val.Value And this.RevMasks(shiftAmount)) \ this.Powers(shiftAmount)) _
                    And this.Masks(MAX_BITSHIFT - shiftAmount)
    ElseIf shiftAmount = MIN_BITSHIFT Then
        ShiftRight.Value = val.Value
    ElseIf shiftAmount >= MAX_BITSHIFT Then
        ShiftRight.Value = &H0
    End If
End Function

''
'@Static
'@Description "Rotates the specified value left by the specified number of bits."
'"
'@Parameters
'   val ULong value
'   shiftAmount The number of bits to rotate-left by.
'@Returns ULong
'   Returns a ULong of the bitwise rotate-left, by the given number of bits.
'
'@Remarks
'   Bitwise operatation shifts each bit to the left, with the highest bit copied in the Carry flag
'   and into the lowest bit.
'
'@References
'   https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Numerics/BitOperations.cs
''
Public Function RotateLeft(ByRef val As ULong, ByVal shiftAmount As Long) As ULong
    RotateLeft.Value = ShiftLeft(val, shiftAmount).Value Or ShiftRight(val, (SIZEOF_UINT32_BITS - shiftAmount)).Value
End Function

''
'@Static
'@Description "Rotates the specified value right by the specified number of bits."
'@Parameters
'   val ULong
'       value
'   shiftAmount
'       The number of bits to rotate-right by.
'@Returns ULong
'   Returns a UInt32 of the bitwise rotate-right, by the given number of bits.
'@Remarks
'   Bitwise operatation shifts each bit to the right, with the lowest bit copied in the Carry flag
'   and into the highest bit.
'
'@Reference
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/Numerics/BitOperations.cs,f37eec40a05d7095
''
Public Function RotateRight(ByRef val As ULong, ByVal shiftAmount As Long) As ULong
    RotateRight.Value = ShiftRight(val, shiftAmount).Value Or ShiftLeft(val, (SIZEOF_UINT32_BITS - shiftAmount)).Value
End Function

'------------------------------------------------------------------
'IConvertible implementation
'------------------------------------------------------------------

''
'@Static
'@Description "Converts a ULong/UInt32 to the VBA decimal type."
'@Parameters
'   val ULong
'@Returns Variant
'   Variant containing a VBA decimal type.
''
Public Function ToDecimal(ByRef val As ULong) As Variant
    If val.Value < MIN_ULONG Then
        ToDecimal = CDec(val.Value) + NegativeInt32
    Else
        ToDecimal = CDec(val.Value)
    End If
End Function

''
'@Static
'@Description "Converts an unsigned 32-bit value to LongLong value."
'@Parameters
'   val ULong
'@Returns LongLong
'
'@Remarks
' Copies the binary value of Long data type to LongLong without VBA numeric conversion.
''
#If Win64 Then
Public Function ToLongLong(ByRef val As ULong) As LongLong
    If val.Value < MIN_ULONG Then
        ToLongLong = val.Value + NegativeInt32
    Else
        ToLongLong = val.Value
    End If
End Function
#End If

''
'@Static
'@Description "Converts an unsigned 32-bit value to Currency value."
'@Parameters
'   val ULong
'@Returns Currency
''
Public Function ToCurrency(ByRef val As ULong) As Currency
    Const NegativeInt32 As Currency = 4294967296@ ' For converting to ULong for negative Currency values
    If val.Value < MIN_ULONG Then
        ToCurrency = val.Value + NegativeInt32
    Else
        ToCurrency = val.Value
    End If
End Function

'------------------------------------------------------------------
'Private Methods
'------------------------------------------------------------------

'2^i
Private Sub InitPowers()
    ReDim this.Powers(SIZEOF_UINT32_BITS)
    
    this.Powers(0) = &H1&         ' 00000000000000000000000000000001    1
    this.Powers(1) = &H2&         ' 00000000000000000000000000000010    2
    this.Powers(2) = &H4&         ' 00000000000000000000000000000100    4
    this.Powers(3) = &H8&         ' 00000000000000000000000000001000    8
    
    this.Powers(4) = &H10&        ' 00000000000000000000000000010000    16
    this.Powers(5) = &H20&        ' 00000000000000000000000000100000    32
    this.Powers(6) = &H40&        ' 00000000000000000000000001000000    64
    this.Powers(7) = &H80&        ' 00000000000000000000000010000000    128

    this.Powers(8) = &H100&       ' 00000000000000000000000100000000    256
    this.Powers(9) = &H200&       ' 00000000000000000000001000000000    512
    this.Powers(10) = &H400&      ' 00000000000000000000010000000000    1024
    this.Powers(11) = &H800&      ' 00000000000000000000100000000000    2048
    
    this.Powers(12) = &H1000&     ' 00000000000000000001000000000000    4096
    this.Powers(13) = &H2000&     ' 00000000000000000010000000000000    8192
    this.Powers(14) = &H4000&     ' 00000000000000000100000000000000    16384
    this.Powers(15) = &H8000&     ' 00000000000000001000000000000000    32768
    
    this.Powers(16) = &H10000     ' 00000000000000010000000000000000    65536
    this.Powers(17) = &H20000     ' 00000000000000100000000000000000    131072
    this.Powers(18) = &H40000     ' 00000000000001000000000000000000    262144
    this.Powers(19) = &H80000     ' 00000000000010000000000000000000    524288
    
    this.Powers(20) = &H100000    ' 00000000000100000000000000000000    1048576
    this.Powers(21) = &H200000    ' 00000000001000000000000000000000    2097152
    this.Powers(22) = &H400000    ' 00000000010000000000000000000000    4194304
    this.Powers(23) = &H800000    ' 00000000100000000000000000000000    8388608
    
    this.Powers(24) = &H1000000   ' 00000001000000000000000000000000    16777216
    this.Powers(25) = &H2000000   ' 00000010000000000000000000000000    33554432
    this.Powers(26) = &H4000000   ' 00000100000000000000000000000000    67108864
    this.Powers(27) = &H8000000   ' 00001000000000000000000000000000    134217728
    
    this.Powers(28) = &H10000000  ' 00010000000000000000000000000000    268435456
    this.Powers(29) = &H20000000  ' 00100000000000000000000000000000    536870912
    this.Powers(30) = &H40000000  ' 01000000000000000000000000000000    1073741824
    this.Powers(31) = &H80000000  ' 10000000000000000000000000000000    -2147483648
    this.Powers(32) = &H0&        ' 00000000000000000000000000000000    0
    
End Sub

'2^i-1
Private Sub InitMasks()
    ReDim this.Masks(SIZEOF_UINT32_BITS)

    this.Masks(0) = &H0&        ' 00000000000000000000000000000000  0
    this.Masks(1) = &H1&        ' 00000000000000000000000000000001  1
    this.Masks(2) = &H3&        ' 00000000000000000000000000000011  3
    this.Masks(3) = &H7&        ' 00000000000000000000000000000111  7

    this.Masks(4) = &HF&        ' 00000000000000000000000000001111  15
    this.Masks(5) = &H1F&       ' 00000000000000000000000000011111  31
    this.Masks(6) = &H3F&       ' 00000000000000000000000000111111  63
    this.Masks(7) = &H7F&       ' 00000000000000000000000001111111  127

    this.Masks(8) = &HFF&       ' 00000000000000000000000011111111  255
    this.Masks(9) = &H1FF&      ' 00000000000000000000000111111111  511
    this.Masks(10) = &H3FF&     ' 00000000000000000000001111111111  1023
    this.Masks(11) = &H7FF&     ' 00000000000000000000011111111111  2047

    this.Masks(12) = &HFFF&     ' 00000000000000000000111111111111  4095
    this.Masks(13) = &H1FFF&    ' 00000000000000000001111111111111  8191
    this.Masks(14) = &H3FFF&    ' 00000000000000000011111111111111  16383
    this.Masks(15) = &H7FFF&    ' 00000000000000000111111111111111  32767

    this.Masks(16) = &HFFFF&    ' 00000000000000001111111111111111  65535
    this.Masks(17) = &H1FFFF    ' 00000000000000011111111111111111  131071
    this.Masks(18) = &H3FFFF    ' 00000000000000111111111111111111  262143
    this.Masks(19) = &H7FFFF    ' 00000000000001111111111111111111  524287

    this.Masks(20) = &HFFFFF    ' 00000000000011111111111111111111  1048575
    this.Masks(21) = &H1FFFFF   ' 00000000000111111111111111111111  2097151
    this.Masks(22) = &H3FFFFF   ' 00000000001111111111111111111111  4194303
    this.Masks(23) = &H7FFFFF   ' 00000000011111111111111111111111  8388607

    this.Masks(24) = &HFFFFFF   ' 00000000111111111111111111111111  16777215
    this.Masks(25) = &H1FFFFFF  ' 00000001111111111111111111111111  33554431
    this.Masks(26) = &H3FFFFFF  ' 00000011111111111111111111111111  67108863
    this.Masks(27) = &H7FFFFFF  ' 00000111111111111111111111111111  134217727

    this.Masks(28) = &HFFFFFFF  ' 00001111111111111111111111111111  268435455
    this.Masks(29) = &H1FFFFFFF ' 00011111111111111111111111111111  536870911
    this.Masks(30) = &H3FFFFFFF ' 00111111111111111111111111111111  1073741823
    this.Masks(31) = &H7FFFFFFF ' 01111111111111111111111111111111  2147483647

    this.Masks(32) = &HFFFFFFFF ' 11111111111111111111111111111111  -1

End Sub

'-2^i
Private Sub InitRevMasks()
    ReDim this.RevMasks(SIZEOF_UINT32_BITS)

    this.RevMasks(0) = &HFFFFFFFF   ' 11111111111111111111111111111111    -1
    this.RevMasks(1) = &HFFFFFFFE   ' 11111111111111111111111111111110    -2
    this.RevMasks(2) = &HFFFFFFFC   ' 11111111111111111111111111111100    -4
    this.RevMasks(3) = &HFFFFFFF8   ' 11111111111111111111111111111000    -8

    this.RevMasks(4) = &HFFFFFFF0   ' 11111111111111111111111111110000    -16
    this.RevMasks(5) = &HFFFFFFE0   ' 11111111111111111111111111100000    -32
    this.RevMasks(6) = &HFFFFFFC0   ' 11111111111111111111111111000000    -64
    this.RevMasks(7) = &HFFFFFF80   ' 11111111111111111111111110000000    -128

    this.RevMasks(8) = &HFFFFFF00   ' 11111111111111111111111100000000    -256
    this.RevMasks(9) = &HFFFFFE00   ' 11111111111111111111111000000000    -512
    this.RevMasks(10) = &HFFFFFC00  ' 11111111111111111111110000000000    -1024
    this.RevMasks(11) = &HFFFFF800  ' 11111111111111111111100000000000    -2048

    this.RevMasks(12) = &HFFFFF000  ' 11111111111111111111000000000000    -4096
    this.RevMasks(13) = &HFFFFE000  ' 11111111111111111110000000000000    -8192
    this.RevMasks(14) = &HFFFFC000  ' 11111111111111111100000000000000    -16384
    this.RevMasks(15) = &HFFFF8000  ' 11111111111111111000000000000000    -32768

    this.RevMasks(16) = &HFFFF0000  ' 11111111111111110000000000000000    -65536
    this.RevMasks(17) = &HFFFE0000  ' 11111111111111100000000000000000    -131072
    this.RevMasks(18) = &HFFFC0000  ' 11111111111111000000000000000000    -262144
    this.RevMasks(19) = &HFFF80000  ' 11111111111110000000000000000000    -524288

    this.RevMasks(20) = &HFFF00000  ' 11111111111100000000000000000000    -1048576
    this.RevMasks(21) = &HFFE00000  ' 11111111111000000000000000000000    -2097152
    this.RevMasks(22) = &HFFC00000  ' 11111111110000000000000000000000    -4194304
    this.RevMasks(23) = &HFF800000  ' 11111111100000000000000000000000    -8388608

    this.RevMasks(24) = &HFF000000  ' 11111111000000000000000000000000    -16777216
    this.RevMasks(25) = &HFE000000  ' 11111110000000000000000000000000    -33554432
    this.RevMasks(26) = &HFC000000  ' 11111100000000000000000000000000    -67108864
    this.RevMasks(27) = &HF8000000  ' 11111000000000000000000000000000    -134217728

    this.RevMasks(28) = &HF0000000  ' 11110000000000000000000000000000    -268435456
    this.RevMasks(29) = &HE0000000  ' 11100000000000000000000000000000    -536870912
    this.RevMasks(30) = &HC0000000  ' 11000000000000000000000000000000    -1073741824
    this.RevMasks(31) = &H80000000  ' 10000000000000000000000000000000    -2147483648
    
    this.RevMasks(32) = &H0&        ' 00000000000000000000000000000000      0
End Sub

''
'Converts Checked Decimal to ULong
''
Private Function DecimalToULong(ByVal val As Variant) As ULong
    If val > MAXInt32 Then
        DecimalToULong.Value = CLng(Fix(val) - NegativeInt32)
    Else
        DecimalToULong.Value = CLng(Fix(val))
    End If
End Function

''
'Converts Checked Currency to ULong
''
Private Function CurrencyToULong(ByVal val As Currency) As ULong
    Const MAXInt32 As Currency = 2147483647@
    Const NegativeInt32 As Currency = 4294967296@ ' Convert negative Currency values to positive ULong value

    If val > MAXInt32 Then
        CurrencyToULong.Value = CLng(Fix(val) - NegativeInt32)
    Else
        CurrencyToULong.Value = CLng(Fix(val))
    End If
End Function

''
'Converts Checked Double to ULong
''
Private Function DoubleToULong(ByVal val As Double) As ULong
    Const MAXInt32 As Double = 2147483647#
    Const NegativeInt32 As Double = 4294967296# ' Convert negative Currency values to positive ULong value

    If val > MAXInt32 Then
        DoubleToULong.Value = CLng(Fix(val) - NegativeInt32)
    Else
        DoubleToULong.Value = CLng(Fix(val))
    End If
End Function

''
'Converts Checked Single to ULong
''
Private Function SingleToULong(ByVal val As Single) As ULong
    Const MAXInt32 As Single = 2147483647
    Const NegativeInt32 As Single = 4294967296# ' Convert negative Single values to positive ULong value
    
    If val > MAXInt32 Then
        SingleToULong.Value = CLng(Fix(val) - NegativeInt32)
    Else
        SingleToULong.Value = CLng(Fix(val))
    End If
End Function

''
'Converts Checked LongLong to ULong
''
#If Win64 Then
Private Function LongLongToULong(ByVal val As LongLong) As ULong
    If val > MAXInt32 Then
        LongLongToULong.Value = CLng(val - NegativeInt32)
    Else
        LongLongToULong.Value = CLng(val)
    End If
End Function
#End If

