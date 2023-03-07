VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UInt32Static"
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
'@Version v1.2 March 6, 2023
'@LastModified March 7, 2023

'@Dependencies
'   ULongType.bas
'   DWordType.bas
'   WordType.bas
'   QWordType.bas

'@Remarks
'   V1.2
'   Changed the behaviour of CUInt32 to raise overflow exception for negative values.
'   Added CBytesUInt32 function to copy the byte value. i.e. Negative values will be converted to
'   larger unsigned 32-bit values.
'   CBytesUInt32 CUInt or Parse functions are the recommended methods to cast a value to a ULong and
'   avoid directly using the ULong.Value property.

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.uint32?view=net-7.0
' https://source.dot.net/#System.Private.CoreLib/src/libraries/System.Private.CoreLib/src/System/UInt32.cs
' https://newtonexcelbach.com/2015/10/26/the-vba-decimal-data-type/
' https://www.vbforums.com/showthread.php?834827-The-Decimal-Data-Type
' https://stackoverflow.com/questions/52099277/what-is-the-largest-possible-decimal-value-in-vba
' https://wiki.sei.cmu.edu/confluence/display/c/INT30-C.+Ensure+that+unsigned+integer+operations+do+not+wrap
' https://en.wikipedia.org/wiki/Division_algorithm

'@TODO
' Add functions for operators <, <=, >, >= currently could use CompareTo
' Add private types for DWORD, WORD, QWORD to remove dependencies?
' Unit testing

Option Explicit

'------------------------------------------------------------------
'Constants
'------------------------------------------------------------------
Private Const CHAR_BIT As Long = 8
Private Const SIZEOF_UINT32 As Long = 4
Private Const SIZEOF_UINT32_BITS As Long = SIZEOF_UINT32 * CHAR_BIT

Private Const MAX_ULONG As String = "4294967295"
Private Const MIN_ULONG As Long = 0
Private Const MAX_UINT32 As Long = &HFFFFFFFF
Private Const MIN_UINT32 As Long = &H0
Private Const MAXVALUE_Int32 As Long = 2147483647

'Bit shifting constants
Private Const MIN_BITSHIFT As Long = 0
Private Const MAX_BITSHIFT As Long = 32
Private Const MOST_SIGNIFICANT_BIT_MASK As Long = &H80000000

'Error number constants
Private Const OverflowException As Long = &H6&
Private Const ArgumentException As Long = &H5&  'Invalid procedure call or argument.
Private Const DivideByZeroException As Long = &H11&
Private Const ArgumentNullException As Long = &H80004003
Private Const NotImplementedException As Long = 9999

'------------------------------------------------------------------
'Types
'------------------------------------------------------------------
Private Type QWORD_Currency
    Value As Currency
End Type

Private Type TUInt32Static
    decMaxUInt32 As Variant 'contains a decimal type

    ulngOne As ULong    'For caculation in IsPow2
    Powers() As Long    'Used for bitwise operations
    Masks() As Long     'Used for bitwise operations
    RevMasks() As Long  'Used for bitwise operations
End Type

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As TUInt32Static

'------------------------------------------------------------------
'Constructors and destructors
'------------------------------------------------------------------
Private Sub Class_Initialize()
    this.decMaxUInt32 = CDec(MAX_ULONG)
    this.ulngOne.Value = 1 'constant for caculation in IsPow2
    
    'Initialise bitwise arrays
    InitPowers
    InitMasks
    InitRevMasks
End Sub

'------------------------------------------------------------------
'Properties
'------------------------------------------------------------------

''
'@Description"Represents the largest possible value of UInt32. This field is constant."
'@Field Value ULong
'   Value = 4294967295#
'@Remarks
'   The value of this constant is 4,294,967,295; that is, hexadecimal 0xFFFFFFFF.
''
Public Property Get MaxValue() As ULong
    MaxValue.Value = MAX_UINT32
End Property

''
'@Description"Represents the smallest possible value of UInt32. This field is constant."
'@Field Value UInt32
'   Value = 0
'@Remarks
'   The value of this constant is 0.
''
Public Property Get MinValue() As ULong
    MinValue.Value = MIN_UINT32
End Property

Friend Property Get MaxValueDecimal() As Variant
    MaxValueDecimal = this.decMaxUInt32
End Property

'------------------------------------------------------------------
'Public Methods
'------------------------------------------------------------------

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
'@Description "Converts the byte value into an unsigned 32-bit integer."
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
'       Raised for when not a valid type to cast of a Byte, Integer, LongLong or Currency
'
'@Remarks
'   Bytes values to be converted maybe of types Byte, Long, Integers or Currency.
'   Negative values are converted into a larger unsigned 32-bit integers.
'
'   For Integer type its WORD value is copied to the ULong DWORD lower WORD.
'   eg. If Integer value is -1 Hex FFFF is converted to ULong value of 65535 Hex 0000FFFF
'
'   For Long type its DWORD value is copied to the ULong DWORD
'   eg. If Long value is -1 Hex FFFFFFFF is converted to ULong value of 4294967295 Hex FFFFFFFF
'
'
'   For Currency type the low DWORD of a currency value is copied to the DWORD of the ULong.
'   Eg Currency value of 0.0001 Hex 00000000 00000001 converts to ULong 00000001 i.e. of value 1.
'
'   For LongLong type the low DWORD of a LongLong value is copied to the DWORD of the ULong.
'   eg. If LongLong value is 42949672958 i.e. Hex 00000009 FFFFFFFE is converted to ULong value Hex FEFFFFFF
'   i.e. value of 4294967294
''
Public Function CBytesUInt32(ByVal val As Variant) As ULong
    Select Case VarType(val)
        Case VBA.vbLong
            CBytesUInt32.Value = val
        Case VBA.vbByte
            CBytesUInt32.Value = val
        Case VBA.vbInteger
            'copy Integer type WORD to lower WORD of ULong
            Dim wVal As WORD
            wVal.Value = val
            LSet CBytesUInt32 = wVal
        #If Win64 Then
        Case VBA.vbLongLong
            'copy low DWORD of LongLong type to ULong
            Dim qwVal As QWORD
            qwVal.Value = val
            LSet CBytesUInt32 = qwVal
        #End If
        Case VBA.vbCurrency
            'copy low DWORD of Currency type to ULong
            Dim qwValCurrency As QWORD_Currency
            qwValCurrency.Value = val
            LSet CBytesUInt32 = qwValCurrency
        Case Else
            Err.Raise ArgumentException, "UInt32Static.CBytesUInt32", "Invalid argument"
    End Select
End Function

''
'@Static
'@Description "Converts a value to unsigned 32-bit integer."
'@Parameters
'   val: Variant
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
'   Decimal values are truncated
'   Negative values return an overflow exception
''
Public Function CUInt32(ByRef val As Variant) As ULong
    Select Case VarType(val)
        Case VBA.vbLong
            If val < 0 Then
                Err.Raise OverflowException, "UInt32Static.CUInt32"
            Else
                CUInt32.Value = val
            End If
        Case VBA.vbString
            CUInt32 = Me.Parse(CStr(val))
        Case VBA.vbByte
            CUInt32.Value = val
        Case VBA.vbInteger
            If val < 0 Then
                Err.Raise OverflowException, "UInt32Static.CUInt32"
            Else
                CUInt32.Value = val
            End If
        Case VBA.vbCurrency
            If val < 0 Or val > 4294967295# Then
                Err.Raise OverflowException, "UInt32Static.CUInt32"
            Else
                CUInt32 = DecimalToULong(val)
            End If
        Case VBA.vbDecimal
            If val < 0 Or val > this.decMaxUInt32 Then
                Err.Raise OverflowException, "UInt32Static.CUInt32"
            Else
                CUInt32 = DecimalToULong(val)
            End If
        #If Win64 Then
        Case VBA.vbLongLong
            If val < 0 Or val > 4294967295# Then
                Err.Raise OverflowException, "UInt32Static.CUInt32"
            Else
                Dim qwVal As QWORD
                qwVal.Value = val
                LSet CUInt32 = qwVal
            End If
        #End If
        Case VBA.vbDouble
            If val < 0 Or val > 4294967295# Then
                Err.Raise OverflowException, "UInt32Static.CUInt32"
            Else
                CUInt32 = DecimalToULong(val)
            End If
        Case VBA.vbSingle
            If val < 0 And val > 4294967295# Then
                Err.Raise OverflowException, "UInt32Static.CUInt32"
            Else
                CUInt32 = DecimalToULong(val)
            End If
        Case Else
            Err.Raise ArgumentException, "UInt32Static.CUInt32", "Invalid argument"
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
'@Description"Returns a value indicating whether two unsigned 32-bit values are equal."
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
    Absolute = val
End Function

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
'@Remarks
'
'@TODO Implement
''
Public Function Clamp(ByRef val As ULong, ByRef valMin As ULong, ByRef valMax As ULong) As ULong
'public static uint Clamp (uint value, uint min, uint max)
    Err.Raise NotImplementedException, "UInt32.Clamp"
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
'
'@Remarks
''
#If Win64 Then
Public Function DivRem(ByRef dividend As ULong, ByRef divisor As ULong, ByRef outRemainder As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "UInt32Static.DivRem"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        outRemainder.Value = 0
        Exit Function
    ElseIf divisor.Value = 1 Then
        LSet DivRem = dividend
        outRemainder.Value = 0
        Exit Function
    End If

    Dim qwDividend As QWORD
    Dim qwDivisor As QWORD
    Dim qwQuotient As QWORD
    Dim qwRemainder As QWORD

    LSet qwDividend = dividend
    LSet qwDivisor = divisor

    qwQuotient.Value = qwDividend.Value \ qwDivisor.Value 'Using integer division which truncates decimal places
    qwRemainder.Value = qwDividend.Value - (qwDivisor.Value * qwQuotient.Value)
    LSet DivRem = qwQuotient
    LSet outRemainder = qwRemainder
End Function

#Else
Public Function DivRem(ByRef dividend As ULong, ByRef divisor As ULong, ByRef outRemainder As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "UInt32Static.DivRem"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        outRemainder.Value = 0
        Exit Function
    ElseIf divisor.Value = 1 Then
        LSet DivRem = dividend
        outRemainder.Value = 0
        Exit Function
    End If

    'convert dividend ULong to decimal
    Dim decimalDividend As Variant
    decimalDividend = CDec(dividend.Value)
    If decimalDividend < 0 Then
        decimalDividend = decimalDividend + 4294967296#
    End If

    'convert divisor ULong to decimal
    Dim decimalDivisor As Variant
    decimalDivisor = CDec(divisor.Value)
    If decimalDivisor < 0 Then
        decimalDivisor = decimalDivisor + 4294967296#
    End If

    Dim decimalQuotient As Variant
    decimalQuotient = Fix(decimalDividend / decimalDivisor)
    
    Dim decimalRemainder As Variant
    decimalRemainder = decimalDividend - (decimalDivisor * decimalQuotient)

    'Convert decimal remainder to ULong
    If decimalRemainder > MAXVALUE_Int32 Then
        outRemainder.Value = decimalRemainder - 4294967296#
    Else
        outRemainder.Value = decimalRemainder
    End If
    
    'Convert decimal quotient to ULong
    If decimalQuotient > MAXVALUE_Int32 Then
        DivRem.Value = decimalQuotient - 4294967296#
    Else
        DivRem.Value = decimalQuotient
    End If
End Function
#End If

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
        Err.Raise DivideByZeroException, "UInt32Static.Modulus"
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
'   left if it is greater than right; otherwise, right.
'
'@Remarks
''
Public Function Max(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    If CompareTo(lhs, rhs) > 0 Then
        Max = lhs
    Else
        Max = rhs
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
        Min = lhs
    Else
        Min = rhs
    End If
End Function

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
'   For Win64 using LongLong type for addition calculation.
'   Long binary value is copied to LongLong value without VBA numeric conversion using LSet.
''
#If Win64 Then
Public Function Add(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Add.Value = rhs.Value
        Exit Function
    ElseIf rhs.Value = 0 Then
        Add.Value = lhs.Value
        Exit Function
    End If

    Dim qwLhs As QWORD
    Dim qwRhs As QWORD
    LSet qwLhs = lhs
    LSet qwRhs = rhs
    
    Dim qwResult As QWORD
    qwResult.Value = qwLhs.Value + qwRhs.Value
    If qwResult.Value > 4294967295# Then
       Err.Raise OverflowException, "UInt32.Add"
    End If
    LSet Add = qwResult 'copy LongLong to Long without VBA numeric conversion which may cause overflow.
End Function

#Else
''
'@Remarks
'   For Win32 addition is performed by addition of low and high WORDs with carry..
''
Public Function Add(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    ' perform a couple of quick checks that have an easy result.
    If lhs.Value = 0 Then
        Add.Value = rhs.Value
        Exit Function
    ElseIf rhs.Value = 0 Then
        Add.Value = lhs.Value
        Exit Function
    End If
      
    'Copy lhs and rhs to obtain low high word values
    Dim dwlhs As DWORDLoHi
    LSet dwlhs = lhs
    Dim dwrhs As DWORDLoHi
    LSet dwrhs = rhs
    Dim dwResult As DWORDLoHi
        
    Dim result As ULong
    Dim wlhs As ULong
    Dim wrhs As ULong
    
    'Copy low word binary values without conversion to long type.
    LSet wlhs = dwlhs.wLow
    LSet wrhs = dwrhs.wLow
    result.Value = wlhs.Value + wrhs.Value
    If result.Value > 65535 Then
        dwResult.wHigh.Value = 1    'carry occured
    End If
    LSet dwResult.wLow = result 'copy low word of result without VBA conversion
    
    'Copy high word binary values without conversion to long type.
    LSet wlhs = dwlhs.wHigh
    LSet wrhs = dwrhs.wHigh
    result.Value = wlhs.Value + wrhs.Value + dwResult.wHigh.Value
    If result.Value > 65535 Then
        Err.Raise OverflowException, "UInt32.Add"
    End If
    LSet dwResult.wHigh = result 'copy low word of result without VBA conversion
    
    LSet Add = dwResult
End Function
#End If

''
'@Static
'@Description "Divides two unsigned 32-bit values."
'@Parameters
'   dividend ULong
'       The dividend
'   divisor ULong
'       The divisor
'
'@Returns ULong
'   The quotient of the of the dividend and divisor.
'
'@Exceptions
'   DivideByZeroException
'       Raised when a divisor is 0, runtime error 11
'
'@Remarks
'   Win64 version using QWORD approximately 40 times better performance.
'   Decimal places are truncated.
''
#If Win64 Then
Public Function Divide(ByRef dividend As ULong, ByRef divisor As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "UInt32Static.Divide"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        Exit Function
    ElseIf divisor.Value = 1 Then
        LSet Divide = dividend
        Exit Function
    End If
    
    Dim qwDividend As QWORD
    Dim qwDivisor As QWORD
    
    LSet qwDividend = dividend
    LSet qwDivisor = divisor
    
    Dim qwResult As QWORD
    'performing integer division decimal places truncated
    qwResult.Value = qwDividend.Value \ qwDivisor.Value
    
    LSet Divide = qwResult
End Function
#Else
Public Function Divide(ByRef dividend As ULong, ByRef divisor As ULong) As ULong
    If divisor.Value = 0 Then
        Err.Raise DivideByZeroException, "UInt32Static.Divide"
    End If
    ' perform a couple of quick checks that have an easy result.
    If dividend.Value = 0 Then
        Exit Function
    ElseIf divisor.Value = 1 Then
        LSet Divide = dividend
        Exit Function
    End If
        
    'convert dividend ULong to decimal
    Dim decimalDividend As Variant
    decimalDividend = CDec(dividend.Value)
    If decimalDividend < 0 Then
        decimalDividend = decimalDividend + 4294967296#
    End If
    
    'convert divisor ULong to decimal
    Dim decimalDivisor As Variant
    decimalDivisor = CDec(divisor.Value)
    If decimalDivisor < 0 Then
        decimalDivisor = decimalDivisor + 4294967296#
    End If
    
    Dim decimalQuotient As Variant
    decimalQuotient = Fix(decimalDividend / decimalDivisor)
    
    'Convert decimal result to ULong
    If decimalQuotient > MAXVALUE_Int32 Then
        Divide.Value = decimalQuotient - 4294967296#
    Else
        Divide.Value = decimalQuotient
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
'   For Win64 using LongLong type for multiply calculation.
''
#If Win64 Then
Public Function Multiply(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    Dim qwLhs As QWORD
    Dim qwRhs As QWORD
    
    LSet qwLhs = lhs
    LSet qwRhs = rhs
    
    Dim qwResult As QWORD
    qwResult.Value = qwLhs.Value * qwRhs.Value
    
    If qwResult.Value > 4294967295# Then
       Err.Raise OverflowException, "UInt32Static.Multiply"
    End If
    LSet Multiply = qwResult
End Function
#Else
''
'@Remarks
'   For Win32 using decimal type for multiply calculation.
''
Public Function Multiply(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    'convert lhs ULong to decimal
    Dim decimalLhs As Variant
    decimalLhs = CDec(lhs.Value)
    If decimalLhs < 0 Then
        decimalLhs = decimalLhs + 4294967296#
    End If
    
    'convert rhs ULong to decimal
    Dim decimalRhs As Variant
    decimalRhs = CDec(rhs.Value)
    If decimalRhs < 0 Then
        decimalRhs = decimalRhs + 4294967296#
    End If
    
    Dim decimalResult As Variant
    decimalResult = decimalLhs * decimalRhs
    If decimalResult > 4294967295# Then
       Err.Raise OverflowException, "UInt32Static.Multiply"
    End If
    
    'Convert decimal result to ULong
    If decimalResult > MAXVALUE_Int32 Then
        Multiply.Value = decimalResult - 4294967296#
    Else
        Multiply.Value = decimalResult
    End If
End Function
#End If

''
'@Static
'@Description "Subtracts two unsigned 32 bit values."
'@Parameters
'   lhs ULong
'       Value to subtract from
'   rhs ULong
'       Value to subtract from lhs
'@Returns ULong
'
'@Exceptions
'   OverflowException
'       Raised when the result is less the minimum UInt32 value of 0.
'
'@Remarks
'   For Win64 subtract is performed by casting binary ULong values to QWORD containing LongLong
'   values without VBA numeric conversion.
''
#If Win64 Then
Public Function Subtract(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    'Cast lhs, rhs binary value to LongLong without conversion
    Dim qwLhs As QWORD
    Dim qwRhs As QWORD
    LSet qwLhs = lhs
    LSet qwRhs = rhs
    
    'Subtract using LongLong type within QWORD
    Dim qwResult As QWORD
    qwResult.Value = qwLhs.Value - qwRhs.Value
    If qwResult.Value < MIN_UINT32 Then
       Err.Raise OverflowException, "UInt32Static.Subtract"
    End If
    LSet Subtract = qwResult
End Function

#Else
''
'@Remarks
'   For Win32 addition is performed by subtraction of low and high WORDs with carry.
'   Calculation of WORDs subtract is performed using Long.
''
Public Function Subtract(ByRef lhs As ULong, ByRef rhs As ULong) As ULong
    Const MINVALUE_WORD As Long = 0
    If lhs.Value = 0 Then
        Subtract.Value = rhs.Value
        Exit Function
    ElseIf rhs.Value = 0 Then
        Subtract.Value = lhs.Value
        Exit Function
    ElseIf Me.CompareTo(lhs, rhs) = -1 Then
        Err.Raise OverflowException, "UInt32Static.Subtract"
        Exit Function
    End If

    'Copy lhs and rhs to obtain low high word values
    Dim dwlhs As DWORDLoHi
    LSet dwlhs = lhs
    Dim dwrhs As DWORDLoHi
    LSet dwrhs = rhs
    Dim dwResult As DWORDLoHi
        
    Dim result As ULong
    Dim wlhs As ULong
    Dim wrhs As ULong
    
    'Copy low word binary values without conversion to long type.
    LSet wlhs = dwlhs.wLow
    LSet wrhs = dwrhs.wLow
    result.Value = wlhs.Value - wrhs.Value
    If result.Value < MINVALUE_WORD Then
        dwResult.wHigh.Value = 1   'carry occured
    End If
    LSet dwResult.wLow = result 'copy low word of result without VBA conversion
    
    'Copy high word binary values without conversion to long type.
    LSet wlhs = dwlhs.wHigh
    LSet wrhs = dwrhs.wHigh
    result.Value = wlhs.Value - wrhs.Value - dwResult.wHigh.Value
    If result.Value < MINVALUE_WORD Then
        Err.Raise OverflowException, "UInt32Static.Subtract"
    End If
    LSet dwResult.wHigh = result 'copy low word of result without VBA conversion
    
    LSet Subtract = dwResult
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
'@TODO Implement
''
Public Function LeadingZeroCount(ByRef val As ULong) As ULong
    Err.Raise NotImplementedException, "UInt32.LeadingZeroCount"
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
    Err.Raise NotImplementedException, "UInt32.PopCount"
End Function

''
'@Static
'@Description "Bitwise logical left-shift operation, by the given number of bits."
'@Parameters
'   val: ULong value
'   offset: The number of bits to left-shift by.
'@Returns Long
'   Returns a Int32 of the bitwise logical left-shift operation, by the given number of bits.
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
'   If offset <= 0 the value is returned
'   If the offset >= the maximum bitshift i.e. 32 then 0 is returned
''
Public Function LeftShift(ByRef val As ULong, ByVal offset As Long) As ULong
    If offset > MIN_BITSHIFT And offset < SIZEOF_UINT32_BITS Then
        LeftShift.Value = (val.Value And this.Masks(31 - offset)) * this.Powers(offset)
        If val.Value And this.Powers(31 - offset) Then
            LeftShift.Value = LeftShift.Value Or MOST_SIGNIFICANT_BIT_MASK  ' set bit 31
        End If
    ElseIf offset <= 0 Then
        LeftShift = val
    ElseIf offset >= MAX_BITSHIFT Then
        LeftShift.Value = &H0
    End If
End Function

''
'@Static
'@Description "Bitwise logical right-shift operation, by the given number of bits."
'@Parameters
'   val: ULong value
'   offset: The number of bits to right-shift by.
'@Returns Long
'   Returns a Int32 of the bitwise logical right-shift operation, by the given number of bits.
'@Remarks
'   Right shifting is equal to dividing value by 2^Shift.
Public Function RightShift(ByRef val As ULong, ByVal offset As Long) As ULong
    If offset > MIN_BITSHIFT And offset < SIZEOF_UINT32_BITS Then
        RightShift.Value = ((val.Value And this.RevMasks(offset)) \ this.Powers(offset)) _
                    And this.Masks(MAX_BITSHIFT - offset)
    ElseIf offset = MIN_BITSHIFT Then
        RightShift = val
    ElseIf offset >= MAX_BITSHIFT Then
        RightShift.Value = &H0
    End If
End Function

''
'@Static
'@Description "Rotates the specified value left by the specified number of bits."
'"
'@Parameters
'   val ULong value
'   offset The number of bits to rotate-left by.
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
Public Function RotateLeft(ByRef val As ULong, ByVal offset As Long) As ULong
    RotateLeft.Value = LeftShift(val, offset).Value Or RightShift(val, (SIZEOF_UINT32_BITS - offset)).Value
End Function

''
'@Static
'@Description "Rotates the specified value right by the specified number of bits."
'@Parameters
'   val ULong
'       value
'   offset
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
Public Function RotateRight(ByRef val As ULong, ByVal offset As Long) As ULong
    RotateRight.Value = RightShift(val, offset).Value Or LeftShift(val, (SIZEOF_UINT32_BITS - offset)).Value
End Function

''
'@Static
'@Description"Determines if a value is a power of two."
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
'   Win64 converted ULong to LongLong type using a QWORD.
'   Win32 converted ULong to decimal type.
''
#If Win64 Then
Public Function ToString(ByRef val As ULong) As String
    Dim qwVal As QWORD
    LSet qwVal = val    'copy binary Long value to longlong without VBA numeric conversion
    ToString = CStr(qwVal.Value)
End Function
#Else
Public Function ToString(ByRef val As ULong) As String
    Dim decimalVal As Variant
    decimalVal = CDec(val.Value)
    If decimalVal < 0 Then
        decimalVal = decimalVal + 4294967296#
    End If
    ToString = CStr(decimalVal)
End Function
#End If

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
'@TODO Check the behaviour of conversion of values with decimals are treated. Round or truncate?
''
#If Win64 Then
Public Function Parse(ByRef s As String) As ULong
    If (Trim(s) = VBA.vbNullString) Then
        Err.Raise ArgumentNullException, "UInt32Static.Parse", "Argument Null Exception"
    End If

    On Error GoTo ErrorHandler
    Dim qwVal As QWORD
    qwVal.Value = s
    If qwVal.Value < 0 Or qwVal.Value > 4294967295# Then
        Err.Raise OverflowException
    Else
        LSet Parse = qwVal
    End If
CleanExit:
    Exit Function
ErrorHandler:
    Err.Raise ArgumentException, "UInt32Static.Parse", "Invalid argument"
    Resume CleanExit
End Function

#Else
Public Function Parse(ByRef s As String) As ULong
    If (Trim(s) = VBA.vbNullString) Then
        Err.Raise ArgumentNullException, "UInt32Static.Parse"
    End If
    On Error GoTo ErrorHandler
    Dim decVal As Variant
    'To cast string to UInt32 convert to rounded decimal as internal values internally are as unsigned with a sign flag and Base 10 exponent
    decVal = Fix(CDec(s))
    If decVal < 0 Or decVal > this.decMaxUInt32 Then
        Err.Raise OverflowException
    Else
        Parse = DecimalToULong(decVal)
    End If
CleanExit:
    Exit Function
ErrorHandler:
    Err.Raise ArgumentException, "UInt32Static.Parse", "Invalid argument"
    Resume CleanExit
End Function
#End If

''
'@Static
'@Description "Converts a ULong/UInt32 to the VBA decimal type."
'@Parameters
'   val ULong
'@Returns Variant
'   Variant containing a VBA decimal type.
'
'@Remarks
''
#If Win64 Then
Public Function ToDecimal(ByRef val As ULong) As Variant
    Dim qwVal As QWORD
    LSet qwVal = val
    ToDecimal = CDec(qwVal.Value)
End Function
#Else
Public Function ToDecimal(ByRef val As ULong) As Variant
    ToDecimal = CDec(val.Value)
    If ToDecimal < 0 Then
        ToDecimal = ToDecimal + 4294967296#
    End If
End Function
#End If

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
    Dim qwVal As QWORD
    LSet qwVal = val
    ToLongLong = qwVal.Value
End Function
#End If

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

Private Function DecimalToULong(ByVal val As Variant) As ULong
    Dim decVal As Variant
    decVal = Fix(val)

    If decVal > MAXVALUE_Int32 Then
        DecimalToULong.Value = decVal - 4294967296#
    Else
        DecimalToULong.Value = decVal
    End If
End Function


