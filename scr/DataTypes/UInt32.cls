VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UInt32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rubberduck annotations
'@Folder("VBACorLib.DataTypes")
'@ModuleDescription "Represents a 32-bit unsigned integer."
'@PredeclaredId

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-Library
'@Version v1.0 March 8, 2023
'@LastModified March 10, 2023

'@Dependencies
'   ULongType.bas
'   ULong32.cls
'   UInt32Static.cls

'Remarks
' Unsigned 32-bit Integer ojbect.

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.uint32?view=net-7.0
' https://github.com/dotnet/docs/blob/main/docs/standard/generics/math.md

'@TODO
'Implement
'   ToBoolean
'   ToByte
'   ToInt16
'   ToInt32
'
'Performance improvements
'   possibly move in ULong32 logic, change friend property value to Long
'   move bitwise arrays to module
'
'ToInt16

Option Explicit

'------------------------------------------------------------------
'Constants
'------------------------------------------------------------------
'Error number constants
Private Const OverflowException As Long = &H6&
Private Const ArgumentException As Long = &H5&  'Invalid procedure call or argument.
Private Const DivideByZeroException As Long = &H11&
Private Const ArgumentNullException As Long = &H80004003
Private Const NotImplementedException As Long = 9999

'------------------------------------------------------------------
'Private Variables
'------------------------------------------------------------------
Private this As ULong

'------------------------------------------------------------------
'Constructors
'------------------------------------------------------------------

''
'@Static
'@Description "Initializes a new instance of an unsigned 32-bit integer to a specified value."
'@Parameters
'   val: Variant
'       value to be converted to an unsigned 32-bit value
'@Returns UInt32
'   value converted to an unsigned 32 bit integer.
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
'   Be careful using decimal places.
'   If the fraction is less than or equal to .5, the result will round down.
'   If the fraction is greater than .5, the result will round up.
''
Public Function Create(Optional ByVal val As Variant) As UInt32
    With New UInt32
        If Not IsMissing(val) Then
            .Value = ULong32.CreateChecked(val)
        End If
        Set Create = .Self
    End With
End Function

''
'@Static
'@Description "Initializes a new instance of an unsigned 32-bit integer to a specified ULong structure."
'@Parameters
'   val: Variant
'       value to be converted to an unsigned 32-bit value
'@Returns UInt32
'   value converted to an unsigned 32 bit integer.
'
''
Public Function CreateFromULong(ByRef val As ULong) As UInt32
    With New UInt32
        .Value = val
        Set CreateFromULong = .Self
    End With
End Function

'------------------------------------------------------------------
'Properties Internal
'------------------------------------------------------------------
Friend Property Get Self() As UInt32
    Set Self = Me
End Property

Friend Property Get Value() As ULong
    Value = this
End Property

Friend Property Let Value(ByRef val As ULong)
    this = val
End Property

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
Public Property Get MaxValue() As UInt32
    Set MaxValue = UInt32Static.MaxValue
End Property

''
'@Description"Represents the smallest possible value of UInt32. This field is constant."
'@Field Value UInt32
'   Value = 0
'@Remarks
'   The value of this constant is 0.
''
Public Property Get MinValue() As UInt32
    Set MinValue = UInt32Static.MinValue
End Property

'------------------------------------------------------------------
'Public Functions
'------------------------------------------------------------------

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
Public Function Clamp(ByVal val As UInt32, ByRef minVal As UInt32, ByRef maxVal As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Clamp(val.Value, minVal.Value, maxVal.Value)
        Set Clamp = .Self
    End With
End Function

''
'@Description "Compares this instance to a specified 32-bit unsigned integer and returns an indication of their relative values."
'@Parameters
'   Value UInt32
'       An unsigned integer to compare.
'
'@Returns Long
'   A signed number indicating the relative values of this instance and value.
'
'   Return value        Description
'   Less than zero      This instance is less than value.
'   Zero                This instance is equal to value.
'   Greater than zero   This instance is greater than value.
''
Public Function CompareTo(ByVal val As UInt32) As Long
    CompareTo = ULong32.CompareTo(this, val.Value)
End Function

''
'@Static
'@Description "Creates an instance of the current type from a value, throwing an overflow exception for any values that fall outside the representable range of the current type."
'@Parameters
'   Value Variant
'       The value which is used to create the instance of an unsigned 32-bit integer.
'
'@Returns UInt32
'   An instance of an unsigned 32-bit integer created from value.
''
Public Function CreateChecked(ByVal val As Variant) As UInt32
    With New UInt32
        .Value = ULong32.CreateChecked(val)
        Set CreateChecked = .Self
    End With
End Function

''
'@Static
'@Description "Creates an instance of the current type from a value, saturating any values that fall outside the representable range of the current type."
'   val Variant
'       The value which is used to create the instance of an unsigned 32-bit integer.
'@Remarks
'   Creates a value, clamping to T.MinValue or T.MaxValue if the input can't fit.
''
Public Function CreateSaturating(ByVal val As Variant) As UInt32
    With New UInt32
        .Value = ULong32.CreateSaturating(val)
        Set CreateSaturating = .Self
    End With
End Function

''
'@Static
'@Description "Creates an instance of the current type from a value, truncating any values that fall outside the representable range of the current type."
'@Parameters
'   val Variant
'       The value to be converted to an unsigned 32-bit integer.
'       The byte value is copied to an unsigned 32-bit integer where negative values are converted
'       into larger unsigned values.
'@Returns UInt32
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
'@References
' https://github.com/dotnet/docs/blob/main/docs/standard/generics/math.md
''
Public Function CreateTruncating(ByVal val As Variant) As UInt32
    With New UInt32
        .Value = ULong32.CreateTruncating(val)
        Set CreateTruncating = .Self
    End With
End Function

''
'@Static
'@Description "Computes the quotient and remainder of two values."
'@Parameters
'   dividend UInt32
'       The value which divisor divides.
'   divisor UInt32
'       The value which divides dividend.
'   outRemainder
'       remainder of left divided-by right.
'
'@Returns
'   The quotient and remainder of left divided-by right.
''
Public Function DivRem(ByVal dividend As UInt32, ByVal divisor As UInt32, ByRef outRemainder As UInt32) As UInt32
    Dim remainder As ULong
    
    With New UInt32
        .Value = ULong32.DivRem(dividend.Value, divisor.Value, remainder)
        Set DivRem = .Self
    End With
    
    If outRemainder Is Nothing Then
        With New UInt32
            .Value = remainder
        Set outRemainder = .Self
        End With
    Else
        outRemainder.Value = remainder
    End If
End Function

''
'@Description "Returns a value indicating whether this instance is equal to a specified UInt32."
'@Parameters
'   obj UInt32
'       A value to compare to this instance.
'
'@Returns Boolean
'   true if obj has the same value as this instance; otherwise, false.
'
''
Public Function Equals(ByVal obj As UInt32) As Boolean
    If (this.Value = obj.Value.Value) Then
        Equals = True
    End If
End Function

''
'@Static
'@Description "Determines if a value represents an even integral number."
'@Parameters
'   Value UInt32
'       The value to be checked.
'
'@Returns Boolean
'   true if value is an even integer; otherwise, false.
''
Public Function IsEvenInteger(ByVal val As UInt32) As Boolean
    IsEvenInteger = ULong32.IsEvenInteger(val.Value)
End Function

''
'@Static
'@Description "Determines if a value represents an odd integral number."
'@Parameters
'   Value UInt32
'       The value to be checked.
'
'@Returns Boolean
'   true if value is an odd integer; otherwise, false.
''
Public Function IsOddInteger(ByVal val As UInt32) As Boolean
    IsOddInteger = ULong32.IsOddInteger(val.Value)
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
Public Function IsPow2(ByVal val As UInt32) As Boolean
    IsPow2 = ULong32.IsPow2(val.Value)
End Function

''
'@Static
'@Description "Compares two values to compute which is greater."
'@Parameters
'   lhs UInt32
'       The value to compare with rhs.
'   rhs UInt32
'       The value to compare with lhs.
'
'@Returns UInt32
'   lhs if it is greater than rhs; otherwise, rhs.
''
Public Function Max(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Max(lhs.Value, rhs.Value)
        Set Max = .Self
    End With
End Function

''
'@Static
'@Description "Compares two values to compute which is lesser."
'@Parameters
'   lhs UInt32
'       The value to compare with rhs.
'   rhs UInt32
'       The value to compare with lhs.
'
'@Returns UInt32
'   lhs if it is less than rhs; otherwise, rhs.
''
Public Function Min(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Min(lhs.Value, rhs.Value)
        Set Min = .Self
    End With
End Function

''
'@Static
'@Description "Converts the string representation of a number to its 32-bit unsigned integer equivalent."
'@Parameters
'   s String
'       A string representing the number to convert.
'
'@Returns UInt32
'   A 32-bit unsigned integer equivalent to the number contained in s.
'
'@Exceptions
'   ArgumentNullException
'       The s parameter is null.

'   FormatException
'       The s parameter is not of the correct format.

'   OverflowException
'       The s parameter represents a number that is less than UInt32.MinValue or greater than UInt32.MaxValue.
''
Public Function Parse(ByRef s As String) As UInt32
    With New UInt32
        .Value = ULong32.CreateChecked(s)
        Set Parse = .Self
    End With
End Function

''
'@Static
'@Description "Rotates a value left by a given amount."
'@Parameters
'   Value UInt32
'       The value which is rotated left by rotateAmount.
'   rotateAmount Int32
'       The amount by which value is rotated left.
'
'@Returns UInt32
'   The result of rotating value left by rotateAmount.
''
Public Function RotateLeft(ByVal val As UInt32, ByVal rotateAmount As Long) As UInt32
    With New UInt32
        .Value = ULong32.RotateLeft(val.Value, rotateAmount)
        Set RotateLeft = .Self
    End With
End Function

''
'@Static
'@Description "Rotates a value right by a given amount."
'@Parameters
'   Value UInt32
'       The value which is rotated right by rotateAmount.
'   rotateAmount Int32
'       The amount by which value is rotated right.
'
'@Returns UInt32
'   The result of rotating value right by rotateAmount.
''
Public Function RotateRight(ByVal val As UInt32, ByVal rotateAmount As Long) As UInt32
    With New UInt32
        .Value = ULong32.RotateRight(val.Value, rotateAmount)
        Set RotateRight = .Self
    End With
End Function

''
'@Description "Converts the numeric value of this instance to its equivalent string representation."
'@Returns String
'   The string representation of the value of this instance, consisting of a sequence of digits
'   ranging from 0 to 9, without a sign or leading zeroes.
''
Public Function ToString() As String
    ToString = ULong32.ToString(this)
End Function

'------------------------------------------------------------------
'Operators
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
Public Sub Assign(ByRef destination As UInt32, ByRef source As UInt32)
    If Not source Is Nothing Then
        If destination Is Nothing Then
            Set destination = Me.CreateFromULong(source.Value)
        Else
            destination.Value = source.Value
        End If
    End If
End Sub

'------------------------------------------------------------------
'Equality Operators
'------------------------------------------------------------------
''
'@Static
'@Description "Returns a boolean value indicating whether two unsigned 32-bit values are not equal."
'@Parameters
'   lhs: UInt32
'       unsigned  32-bit value to compare rhs.
'   rhs: UInt32
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if values are not equal.
''
Public Function NotEquals(ByVal lhs As UInt32, ByRef rhs As UInt32) As Boolean
    NotEquals = ULong32.NotEquals(lhs.Value, rhs.Value)
End Function

'------------------------------------------------------------------
'ComparisonOperators
'------------------------------------------------------------------

''
'@Static
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is less than another."
'@Parameters
'   lhs: UInt32
'       UInt32  32-bit value to compare rhs.
'   rhs: ULong
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is less than rhs value.
''
Public Function LessThan(ByVal lhs As UInt32, ByVal rhs As UInt32) As Boolean
    LessThan = ULong32.LessThan(lhs.Value, rhs.Value)
End Function

''
'@Static
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is less than or equal to another."
'@Parameters
'   lhs: UInt32
'       unsigned  32-bit value to compare rhs.
'   rhs: UInt32
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is less than or equal to rhs value.
''
Public Function LessThanOrEqual(ByVal lhs As UInt32, ByVal rhs As UInt32) As Boolean
    LessThanOrEqual = ULong32.LessThanOrEqual(lhs.Value, rhs.Value)
End Function

''
'@Static
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is greater than another."
'@Parameters
'   lhs: UInt32
'       unsigned  32-bit value to compare rhs.
'   rhs: UInt32
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is greater than rhs value.
''
Public Function GreaterThan(ByVal lhs As UInt32, ByVal rhs As UInt32) As Boolean
    GreaterThan = ULong32.GreaterThan(lhs.Value, rhs.Value)
End Function

''
'@Static
'@Description "Returns a boolean value indicating whether an unsigned 32-bit value is greater than or equal to another."
'@Parameters
'   lhs: UInt32
'       unsigned  32-bit value to compare rhs.
'   rhs: UInt32
'       unsigned  32-bit value  to compare to lhs.
'
'@Returns Boolean
'   Returns true if lhs is greater than or equal to rhs value.
''
Public Function GreaterThanOrEqual(ByVal lhs As UInt32, ByVal rhs As UInt32) As Boolean
    GreaterThanOrEqual = ULong32.GreaterThanOrEqual(lhs.Value, rhs.Value)
End Function

'------------------------------------------------------------------
'Decrement Operator
'------------------------------------------------------------------

''
'@Static
'@Description "Decrement an unsigned 32-bit integer."
'@Parameters
'   val: UInt32
'       unsigned  32-bit value to decrement.
''
Public Sub Decrement(ByRef val As UInt32)
    If Me.GreaterThan(val, Me.MinValue) Then
       val.Value = ULong32.Subtract(val.Value, ULong32.One)
    Else
        Err.Raise OverflowException, "UInt32.Decrement"
    End If
End Sub

'------------------------------------------------------------------
'Increment Operator
'------------------------------------------------------------------

''
'@Static
'@Description "Increment an unsigned 32-bit integer."
'@Parameters
'   val: UInt32
'       unsigned  32-bit value to increment.
''
Public Sub Increment(ByRef val As UInt32)
    If Me.LessThan(val, Me.MaxValue) Then
       val.Value = ULong32.Add(val.Value, ULong32.One)
    Else
        Err.Raise OverflowException, "UInt32.Increment"
    End If
End Sub

'------------------------------------------------------------------
'Math Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Divides two unsigned 32-bit values."
'@Parameters
'   dividend UInt32
'       The dividend to be divided by the divisor
'   divisor UInt32
'       The divisor to divide the dividend
'
'@Returns UInt32
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
''
Public Function Divide(ByVal dividend As UInt32, ByVal divisor As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Divide(dividend.Value, divisor.Value)
        Set Divide = .Self
    End With
End Function

''
'@Static
'@Description "Adds two unsigned 32-bit integers."
'@Parameters
'   lhs UInt32
'       value to be added to rhs
'   rhs UInt32
'       value to be addded  to lhs
'
'@Returns ULong
'   The result of the addition of two unsigned 32 bit integers.
'
'@Exceptions
'   OverflowException
'       Raised when a result exceeds the max UInt32 value of 4294967295
''
Public Function Add(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Add(lhs.Value, rhs.Value)
        Set Add = .Self
    End With
End Function

''
'@Static
'@Description "Multiplies two unsigned 32-bit integers."
'@Parameters
'   lhs UInt32
'       value to be multiply to rhs
'   rhs UInt32
'       value to be multiply to lhs
'
'@Returns UInt32
'   The result of the multiplying of two unsigned 32 bit values
'
'@Exceptions
'   OverflowException
'       Raised when the result exceeds the max unsigned 32-bit value of 4294967295
''
Public Function Multiply(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Multiply(lhs.Value, rhs.Value)
        Set Multiply = .Self
    End With
End Function

''
'@Static
'@Description "Subtracts two unsigned 32 bit values."
'@Parameters
'   lhs UInt32
'       Value to subtract from
'   rhs UInt32
'       Value to subtract from lhs
'@Returns UInt32
'
'@Exceptions
'   OverflowException
'       Raised when the result is less the minimum UInt32 value of 0.
'
'@Remarks
'   For Win64 subtract is performed by casting binary ULong values to QWORD containing LongLong
'   values without VBA numeric conversion.
''
Public Function Subtract(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Subtract(lhs.Value, rhs.Value)
        Set Subtract = .Self
    End With
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
Public Function Modulus(ByVal dividend As UInt32, ByRef divisor As UInt32) As UInt32
    With New UInt32
        .Value = ULong32.Modulus(dividend.Value, divisor.Value)
        Set Modulus = .Self
    End With
End Function

'------------------------------------------------------------------
'Bitwise Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Bitwise AND of two unsigned 32-bit value."
'@Parameters
'   lhs UInt32
'   rhs UInt32
'
'@Returns UInt32
''
Public Function BitwiseAnd(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value.Value = lhs.Value.Value And rhs.Value.Value
        Set BitwiseAnd = .Self
    End With
End Function

''
'@Static
'@Description "Bitwise NOT for an unsigned 32-bit value."
'@Parameters
'   val UInt32
'
'@Returns UInt32
''
Public Function BitwiseNot(ByVal val As UInt32) As UInt32
    With New UInt32
        .Value.Value = Not val.Value.Value
        Set BitwiseNot = .Self
    End With
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
Public Function BitwiseOr(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value.Value = lhs.Value.Value Or rhs.Value.Value
        Set BitwiseOr = .Self
    End With
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
Public Function BitwiseXor(ByVal lhs As UInt32, ByVal rhs As UInt32) As UInt32
    With New UInt32
        .Value.Value = lhs.Value.Value Xor rhs.Value.Value
        Set BitwiseXor = .Self
    End With
End Function

Public Function LeadingZeroCount(ByVal val As UInt32) As UInt32
    Err.Raise NotImplementedException, "UInt32.LeadingZeroCount"
End Function

Public Function PopCount(ByVal val As UInt32) As UInt32
    Err.Raise NotImplementedException, "UInt32.PopCount"
End Function

'------------------------------------------------------------------
'Shift Operators
'------------------------------------------------------------------

''
'@Static
'@Description "Bitwise logical left-shift operation, by the given number of bits."
'@Parameters
'   val UInt32
'       value
'   shiftAmount
'       The number of bits to left-shift by.
'
'@Returns UInt32
'   Returns a Int32 of the bitwise logical left-shift operation, by the given number of bits.
'
'@Remarks
'   Left shifting is equal to multiplying Value by 2Shift.
'   If offset <= 0 the value is returned
'   If the offset >= the maximum bitshift i.e. 32 then 0 is returned
''
Public Function ShiftLeft(ByVal val As UInt32, ByVal shiftAmount As Long) As UInt32
    With New UInt32
        .Value = ULong32.ShiftLeft(val.Value, shiftAmount)
        Set ShiftLeft = .Self
    End With
End Function

''
'@Static
'@Description "Bitwise logical right-shift operation, by the given number of bits."
'@Parameters
'   val UInt32
'       value
'   shiftAmount
'       The number of bits to right-shift by.
'
'@Returns UInt32
'   Returns a Int32 of the bitwise logical right-shift operation, by the given number of bits.
'
'@Remarks
'   Right shifting is equal to dividing value by 2^Shift.
''
Public Function ShiftRight(ByVal val As UInt32, ByVal shiftAmount As Long) As UInt32
    With New UInt32
        .Value = ULong32.ShiftRight(val.Value, shiftAmount)
        Set ShiftRight = .Self
    End With
End Function

'------------------------------------------------------------------
'IConvertible implementation
'------------------------------------------------------------------
Public Function ToBoolean() As Boolean
    Err.Raise NotImplementedException, "UInt32.ToBoolean"
End Function

Public Function ToByte() As Byte
    Err.Raise NotImplementedException, "UInt32.ToByte"
End Function

'Public Function ToChar() As Char
'End Function

''
'@Static
'@Description "Converts the instance into the VBA decimal type."
'
'@Returns Variant
'   The decimal value created fom the current instance.
''
Public Function ToDecimal() As Variant
    ToDecimal = ULong32.ToDecimal(this)
End Function

Public Function ToInt16() As Integer
    Err.Raise NotImplementedException, "UInt32.ToInt16"
End Function

Public Function ToInt32() As Long
    Err.Raise NotImplementedException, "ToInt32.LeadingZeroCount"
End Function

'Public Function ToInt64() As Int64
'End Function

''
'@Static
'@Description "Converts the instance into the LongLong type."
'
'@Returns LongLong
'   The LongLong value created fom the current instance.
''
#If Win64 Then
Public Function ToLongLong() As LongLong
    ToLongLong = ULong32.ToLongLong(this)
End Function
#End If

'Public Function ToUInt16() As UInt16
'End Function
