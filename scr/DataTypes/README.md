# Unsigned 32-bit data type 

Compatible with VBA Win64, Win32, Mac, VB6

Data Types

ULong32.cls Static class for processing unsigned 32 bit integers. This class processes the UDT ULong.  

  **Dependencies:**
  
    - [ULongType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/ULongType.bas)     


ULongType.bas

-ULong data structure for unsigned 32-bit integers	

UInt32.cls Class for 32-bit integer objects.

  **Dependencies:**
  
    - [ULongType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/ULongType.bas)  
    
    - [ULong32.cls](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/ULong32.cls)
    
    - [UInt32Static.cls](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/UInt32Static.cls) 
    

For Win32 the VBA Currency data type is used for the Multiply, Divide and DivRem math operations and results converted to an unsigned 32-bit integer.  Win64 uses the LongLong type for math operations and results converted to an unsigned 32-bit integer.

**Version 1.4 March 21, 2023 ULong32.cls**

- Performance improvements to math functions, Add, Subtract, Multiply, Divide and DivRem.  

- Updated Parsing function to support Hex and Octal string literals.  

- Updated CreateSaturating function to truncate decimals as per .Net 7.0 behaviour. 

- Updated CreateTruncating to support Decimal, Double, Single data types. Note behaviour for these types is same as CreateSaturating as per .Net 7.0 behaviour.


@TODO

Implement UInt32 class

Implement functions

- PopCount

- TrailingZeroCount

- TryFormat

- TryParse

- Parse Add overloads to parse formatting for currency sign, commas etc.

Testing for Bitwise functions

Unit Testing

Examples of use


Also working on UInt32 and UInt64 classes for unsigned integer objects which will be posted soon.
