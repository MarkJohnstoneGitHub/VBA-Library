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
    


For Win32 the VBA Decimal data type is used for the Multiply, Divide and DivRem math operations and results converted to an unsigned 32-bit integer.  Win64 uses the LongLong type for math operations and results converted to an unsigned 32-bit integer.

@TODO

Implement UInt32 class

Implement functions

PopCount

LeadingZeroCount

Testing for Bitwise functions

Unit Testing

Examples of use


Also working on UInt32 and UInt64 classes for unsigned integer objects which will be posted soon.
