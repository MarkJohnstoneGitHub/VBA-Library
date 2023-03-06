# Unsigned 32 bit data type

Data Types
UInt32Static.cls Static class for unsigned 32 bit values

  **Dependencies:**
  
    - [ULongType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/ULongType.bas)
    
    - [DWordType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/DWordType.bas)
    
    - [QWordType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/QWordType.bas) 
    
    - [WordType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/WordType.bas)      
   
      
ULongType.bas

-ULong data structure for unsigned 32 bit values	

For Win32 the VBA Decimal data type is used for the Multiply, Divide and DivRem math operations and results converted to an unsigned 32-bit integer.  Win64 uses the LongLong type for math operations and results converted to an unsigned 32-bit integer.

@TODO

Are currency, single, double data types appropriate to typecasted to ULong?  Require investigating how other languages handle these types and unsigned integers.

Unit Testing

Examples of use
