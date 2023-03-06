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

Currently for Win32 using the VBA Decimal data type for math operations as internally it is treated as an unsigned value making for reliable method for converting to unsigned 32 bit value.

@TODO

Are currency, single, double appropriate to typecasted to ULong?  Require investigating how other languages handle these types and unsigned integers.

Unit Testing

Examples of use
