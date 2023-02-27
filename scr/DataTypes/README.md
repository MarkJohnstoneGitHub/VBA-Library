# Unsigned 32 bit data type

Data Types
UInt32Static.cls Static class for unsigned 32 bit values

  **Dependencies:**
  
    - [ULongType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/ULongType.bas)
    
    - [DWordType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/DWordType.bas)
    
    - [QWordType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/QWordType.bas) 
    
    - [WordType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/WordType.bas)      
    
    - [VBADecimalType.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/DataTypes/VBADecimalType.bas)
    
    - [CopyMemoryAPI.bas](https://github.com/MarkJohnstoneGitHub/VBA-Library/blob/main/scr/API/CopyMemoryAPI.bas)    

ULongType.bas

-ULong data structure for unsigned 32 bit values	

VBADecimalType.bas

-VBA Decimal type structure within a variant.

CopyMemoryAPI.bas

-API declarations for copy memory by pointer for Windows and Mac, with VBA6 and VBA7 compatibility.


Currently predominantly using the VBA Decimal data type for math operations as internally it is treated unsigned value making for reliable method for converting to unsigned 32 bit value. Updated performance for multipy and add so far Win64 to use QWORD  and LongLong type. 

@TODO

Unit Testing

Examples of use

Investigate methods to improve performance for various math operations and functions

Possible solutions:

-Efficient bitwise methods

-COM addin

-Using other existing data types and bitshifting
