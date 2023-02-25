# Unsigned 32 bit data type

Data Types
UInt32Static.cls Static class for unsigned 32 bit values

  Dependances:
  
    ULongType.bas
    
    VBADecimalType.bas
    
    CopyMemoryAPI.bas

ULongType.bas

-ULong data structure for unsigned 32 bit values	

VBADecimalType.bas

-VBA Decimal type structure within a variant.

CopyMemoryAPI.bas

-API declarations for copy memory by pointer for Windows and Mac, with VBA6 and VBA7 compatibility.


Currently predominantly using the VBA Decimal data type for math operations as internally it is treated unsigned value making for reliable method for converting to unsigned 32 bit value.

@TODO
Investigate methods to improve performance for various math operations and functions

Possible solutions:

-Efficient bitwise methods

-COM addin

-Using other existing data types and bitshifting
