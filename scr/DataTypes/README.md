# Unsigned 32 bit value

Data Types
UInt32Static.cls Static class for unsigned 32 bit values

  Dependances:
  
    ULongType.bas
    
    CopyMemoryAPI.bas

ULongType.bas 
ULong data structure for unsigned 32 bit values	

Currently predominantly using the VBA Decimal data type for math operations as internally it is treated unsigned value making for reliable method for converting to unsigned 32 bit value.

@TODO
Investigate methods to improve performance for various math operations and functions

Possible solutions:

-Efficient bitwise methods

-COM addin

-Using other existing data types and bitshifting
