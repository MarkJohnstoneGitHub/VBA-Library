# Unsigned 32 bit value

VBA Utility Classes.
UInt32Static.cls Static class for unsigned 32 bit values

ULongType.bas ULong data structure for unsigned 32 bit values	

Currently predominantly using the Decimal data type for math operatioons as internally it is treated unsigned value making for reliable and compatible method for converting to unsigned 32 bit value.

@TODO
Investigate methods to improve performance for various math operations and functions
Possible solutions
-Efficient bitwise methods
-COM addin
-Using other existing data types and bitshifting
