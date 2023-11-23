VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Manages a compact array of bit values, which are represented as Booleans, where true indicates that the bit is on (1) and false indicates the bit is off (0)."
'@Folder "VBADotNetLib.System.Collections"
'@PredeclaredId
'@ModuleDescription("Manages a compact array of bit values, which are represented as Booleans, where true indicates that the bit is on (1) and false indicates the bit is off (0).")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified November 2, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray?view=netframework-4.8.1

'@Remarks
'Singleton class for BitArray containing factory methods and static members.

Option Explicit

Private Type TBitArraySingleton
   BitArraySingleton As DotNetLib.BitArraySingleton
End Type

Private this As TBitArraySingleton

Private Sub Class_Initialize()
   Set this.BitArraySingleton = New DotNetLib.BitArraySingleton
End Sub

Private Sub Class_Terminate()
   Set this.BitArraySingleton = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the BitArray class that can hold the specified number of bit values, which are initially set to false.")
Public Function Create(ByVal length As Long) As DotNetLib.BitArray
Attribute Create.VB_Description = "Initializes a new instance of the BitArray class that can hold the specified number of bit values, which are initially set to false."
   Set Create = this.BitArraySingleton.Create(length)
End Function

'@Static
'@Description("Initializes a new instance of the BitArray class that can hold the specified number of bit values, which are initially set to the specified value.")
Public Function Create2(ByVal length As Long, ByVal defaultValue As Boolean) As DotNetLib.BitArray
Attribute Create2.VB_Description = "Initializes a new instance of the BitArray class that can hold the specified number of bit values, which are initially set to the specified value."
   Set Create2 = this.BitArraySingleton.Create_2(length, defaultValue)
End Function

'@Static
'@Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified array of bytes.")
Public Function Create3(ByRef bytes() As Byte) As DotNetLib.BitArray
Attribute Create3.VB_Description = "Initializes a new instance of the BitArray class that contains bit values copied from the specified array of bytes."
   Set Create3 = this.BitArraySingleton.Create_3(bytes)
End Function

'@Static
'@Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified array of Booleans.")
Public Function Create4(ByRef values() As Boolean) As DotNetLib.BitArray
Attribute Create4.VB_Description = "Initializes a new instance of the BitArray class that contains bit values copied from the specified array of Booleans."
   Set Create4 = this.BitArraySingleton.Create_4(values)
End Function

'@Static
'@Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified array of 32-bit integers.")
Public Function Create5(ByRef values() As Long) As DotNetLib.BitArray
Attribute Create5.VB_Description = "Initializes a new instance of the BitArray class that contains bit values copied from the specified array of 32-bit integers."
   Set Create5 = this.BitArraySingleton.Create_5(values)
End Function

'@Static
'@Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified BitArray.")
Public Function Create6(ByVal bits As DotNetLib.BitArray) As DotNetLib.BitArray
Attribute Create6.VB_Description = "Initializes a new instance of the BitArray class that contains bit values copied from the specified BitArray."
   Set Create6 = this.BitArraySingleton.Create_6(bits)
End Function
