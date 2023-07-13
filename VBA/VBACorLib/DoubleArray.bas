Attribute VB_Name = "DoubleArray"
'Rubberduck annotations
'@Folder "VBACorLib.DataTypes.Array"

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 13, 2023

Option Explicit

Public Function ArrayDouble(ParamArray values() As Variant) As Double()
    Dim result() As Double
    ReDim result(LBound(values) To UBound(values))
    Dim index As Long
    For index = LBound(values) To UBound(values)
        result(index) = values(index)
    Next index
    ArrayDouble = result
End Function

