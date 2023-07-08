Attribute VB_Name = "DoubleArray"
'@Folder "VBACorLib.DataTypes.Array"

'https://github.com/MarkJohnstoneGitHub/
'@Version v1.0 February 10, 2023
'@LastModified February 10, 2023

Option Explicit

Public Function ArrayDouble(ParamArray values() As Variant) As Double()
    Dim result() As Double
    ReDim result(LBound(values) To UBound(values))
    Dim i As Long
    
    For i = LBound(values) To UBound(values)
        result(i) = values(i)
    Next i
    ArrayDouble = result
End Function



