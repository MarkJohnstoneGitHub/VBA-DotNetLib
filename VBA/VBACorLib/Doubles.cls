VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Doubles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBACorLib.DataTypes.Array")
'@PredeclaredId

'Copyright(c) 2023 Mark Johnstone
'MarkJohnstoneGitHub/VBA-dotnet-CorLib is licensed under the MIT License
'@Version v1.0 July 14, 2023
'@LastModified July 14, 2023

Option Explicit

Public Function ToArray(ParamArray values() As Variant) As Double()
    Dim result() As Double
    ReDim result(LBound(values) To UBound(values))
    Dim index As Long
    For index = LBound(values) To UBound(values)
        result(index) = values(index)
    Next index
    ToArray = result
End Function
