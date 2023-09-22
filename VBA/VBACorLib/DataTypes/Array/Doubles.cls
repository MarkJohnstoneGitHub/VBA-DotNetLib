VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Doubles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "VBACorLib.DataTypes.Array"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified July 16, 2023

Option Explicit

''
'@Static
'@Description "Creates a one dimensional double array from a list of parameters"
'@Parameters
'   values() ParamArray
'       List of parameters to be converted into a string array
'@Returns Double()
'   Returns a double array.
''
Public Function ToArray(ParamArray values() As Variant) As Double()
Attribute ToArray.VB_Description = "Creates a one dimensional double array from a list of parameters"
   Dim result() As Double
   '@TODO What if values is empty?
   ReDim result(LBound(values) To UBound(values))
   Dim index As Long
   For index = LBound(values) To UBound(values)
       result(index) = values(index)
   Next index
   ToArray = result
End Function
