VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBACorLib.DataTypes.Array")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 14, 2023
'@LastModified July 16, 2023

Option Explicit

''
'@Static
'@Description "Creates a one dimensional string array from a list of parameters"
'@Parameters
'   values() ParamArray
'       List of parameters to be converted into a string array
'@Returns String()
'   Returns a string array.
''
Public Function ToArray(ParamArray values() As Variant) As String()
Attribute ToArray.VB_Description = "Creates a one dimensional string array from a list of parameters"
   Dim result() As String
   ReDim result(LBound(values) To UBound(values))
   Dim index As Long
   
   For index = LBound(values) To UBound(values)
       result(index) = CStr(values(index))
   Next index
   ToArray = result
End Function

'@TODO
'''
''@Static
''@Description "Creates a dictionary of string key value pairs from a list of parameters"
''@Parameters
''   values() ParamArray
''       List of parameters to be converted into a dictionary of string key and string value pairs.
''@Returns String()
''   Returns Dictionary
''       Dictionary of string keys and string values.
'''
'Public Function ToDictionary(ParamArray keyValuePairs() As Variant) As Scripting.Dictionary
'    'Must be key,value pairs i.e. even amount of items
'    If (UBound(keyValuePairs) + 1) Mod 2 Then
'        Err.Raise 5
'    End If
'    Dim result As Scripting.Dictionary
'    Set result = New Scripting.Dictionary
'    Dim index As Long
'    For index = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
'        If result.Exists(keyValuePairs(index)) Then
'            Err.Raise 457
'        End If
'        result.Add key:=CStr(keyValuePairs(index)), Item:=CStr(keyValuePairs(index + 1))
'        Next
'    Set ToDictionary = result
'End Function
