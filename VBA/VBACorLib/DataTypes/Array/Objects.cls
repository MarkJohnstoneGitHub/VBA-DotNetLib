VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Objects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "VBACorLib.DataTypes.Array"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified July 16, 2023

Option Explicit

'@Description("Copies a ParamArray of objects to an array.")
Public Sub ToArray(ByRef destination As Variant, ParamArray values() As Variant)
Attribute ToArray.VB_Description = "Copies a ParamArray of objects to an array."
   If VBA.IsArray(destination) Then
      ReDim destination(LBound(values) To UBound(values))
      Dim index As Long
      Dim varValue As Variant
      On Error GoTo Catch
      For Each varValue In values
         Set destination(index) = varValue
         index = index + 1
      Next
   Else
      Err.Raise VBAException.INVALID_PROCEDURE_CALL, , "Only arrays are supported."
   End If
Exit Sub
Catch:
      Err.Raise VBAException.INVALID_PROCEDURE_CALL, , "Invalid argument, array item is an invalid type for " & TypeName(destination)
End Sub
