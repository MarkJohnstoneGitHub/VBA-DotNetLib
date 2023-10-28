VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "VBACorLib.DataTypes.Array"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 16, 2023
'@LastModified October 29, 2023

'@TODO Work in progress

Option Explicit

'@Description("Creates a one-dimensional array of the specified Type with initial values")
Public Sub CreateInitialize1D(ByRef destination As Variant, ParamArray Values() As Variant)
Attribute CreateInitialize1D.VB_Description = "Creates a one-dimensional array of the specified Type with initial values"
   If VBA.IsArray(destination) Then
      '@TODO What if values is empty?
      ReDim destination(LBound(Values) To UBound(Values))
      Dim index As Long
      Dim varValue As Variant
      On Error GoTo Catch
      For Each varValue In Values
         If IsObject(varValue) Then
            Set destination(index) = varValue
         Else
            destination(index) = varValue
         End If
         index = index + 1
      Next
   Else
      Err.Raise VBAException.INVALID_PROCEDURE_CALL, , "Only arrays are supported."
   End If
Exit Sub
Catch:
      Err.Raise VBAException.INVALID_PROCEDURE_CALL, , "Invalid argument, array item is an invalid type for " & typeName(destination)
End Sub

'@Description("Copies a ParamArray of vaules or objects to an array of a type.")
'Return value of False indictates that errors occurred and not all items where transfered to destination array
Public Function TryToArray(ByRef destination As Variant, ParamArray Values() As Variant) As Boolean
Attribute TryToArray.VB_Description = "Copies a ParamArray of vaules or objects to an array of a type."
   On Error Resume Next
   If (TryToArray = VBA.IsArray(destination)) Then
      Dim Resize As Boolean
      Resize = False
      
      ReDim destination(LBound(Values) To UBound(Values))
      Dim index As Long
      Dim varValue As Variant
      For Each varValue In Values
         If IsObject(varValue) Then
            Set destination(index) = varValue
         Else
            destination(index) = varValue
         End If
         If Err.Number = 0 Then
            index = index + 1
         Else
            TryToArray = False
            Resize = True
         End If
      Next
      If Resize Then
         ReDim Preserve destination(index - 1)
      End If
   Else
      Err.Raise VBAException.INVALID_PROCEDURE_CALL, , "Only arrays are supported."
   End If
   On Error GoTo 0 'Stop code and display error
End Function

'''
'@Description("Returns if an array variable is initialized.")
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'@Reference http://www.cpearson.com/excel/vbaarrays.htm
'''
Public Function IsAllocated(ByRef Source As Variant) As Boolean
Attribute IsAllocated.VB_Description = "Returns if an array variable is initialized."
   Dim sourceUBound As Long
   On Error Resume Next
   
   ' if Arr is not an array, return FALSE and get out.
   If IsArray(Source) = False Then
       IsAllocated = False
       Exit Function
   End If
   
   ' Attempt to get the UBound of the array. If the array has not been allocated,
   ' an error will occur. Test Err.Number to see if an error occurred.
   sourceUBound = UBound(Source, 1)
   If (Err.Number = 0) Then
       ''''''''''''''''''''''''''''''''''''''
       ' Under some circumstances, if an array
       ' is not allocated, Err.Number will be
       ' 0. To acccomodate this case, we test
       ' whether LBound <= Ubound. If this
       ' is True, the array is allocated. Otherwise,
       ' the array is not allocated.
       '''''''''''''''''''''''''''''''''''''''
       IsAllocated = (LBound(Source) <= UBound(Source))
   Else
       ' error. unallocated array
       IsAllocated = False
   End If
   On Error GoTo 0 'Stop code and display error
End Function
