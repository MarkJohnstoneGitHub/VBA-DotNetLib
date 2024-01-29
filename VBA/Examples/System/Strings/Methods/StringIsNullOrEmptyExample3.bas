Attribute VB_Name = "StringIsNullOrEmptyExample3"
'@IgnoreModule VariableNotAssigned
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 29, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.isnullorempty?view=netframework-4.8.1#what-is-a-null-string

Option Explicit

''
' A string is null if it has not been assigned a value or if it has explicitly
' been assigned a value of null. Although the composite formatting feature can
' gracefully handle a null string, as the following example shows, attempting to
' call one if its members throws a Object variable or With block variable not set.
''
Public Sub StringIsNullOrEmptyExample3()
    Dim s As DotNetLib.String
    Debug.Print VBString.Format("The value of the string is '{0}'", s)
    On Error Resume Next
    Debug.Print VBString.Format("String length is {0}", s.Length)
    If Err.Number Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

' The example displays the following output:
'    The value of the string is ''
'    Object variable or With block variable not set

