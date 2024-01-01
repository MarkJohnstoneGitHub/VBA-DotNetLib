Attribute VB_Name = "StringIsNullOrWhiteSpaceExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 2, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.isnullorwhitespace?view=netframework-4.8.1#examples

Option Explicit

''
' The following example creates a string array, and then passes each element of
' the array to the IsNullOrWhiteSpace method.
''
Public Sub StringIsNullOrWhiteSpaceExample()
    Dim values() As DotNetLib.String
    Call ArrayEx.CreateInitialize1D(values, Nothing, Strings.EmptyString, Strings.Create("ABCDE"), _
                                    Strings.Create2(" ", 20), Strings.CreateUnescape("  \t   "), _
                                     Strings.Create2(VBAString.Unescape("\u2000"), 10))
    Dim varValue As Variant
    For Each varValue In values
        Debug.Print Strings.IsNullOrWhiteSpace(varValue)
    Next
End Sub

' The example displays the following output:
'       True
'       True
'       False
'       True
'       True
'       True
