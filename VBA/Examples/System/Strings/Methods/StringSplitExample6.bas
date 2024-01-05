Attribute VB_Name = "StringSplitExample6"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 6, 2024
'@LastModified January 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.split?view=netframework-4.8.1#system-string-split(system-string()-system-stringsplitoptions)

Option Explicit

''
' The following example illustrates the difference in the arrays returned by
' calling a string's String.Split(String[], StringSplitOptions) method with
' its options parameter equal to StringSplitOptions.None and
' StringSplitOptions.RemoveEmptyEntries.
''
Public Sub StringSplitExample6()
    Dim source As DotNetLib.String
    Set source = Strings.Create("[stop]ONE[stop][stop]TWO[stop][stop][stop]THREE[stop][stop]")
    Dim stringSeparators() As String
    Call ArrayEx.CreateInitialize1D(stringSeparators, "[stop]")
    Dim result() As String
    
    ' Display the original string and delimiter string.
    Debug.Print VBAString.Format(VBAString.Unescape("Splitting the string:\n   ""{0}""."), source)
    Debug.Print
    Debug.Print VBAString.Format(VBAString.Unescape("Using the delimiter string:\n   ""{0}"""), stringSeparators(0))
    Debug.Print

    ' Split a string delimited by another string and return all elements.
    result = source.Split3(stringSeparators, StringSplitOptions.StringSplitOptions_None)
    Debug.Print VBAString.Format("Result including all elements ({0} elements):", UBound(result) + 1)
    Debug.Print "   ";
    Dim s As Variant
    For Each s In result
        Debug.Print VBAString.Format("'{0}' ", IIf(VBAString.IsNullOrEmpty(s), "<>", s));
    Next
    Debug.Print
    Debug.Print

    ' Split delimited by another string and return all non-empty elements.
    result = source.Split3(stringSeparators, StringSplitOptions.StringSplitOptions_RemoveEmptyEntries)
    Debug.Print VBAString.Format("Result including non-empty elements ({0} elements):", UBound(result) + 1)
    Debug.Print "   ";
    For Each s In result
        Debug.Print VBAString.Format("'{0}' ", IIf(VBAString.IsNullOrEmpty(s), "<>", s));
    Next
    Debug.Print
End Sub

' The example displays the following output:
'    Splitting the string:
'       "[stop]ONE[stop][stop]TWO[stop][stop][stop]THREE[stop][stop]".
'
'    Using the delimiter string:
'       "[stop]"
'
'    Result including all elements (9 elements):
'       '<>' 'ONE' '<>' 'TWO' '<>' '<>' 'THREE' '<>' '<>'
'
'    Result including non-empty elements (3 elements):
'       'ONE' 'TWO' 'THREE'
