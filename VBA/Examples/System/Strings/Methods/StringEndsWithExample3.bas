Attribute VB_Name = "StringEndsWithExample3"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified December 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.endswith?view=netframework-4.8.1

Option Explicit

''
' This example demonstrates the
' System.String.EndsWith(String, StringComparison) method.
''
Public Sub StringEndsWithExample3()
    Dim intro As DotNetLib.String
    Set intro = Strings.CreateUnescape("Determine whether a string ends with another string, " + _
                   "using\n  different values of StringComparison.")
    Dim scValues() As mscorlib.StringComparison
    
     Call VBArray.CreateInitialize1D(scValues, _
            StringComparison.StringComparison_CurrentCulture, _
            StringComparison.StringComparison_CurrentCultureIgnoreCase, _
            StringComparison.StringComparison_InvariantCulture, _
            StringComparison.StringComparison_InvariantCultureIgnoreCase, _
            StringComparison.StringComparison_Ordinal, _
            StringComparison.StringComparison_OrdinalIgnoreCase)
            
    Debug.Print intro.ToString()
    
    ' Display the current culture because the culture-specific comparisons
    ' can produce different results with different cultures.
    Debug.Print VBString.Format(Regex.Unescape("The current culture is {0}.\n"), _
                           CultureInfo.CurrentCulture.Name)
                           
    ' Determine whether three versions of the letter I are equal to each other.
    Dim sc As Variant
    For Each sc In scValues
        Debug.Print VBString.Format("StringComparison.{0}:", StringComparisonHelper.ToString(sc))
        Call test(Strings.Create("abcXYZ"), Strings.Create("XYZ"), sc)
        Call test(Strings.Create("abcXYZ"), Strings.Create("xyz"), sc)
        Debug.Print
    Next
End Sub

Private Sub test(ByVal x As DotNetLib.String, ByVal y As DotNetLib.String, ByVal Comparison As mscorlib.StringComparison)
    Dim resultFmt As String
    resultFmt = """{0}"" {1} with ""{2}""."
    Dim result As String
    result = "does not end"
    If (x.EndsWith(y, Comparison)) Then
        result = "ends"
    End If
    Debug.Print VBString.Format(resultFmt, x, result, y)
End Sub

'/*
'This code example produces the following results:
'
'Determine whether a string ends with another string, using
'  different values of StringComparison.
'The current culture is en-US.
'
'StringComparison.CurrentCulture:
'"abcXYZ" ends with "XYZ".
'"abcXYZ" does not end with "xyz".
'
'StringComparison.CurrentCultureIgnoreCase:
'"abcXYZ" ends with "XYZ".
'"abcXYZ" ends with "xyz".
'
'StringComparison.InvariantCulture:
'"abcXYZ" ends with "XYZ".
'"abcXYZ" does not end with "xyz".
'
'StringComparison.InvariantCultureIgnoreCase:
'"abcXYZ" ends with "XYZ".
'"abcXYZ" ends with "xyz".
'
'StringComparison.Ordinal:
'"abcXYZ" ends with "XYZ".
'"abcXYZ" does not end with "xyz".
'
'StringComparison.OrdinalIgnoreCase:
'"abcXYZ" ends with "XYZ".
'"abcXYZ" ends with "xyz".
'
'*/


