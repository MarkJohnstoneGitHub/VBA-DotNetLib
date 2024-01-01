Attribute VB_Name = "StringLastIndexOfExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 2, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-string-system-int32-system-int32-system-stringcomparison)

Option Explicit

''
' This code example demonstrates the
' System.String.LastIndexOf(String, ..., StringComparison) methods.
' The following example demonstrates three overloads of the LastIndexOf method
' that find the last occurrence of a string within another string using
' different values of the StringComparison enumeration.
''
Public Sub StringLastIndexOfExample()
    Dim intro As String
    intro = "Find the last occurrence of a character using different " + _
                   "values of StringComparison."
    Dim resultFmt As String
    resultFmt = "Comparison: {0,-28} Location: {1,3}"

    '// Define a string to search for.
    '// U+00c5 = LATIN CAPITAL LETTER A WITH RING ABOVE
    Dim CapitalAWithRing As DotNetLib.String
    Set CapitalAWithRing = Strings.CreateUnescape("\u00c5")

    ' Define a string to search.
    ' The result of combining the characters LATIN SMALL LETTER A and COMBINING
    ' RING ABOVE (U+0061, U+030a) is linguistically equivalent to the character
    ' LATIN SMALL LETTER A WITH RING ABOVE (U+00e5).
    Dim cat As DotNetLib.String
    Set cat = Strings.CreateUnescape("A Cheshire c" + "\u0061\u030a" + "t")
    
    Dim loc As Long
    loc = 0
    Dim scValues() As mscorlib.StringComparison
    Call ArrayEx.CreateInitialize1D(scValues, _
                StringComparison.StringComparison_CurrentCulture, _
                StringComparison.StringComparison_CurrentCultureIgnoreCase, _
                StringComparison.StringComparison_InvariantCulture, _
                StringComparison.StringComparison_InvariantCultureIgnoreCase, _
                StringComparison.StringComparison_Ordinal, _
                StringComparison.StringComparison_OrdinalIgnoreCase)

    ' Clear the screen and display an introduction.
    Debug.Print intro
    
    ' Display the current culture because culture affects the result. For example,
    ' try this code example with the "sv-SE" (Swedish-Sweden) culture.
    
    Set CultureInfo.CurrentCulture = CultureInfo.CreateFromName("en-US")
    Debug.Print VBAString.Format("The current culture is ""{0}"" - {1}.", _
                        CultureInfo.CurrentCulture.Name, _
                        CultureInfo.CurrentCulture.DisplayName)

    ' Display the string to search for and the string to search.
    Debug.Print VBAString.Format("Search for the string ""{0}"" in the string ""{1}""", _
                CapitalAWithRing, cat)
    Debug.Print

    ' Note that in each of the following searches, we look for
    ' LATIN CAPITAL LETTER A WITH RING ABOVE in a string that contains
    ' LATIN SMALL LETTER A WITH RING ABOVE. A result value of -1 indicates
    ' the string was not found.
    ' Search using different values of StringComparison. Specify the start
    ' index and count.
    
    Debug.Print "Part 1: Start index and count are specified."
    Dim sc As Variant
    For Each sc In scValues
        loc = cat.LastIndexOf6(CapitalAWithRing, cat.length - 1, cat.length, sc)
        Debug.Print VBAString.Format(resultFmt, StringComparisonHelper.ToString(sc), loc)
    Next
    
    ' Search using different values of StringComparison. Specify the
    ' start index.
    Debug.Print Regex.Unescape("\nPart 2: Start index is specified.")
    For Each sc In scValues
        loc = cat.LastIndexOf5(CapitalAWithRing, cat.length - 1, sc)
        Debug.Print VBAString.Format(resultFmt, StringComparisonHelper.ToString(sc), loc)
    Next
    
    ' Search using different values of StringComparison.
    Debug.Print Regex.Unescape("\nPart 3: Neither start index nor count is specified.")
    For Each sc In scValues
        loc = cat.LastIndexOf4(CapitalAWithRing, sc)
        Debug.Print VBAString.Format(resultFmt, StringComparisonHelper.ToString(sc), loc)
    Next
End Sub

'/*
'Note: This code example was executed on a console whose user interface
'culture is "en-US" (English-United States).
'
'This code example produces the following results:
'
'Find the last occurrence of a character using different values of StringComparison.
'The current culture is "en-US" - English (United States).
'Search for the string "Å" in the string "A Cheshire ca°t"
'
'Part 1: Start index and count are specified.
'Comparison: CurrentCulture               Location:  -1
'Comparison: CurrentCultureIgnoreCase     Location:  12
'Comparison: InvariantCulture             Location:  -1
'Comparison: InvariantCultureIgnoreCase   Location:  12
'Comparison: Ordinal                      Location:  -1
'Comparison: OrdinalIgnoreCase            Location:  -1
'
'Part 2: Start index is specified.
'Comparison: CurrentCulture               Location:  -1
'Comparison: CurrentCultureIgnoreCase     Location:  12
'Comparison: InvariantCulture             Location:  -1
'Comparison: InvariantCultureIgnoreCase   Location:  12
'Comparison: Ordinal                      Location:  -1
'Comparison: OrdinalIgnoreCase            Location:  -1
'
'Part 3: Neither start index nor count is specified.
'Comparison: CurrentCulture               Location:  -1
'Comparison: CurrentCultureIgnoreCase     Location:  12
'Comparison: InvariantCulture             Location:  -1
'Comparison: InvariantCultureIgnoreCase   Location:  12
'Comparison: Ordinal                      Location:  -1
'Comparison: OrdinalIgnoreCase            Location:  -1
'
'*/
