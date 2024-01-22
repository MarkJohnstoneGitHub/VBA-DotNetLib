Attribute VB_Name = "StringEndsWithExample4"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.endswith?view=netframework-4.8.1#system-string-endswith(system-string-system-boolean-system-globalization-cultureinfo)

'@Todo Further testing required expected output not matching
' Example and COM Interop for String.EndsWith2 appear to be implemented correctly.
' https://stackoverflow.com/questions/68467553/c-sharp-endswith-sometimes-gives-false-results

Option Explicit

''
' This code example demonstrates the System.String.EndsWith(String, ..., CultureInfo) method.
' The following example determines whether a string occurs at the end of another string.
' The EndsWith method is called several times using case sensitivity, case insensitivity,
' and different cultures that influence the results of the search.
''
Public Sub StringEndsWithExample4()
    Dim msg1 As DotNetLib.String
    Set msg1 = Strings.CreateUnescape("Search for the target string ""{0}"" in the string ""{1}"".\n")
    Dim msg2 As DotNetLib.String
    Set msg2 = Strings.Create("Using the {0} - ""{1}"" culture:")
    Dim msg3 As DotNetLib.String
    Set msg3 = Strings.Create("  The string to search ends with the target string: {0}")
    Dim result As Boolean
    result = False
    Dim ci As DotNetLib.CultureInfo
    
    ' Define a target string to search for.
    ' U+00c5 = LATIN CAPITAL LETTER A WITH RING ABOVE
    Dim capitalARing As DotNetLib.String
    Set capitalARing = Strings.CreateUnescape("\u00c5")
    
    ' Define a string to search.
    ' The result of combining the characters LATIN SMALL LETTER A and COMBINING
    ' RING ABOVE (U+0061, U+030a) is linguistically equivalent to the character
    ' LATIN SMALL LETTER A WITH RING ABOVE (U+00e5).
    Dim xyzARing As DotNetLib.String
    Set xyzARing = Strings.CreateUnescape("xyz" + "\u0061\u030a")
    
    ' Display the string to search for and the string to search.
    Debug.Print VBString.Format(msg1.ToString, capitalARing, xyzARing)
       
    ' Search using English-United States culture.
    Set ci = CultureInfo.CreateFromName("en-US")
    Debug.Print VBString.Format(msg2.ToString(), ci.DisplayName, ci.name)

    Debug.Print "Case sensitive:"
    result = xyzARing.EndsWith2(capitalARing, False, ci)
    Debug.Print VBString.Format(msg3.ToString(), result)
    
    Debug.Print "Case insensitive:"
    result = xyzARing.EndsWith2(capitalARing, True, ci)
    Debug.Print VBString.Format(msg3.ToString(), result)
    Debug.Print

    ' Search using Swedish-Sweden culture.
    Set ci = CultureInfo.CreateFromName("sv-SE")
    Debug.Print VBString.Format(msg2.ToString(), ci.DisplayName, ci.name)

    Debug.Print "Case sensitive:"
    result = xyzARing.EndsWith2(capitalARing, False, ci)
    Debug.Print VBString.Format(msg3.ToString(), result)
    
    Debug.Print "Case insensitive:"
    result = xyzARing.EndsWith2(capitalARing, True, ci)
    Debug.Print VBString.Format(msg3.ToString, result)
End Sub

'/*
'This code example produces the following results (for en-us culture):
'
'Search for the target string "Å" in the string "xyza°".
'
'Using the English (United States) - "en-US" culture:
'Case sensitive:
'  The string to search ends with the target string: False
'Case insensitive:
'  The string to search ends with the target string: True
'
'Using the Swedish (Sweden) - "sv-SE" culture:
'Case sensitive:
'  The string to search ends with the target string: False
'Case insensitive:
'  The string to search ends with the target string: False
'
'*/

