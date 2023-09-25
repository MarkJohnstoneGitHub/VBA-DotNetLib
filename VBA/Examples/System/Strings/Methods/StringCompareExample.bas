Attribute VB_Name = "StringCompareExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified September 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.compare?view=net-7.0#system-string-compare(system-string-system-string)

Option Explicit

''
' The following example calls the Compare(String, String) method to compare
' three sets of strings.
''
Public Sub StringCompare()
    ' Create upper-case characters from their Unicode code units.
    Dim stringUpper As String
    stringUpper = "\x41\x42\x43"
    stringUpper = Regex.Unescape(stringUpper)

    ' Create lower-case characters from their Unicode code units.
    Dim stringLower As String
    stringLower = "\x61\x62\x63"
    stringLower = Regex.Unescape(stringLower)
    
    ' Display the strings.
    Dim output As String
    output = Strings.Format("Comparing '{0}' and '{1}':", stringUpper, stringLower)
    Debug.Print output
    
    ' Compare the uppercased strings; the result is true.
    Debug.Print Strings.Format("The Strings are equal when capitalized? {0}", _
            IIf(Strings.Compare(UCase(stringUpper), UCase(stringLower)) = 0, "true", "false"))
    
    ' The previous method call is equivalent to this Compare method, which ignores case.
    Debug.Print Strings.Format("The Strings are equal when case is ignored? {0}", _
            IIf(Strings.Compare(stringUpper, stringLower, True) = 0, "true", "false"))
End Sub

' The example displays the following output:
'       Comparing 'ABC' and 'abc':
'       The Strings are equal when capitalized? true
'       The Strings are equal when case is ignored? true
