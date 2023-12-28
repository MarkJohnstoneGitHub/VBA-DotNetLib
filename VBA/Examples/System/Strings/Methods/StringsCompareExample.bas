Attribute VB_Name = "StringsCompareExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 23, 2023
'@LastModified September 26, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.compare?view=netframework-4.8.1#system-string-compare(system-string-system-string)

Option Explicit

''
' The following example calls the Compare(String, String) method to compare
' three sets of strings.
''
Public Sub StringCompare()
    ' Create upper-case characters from their Unicode code units.
    Dim stringUpper As String
    stringUpper = Regex.Unescape("\x41\x42\x43")

    ' Create lower-case characters from their Unicode code units.
    Dim stringLower As String
    stringLower = Regex.Unescape("\x61\x62\x63")
    
    ' Display the strings.
    Dim Output As String
    Output = VBAString.Format("Comparing '{0}' and '{1}':", stringUpper, stringLower)
    Debug.Print Output
    
    ' Compare the uppercased strings; the result is true.
    Debug.Print VBAString.Format("The Strings are equal when capitalized? {0}", _
            IIf(VBAString.Compare(UCase$(stringUpper), UCase$(stringLower)) = 0, "true", "false"))
    
    ' The previous method call is equivalent to this Compare method, which ignores case.
    Debug.Print VBAString.Format("The Strings are equal when case is ignored? {0}", _
            IIf(VBAString.Compare(stringUpper, stringLower, True) = 0, "true", "false"))
End Sub

' The example displays the following output:
'       Comparing 'ABC' and 'abc':
'       The Strings are equal when capitalized? true
'       The Strings are equal when case is ignored? true


