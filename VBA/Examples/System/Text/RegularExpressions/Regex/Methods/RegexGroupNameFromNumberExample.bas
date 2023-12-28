Attribute VB_Name = "RegexGroupNameFromNumberExample"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 3, 2023
'@LastModified October 3, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.groupnamefromnumber?view=netframework-4.8.1#examples

Option Explicit

''
' The following example defines a regular expression pattern that matches an
' address line containing a U.S. city name, state name, and zip code.
' The example uses the GroupNameFromNumber method to retrieve the names of
' capturing groups. It then uses these names to retrieve the corresponding
' captured groups for matches.
''
Public Sub RegexGroupNameFromNumber()
    Dim pattern As String
    pattern = "(?<city>[A-Za-z\s]+), (?<state>[A-Za-z]{2}) (?<zip>\d{5}(-\d{4})?)"
    Dim cityLines() As String
    cityLines = StringArray.ToArray("New York, NY 10003", "Brooklyn, NY 11238", "Detroit, MI 48204", _
                                    "San Francisco, CA 94109", "Seattle, WA 98109")
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create(pattern)
    Dim names As DotNetLib.ListString
    Set names = ListString.Create()
    
    Dim ctr As Long
    ctr = 1
    Dim exitFlag As Boolean
    exitFlag = False
    ' Get group names.
    Do
        Dim Name As String
        Name = rgx.GroupNameFromNumber(ctr)
        If (Not VBAString.IsNullOrEmpty(Name)) Then
            names.Add Name
            ctr = ctr + 1
        Else
            exitFlag = True
        End If
    Loop While (Not exitFlag)
    Dim cityLine As Variant
    For Each cityLine In cityLines
        Dim pvtMatch As DotNetLib.Match
        Set pvtMatch = rgx.Match(cityLine)
        If (pvtMatch.Success) Then
            Debug.Print VBAString.Format("Zip code {0} is in {1}, {2}.", _
                                        pvtMatch.Groups.Item_2(names(3)), _
                                        pvtMatch.Groups.Item_2(names(1)), _
                                        pvtMatch.Groups.Item_2(names(2)))
        End If
    Next
End Sub

' The example displays the following output:
'       Zip code 10003 is in New York, NY.
'       Zip code 11238 is in Brooklyn, NY.
'       Zip code 48204 is in Detroit, MI.
'       Zip code 94109 is in San Francisco, CA.
'       Zip code 98109 is in Seattle, WA.


