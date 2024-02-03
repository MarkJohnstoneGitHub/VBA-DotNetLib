Attribute VB_Name = "RegexIsMatchExample2"
'@Folder("Examples.System.Text.RegularExpressions.Regex.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 3, 2024
'@LastModified February 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.ismatch?view=netframework-4.8.1#system-text-regularexpressions-regex-ismatch(system-string)

Option Explicit

''
' The following example illustrates the use of the IsMatch(String) method to
' determine whether a string is a valid part number. The regular expression
' assumes that the part number has a specific format that consists of three
' sets of characters separated by hyphens. The first set, which contains four
' characters, must consist of an alphanumeric character followed by two numeric
' characters followed by an alphanumeric character. The second set, which
' consists of three characters, must be numeric. The third set, which consists
' of four characters, must have three numeric characters followed by an
' alphanumeric character.
''
Public Sub RegexIsMatchExample2()
    Dim partNumbers() As String
    partNumbers = StringArray.CreateInitialize1D("1298-673-4192", "A08Z-931-468A", _
                            "_A90-123-129X", "12345-KKA-1230", _
                            "0919-2893-1256")
    Dim rgx As DotNetLib.Regex
    Set rgx = Regex.Create("^[a-zA-Z0-9]\d{2}[a-zA-Z0-9](-\d{3}){2}[A-Za-z0-9]$")
    
    Dim partNumber As Variant
    For Each partNumber In partNumbers
        Debug.Print VBString.Format("{0} {1} a valid part number.", _
                           partNumber, _
                           IIf(rgx.IsMatch(partNumber), "is", "is not"))
    Next
End Sub

' The example displays the following output:
'       1298-673-4192 is a valid part number.
'       A08Z-931-468A is a valid part number.
'       _A90-123-129X is not a valid part number.
'       12345-KKA-1230 is not a valid part number.
'       0919-2893-1256 is not a valid part number.
