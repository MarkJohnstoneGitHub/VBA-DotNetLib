Attribute VB_Name = "StringBuilderAppendFormatEg2"
'@Folder("Examples.System.Text.StringBuilder.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 28, 2023
'@LastModified January 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.appendformat?view=netframework-4.8.1#system-text-stringbuilder-appendformat(system-iformatprovider-system-string-system-object)

Option Explicit

''
' The following includes two calls to the
' AppendFormat(IFormatProvider, String, Object) method. Both use the formatting
' conventions of the English-United Kingdom (en-GB) culture. The first inserts
' the string representation of a Decimal value currency in a result string.
' The second inserts a DateTime value in two places in a result string, the
' first including only the short date string and the second the short time string.
''
Public Sub StringBuilderAppendFormatExample2()
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create()
    Dim pvtValue As Variant
    pvtValue = CDec(16.95)
    Dim enGB As DotNetLib.CultureInfo
    Set enGB = CultureInfo.CreateSpecificCulture("en-GB")
    Dim dateToday As DotNetLib.DateTime
    Set dateToday = DateTime.Now
    Call sb.AppendFormat5(enGB, "Final Price: {0:C2}", pvtValue)
    Call sb.AppendLine
    Call sb.AppendFormat5(enGB, "Date and Time: {0:d} at {0:t}", dateToday)
    Debug.Print sb.ToString()
End Sub

' The example displays the following output:
'       Final Price: £16.95
'       Date and Time: 01/10/2014 at 10:22
