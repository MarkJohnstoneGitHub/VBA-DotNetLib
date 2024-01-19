Attribute VB_Name = "DateTimeNowExample"
'@Folder "Examples.System.DateTime.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 14, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.now?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses the Now and UtcNow properties to retrieve the
' current local date and time and the current universal coordinated (UTC)
' date and time. It then uses the formatting conventions of a number of
' cultures to display the strings, along with the values of their Kind
' properties.
''
Public Sub DateTimeNow()
    Dim localDate As DotNetLib.DateTime
    Set localDate = DateTime.Now
    Dim utcDate As DotNetLib.DateTime
    Set utcDate = DateTime.UtcNow
    
    Dim cultureNames() As String
    cultureNames = StringArray.CreateInitialize1D("en-US", "en-GB", "fr-FR", _
                                    "de-DE", "ru-RU")
    Dim cultureName As Variant
    For Each cultureName In cultureNames
        Dim culture As DotNetLib.CultureInfo
        Set culture = CultureInfo.CreateFromName(cultureName)
        Debug.Print VBString.Format("{0}:", culture.NativeName)
        Debug.Print VBString.Format("   Local date and time: {0}, {1:G}", _
                          localDate.ToString3(culture), DateTimeKindHelper.ToString(localDate.Kind))
        Debug.Print VBString.Format(VBString.Unescape("   UTC date and time: {0}, {1:G}\n"), _
                          utcDate.ToString3(culture), DateTimeKindHelper.ToString(utcDate.Kind))
    Next
End Sub

' The example displays the following output:
'       English (United States):
'          Local date and time: 6/19/2015 10:35:50 AM, Local
'          UTC date and time: 6/19/2015 5:35:50 PM, Utc
'
'       English (United Kingdom):
'          Local date and time: 19/06/2015 10:35:50, Local
'          UTC date and time: 19/06/2015 17:35:50, Utc
'
'       français (France):
'          Local date and time: 19/06/2015 10:35:50, Local
'          UTC date and time: 19/06/2015 17:35:50, Utc
'
'       Deutsch (Deutschland):
'          Local date and time: 19.06.2015 10:35:50, Local
'          UTC date and time: 19.06.2015 17:35:50, Utc
'
'       русский (Россия):
'          Local date and time: 19.06.2015 10:35:50, Local
'          UTC date and time: 19.06.2015 17:35:50, Utc


