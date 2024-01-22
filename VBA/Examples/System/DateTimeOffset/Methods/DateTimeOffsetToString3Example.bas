Attribute VB_Name = "DateTimeOffsetToString3Example"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 26, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.tostring?view=netframework-4.8.1#system-datetimeoffset-tostring(system-iformatprovider)

Option Explicit

''
' The following example displays a DateTimeOffset object using CultureInfo objects
' that represent the invariant culture, as well as four other cultures.
''
Public Sub DateTimeOffsetToString3()
    Dim cultures() As DotNetLib.CultureInfo
    ObjectArray.CreateInitialize1D cultures, _
        CultureInfo.InvariantCulture, _
        CultureInfo.CreateFromName("en-us"), _
        CultureInfo.CreateFromName("fr-fr"), _
        CultureInfo.CreateFromName("de-DE"), _
        CultureInfo.CreateFromName("es-ES")

    Dim thisDate As DotNetLib.DateTimeOffset
    Set thisDate = DateTimeOffset.CreateFromDateTimeParts(2007, 5, 1, 9, 0, 0, TimeSpan.Zero)

    Dim varCulture As Variant
    For Each varCulture In cultures
        Dim culture As DotNetLib.CultureInfo
        Set culture = varCulture
        Dim cultureName As String
        If culture.name = vbNullString Then
            cultureName = culture.NativeName
        Else
            cultureName = culture.name
        End If
        Debug.Print VBString.Format("In {0}, {1}", _
                                    cultureName, thisDate.ToString3(culture))
    Next
End Sub

' The example produces the following output:
'    In Invariant Language (Invariant Country), 05/01/2007 09:00:00 +00:00
'    In en-US, 5/1/2007 9:00:00 AM +00:00
'    In fr-FR, 01/05/2007 09:00:00 +00:00
'    In de-DE, 01.05.2007 09:00:00 +00:00
'    In es-ES, 01/05/2007 9:00:00 +00:00


