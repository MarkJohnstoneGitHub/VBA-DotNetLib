Attribute VB_Name = "DateTimeToString3Example"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 14, 2023
'@LastModified August 14, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tostring?view=netframework-4.8.1#system-datetime-tostring(system-iformatprovider)

Option Explicit

'@Description("The following example displays the string representation of a date and time using CultureInfo objects that represent five different cultures.")
Public Sub DateTimeToString3()
Attribute DateTimeToString3.VB_Description = "The following example displays the string representation of a date and time using CultureInfo objects that represent five different cultures."
    ' Create an array of four cultures.
    Dim cultures() As DotNetLib.CultureInfo
    Objects.ToArray cultures, _
                CultureInfo.InvariantCulture, _
                CultureInfo.GetCultureInfo2("en-us"), _
                CultureInfo.GetCultureInfo2("fr-fr"), _
                CultureInfo.GetCultureInfo2("de-DE"), _
                CultureInfo.GetCultureInfo2("es-ES"), _
                CultureInfo.GetCultureInfo2("ja-JP")
    
    Dim thisDate As DotNetLib.DateTime
    Set thisDate = DateTime.CreateFromDateTime(2009, 5, 1, 9, 0, 0)
    
    Dim varCulture As Variant
    For Each varCulture In cultures
        Dim culture As DotNetLib.CultureInfo
        Set culture = varCulture
        Dim cultureName As String
        
        If culture.Name = vbNullString Then
            cultureName = culture.NativeName
        Else
            cultureName = culture.Name
        End If
        Debug.Print "In "; cultureName; ", "; thisDate.ToString3(culture)
    Next
                  
End Sub

' The example produces the following output:
'    In Invariant Language (Invariant Country), 05/01/2009 09:00:00
'    In en-US, 5/1/2009 9:00:00 AM
'    In fr-FR, 01/05/2009 09:00:00
'    In de-DE, 01.05.2009 09:00:00
'    In es-ES, 01/05/2009 9:00:00
'    In ja-JP, 2009/05/01 9:00:00
