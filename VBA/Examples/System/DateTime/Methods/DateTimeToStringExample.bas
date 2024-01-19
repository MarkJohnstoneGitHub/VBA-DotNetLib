Attribute VB_Name = "DateTimeToStringExample"
'@Folder "Examples.System.DateTime.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 13, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.tostring?view=netframework-4.8.1#system-datetime-tostring

Option Explicit

''
' The following example illustrates how the string representation of a DateTime
' value returned by the ToString() method depends on the thread current culture.
' It changes the current culture to en-US, fr-FR, and ja-JP, and in each case
' calls the ToString() method to return the string representation of a date and
' time value using that culture.
''
Public Sub DateTimeToString()
    Dim pvtCurrentCulture As DotNetLib.CultureInfo
    Set pvtCurrentCulture = CultureInfo.CurrentCulture
    Dim exampleDate As DotNetLib.DateTime
    Set exampleDate = DateTime.CreateFromDateTime(2021, 5, 1, 18, 32, 6)
        
    ' Change the current culture to en-US and display the date.
    Set CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo2("en-US")
    Debug.Print exampleDate.ToString()
    
    ' Change the current culture to fr-FR and display the date.
    Set CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo2("fr-FR")
    Debug.Print exampleDate.ToString()
    
    ' Change the current culture to ja-JP and display the date.
    Set CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo2("ja-JP")
    Debug.Print exampleDate.ToString()

    ' Restore the original culture
    Set CultureInfo.CurrentCulture = pvtCurrentCulture
End Sub

' The example displays the following output to the console:
'       5/1/2021 6:32:06 PM
'       01/05/2021 18:32:06
'       2021/05/01 18:32:06

