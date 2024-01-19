Attribute VB_Name = "TextInfoToLowerExample"
'@Folder "Examples.System.Globalization.TextInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 5, 2023
'@LastModified September 5, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo.tolower?view=netframework-4.8.1#system-globalization-textinfo-tolower(system-string)

Option Explicit

' The following code example changes the casing of a string based on the English
' (United States) culture, with the culture name en-US.
Public Sub TextInfoToLower()
    ' Defines the string with mixed casing.
    Dim myString As String
    myString = "wAr aNd pEaCe"
    
    ' Creates a TextInfo based on the "en-US" culture.
    Dim myTI As DotNetLib.TextInfo
    Set myTI = CultureInfo.CreateFromName("en-US", False).TextInfo
    
    ' Changes a string to lowercase.
    Debug.Print """"; myString; """"; " to lowercase: "; myTI.ToLower(myString)
    
    ' Changes a string to uppercase.
    Debug.Print """"; myString; """"; " to uppercase: "; myTI.ToUpper(myString)
    
    ' Changes a string to titlecase.
    Debug.Print """"; myString; """"; " to titlecase: "; myTI.ToTitleCase(myString)
End Sub

'/*
'This code produces the following output.
'
'"wAr aNd pEaCe" to lowercase: war and peace
'"wAr aNd pEaCe" to uppercase: WAR AND PEACE
'"wAr aNd pEaCe" to titlecase: War And Peace
'
'*/
