Attribute VB_Name = "CultureInfoCurrentCultureEg"
'@Folder "Examples.System.Globalization.CultureInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 10, 2023
'@LastModified August 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.currentculture?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates how to change the CurrentCulture and CurrentUICulture of the current thread.")
Public Sub CultureInfoCurrentCulture()
Attribute CultureInfoCurrentCulture.VB_Description = "The following example demonstrates how to change the CurrentCulture and CurrentUICulture of the current thread."
    ' Display the name of the current culture.
    Debug.Print "CurrentCulture is "; CultureInfo.CurrentCulture.Name; "."

    ' Change the current culture to th-TH.
    Set CultureInfo.CurrentCulture = CultureInfo.Create4("th-TH", False)
    Debug.Print "CurrentCulture is now "; CultureInfo.CurrentCulture.Name; "."
    
    ' Display the name of the current UI culture.
    Debug.Print "CurrentUICulture is "; CultureInfo.CurrentUICulture.Name; "."
    
    ' Change the current UI culture to ja-JP.
    Set CultureInfo.CurrentUICulture = CultureInfo.Create4("ja-JP", False)
    Debug.Print "CurrentUICulture is now "; CultureInfo.CurrentUICulture.Name; "."
End Sub

' The example displays the following output:
'       CurrentCulture is en-US.
'       CurrentCulture is now th-TH.
'       CurrentUICulture is en-US.
'       CurrentUICulture is now ja-JP.

'@Description("The following example demonstrates how to change the CurrentCulture and CurrentUICulture of the current thread.")
Public Sub CultureInfoCurrentCultureV2()
Attribute CultureInfoCurrentCultureV2.VB_Description = "The following example demonstrates how to change the CurrentCulture and CurrentUICulture of the current thread."
    ' Display the name of the current culture.
    Debug.Print "CurrentCulture is "; CultureInfo.CurrentCulture.Name; "."
    Debug.Print CultureInfo.CurrentCulture.Calendar.ToString
End Sub
