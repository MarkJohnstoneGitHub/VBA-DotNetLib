Attribute VB_Name = "DateTimeGetDateTimeFormatsEg"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified July 11, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.getdatetimeformats?view=netframework-4.8.1

Option Explicit

'@Description("The following example demonstrates the DateTime.GetDateTimeFormats() method. It displays the string representation of a date using all possible standard date and time formats in the computer's current culture, which in this case is en-US."
Public Sub DateTimeGetDateTimeFormats()
   '@Ignore UseMeaningfulName
   Dim july28 As DateTime
   Set july28 = DateTime.CreateFromDateTime(2009, 7, 28, 5, 23, 15, 16)
   
   Dim july28Formats() As String
   july28Formats = july28.GetDateTimeFormats()
   
   ' Print out july28 in all DateTime formats using the default culture.
   Dim varFormat As Variant
   For Each varFormat In july28Formats
      Debug.Print varFormat
   Next

'The example displays the following output:
'
'7/28/2009
'7/28/09
'07/28/09
'07/28/2009
'09/07/28
'2009-07-28
'28-Jul-09
'Tuesday, July 28, 2009
'July 28, 2009
'Tuesday, 28 July, 2009
'28 July , 2009
'Tuesday, July 28, 2009 5:23 AM
'Tuesday, July 28, 2009 05:23 AM
'Tuesday, July 28, 2009 5:23
'Tuesday, July 28, 2009 05:23
'July 28, 2009 5:23 AM
'July 28, 2009 05:23 AM
'July 28, 2009 5:23
'July 28, 2009 05:23
'Tuesday, 28 July, 2009 5:23 AM
'Tuesday, 28 July, 2009 05:23 AM
'Tuesday, 28 July, 2009 5:23
'Tuesday, 28 July, 2009 05:23
'28 July, 2009 5:23 AM
'28 July, 2009 05:23 AM
'28 July, 2009 5:23
'28 July, 2009 05:23
'Tuesday, July 28, 2009 5:23:15 AM
'Tuesday, July 28, 2009 05:23:15 AM
'Tuesday, July 28, 2009 5:23:15
'Tuesday, July 28, 2009 05:23:15
'July 28, 2009 5:23:15 AM
'July 28, 2009 05:23:15 AM
'July 28, 2009 5:23:15
'July 28, 2009 05:23:15
'Tuesday, 28 July, 2009 5:23:15 AM
'Tuesday, 28 July, 2009 05:23:15 AM
'Tuesday, 28 July, 2009 5:23:15
'Tuesday, 28 July, 2009 05:23:15
'28 July, 2009 5:23:15 AM
'28 July, 2009 05:23:15 AM
'28 July, 2009 5:23:15
'28 July, 2009 05:23:15
'7/28/2009 5:23 AM
'7/28/2009 05:23 AM
'7/28/2009 5:23
'7/28/2009 05:23
'7/28/09 5:23 AM
'7/28/09 05:23 AM
'7/28/09 5:23
'7/28/09 05:23
'07/28/09 5:23 AM
'07/28/09 05:23 AM
'07/28/09 5:23
'07/28/09 05:23
'07/28/2009 5:23 AM
'07/28/2009 05:23 AM
'07/28/2009 5:23
'07/28/2009 05:23
'09/07/28 5:23 AM
'09/07/28 05:23 AM
'09/07/28 5:23
'09/07/28 05:23
'2009-07-28 5:23 AM
'2009-07-28 05:23 AM
'2009-07-28 5:23
'2009-07-28 05:23
'28-Jul-09 5:23 AM
'28-Jul-09 05:23 AM
'28-Jul-09 5:23
'28-Jul-09 05:23
'7/28/2009 5:23:15 AM
'7/28/2009 05:23:15 AM
'7/28/2009 5:23:15
'7/28/2009 05:23:15
'7/28/09 5:23:15 AM
'7/28/09 05:23:15 AM
'7/28/09 5:23:15
'7/28/09 05:23:15
'07/28/09 5:23:15 AM
'07/28/09 05:23:15 AM
'07/28/09 5:23:15
'07/28/09 05:23:15
'07/28/2009 5:23:15 AM
'07/28/2009 05:23:15 AM
'07/28/2009 5:23:15
'07/28/2009 05:23:15
'09/07/28 5:23:15 AM
'09/07/28 05:23:15 AM
'09/07/28 5:23:15
'09/07/28 05:23:15
'2009-07-28 5:23:15 AM
'2009-07-28 05:23:15 AM
'2009-07-28 5:23:15
'2009-07-28 05:23:15
'28-Jul-09 5:23:15 AM
'28-Jul-09 05:23:15 AM
'28-Jul-09 5:23:15
'28-Jul-09 05:23:15
'July 28
'July 28
'2009-07-28T05:23:15.0160000
'2009-07-28T05:23:15.0160000
'Tue, 28 Jul 2009 05:23:15 GMT
'Tue, 28 Jul 2009 05:23:15 GMT
'2009-07-28T05:23:15
'5:23 AM
'05:23 AM
'5:23
'05:23
'5:23:15 AM
'05:23:15 AM
'5:23:15
'05:23:15
'2009-07-28 05:23:15Z
'Tuesday, July 28, 2009 12:23:15 PM
'Tuesday, July 28, 2009 12:23:15 PM
'Tuesday, July 28, 2009 12:23:15
'Tuesday, July 28, 2009 12:23:15
'July 28, 2009 12:23:15 PM
'July 28, 2009 12:23:15 PM
'July 28, 2009 12:23:15
'July 28, 2009 12:23:15
'Tuesday, 28 July, 2009 12:23:15 PM
'Tuesday, 28 July, 2009 12:23:15 PM
'Tuesday, 28 July, 2009 12:23:15
'Tuesday, 28 July, 2009 12:23:15
'28 July, 2009 12:23:15 PM
'28 July, 2009 12:23:15 PM
'28 July, 2009 12:23:15
'28 July, 2009 12:23:15
'July , 2009
'July , 2009
   
End Sub

