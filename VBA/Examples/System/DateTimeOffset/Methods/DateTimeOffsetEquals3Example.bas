Attribute VB_Name = "DateTimeOffsetEquals3Example"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.equals?view=netframework-4.8.1#system-datetimeoffset-equals(system-datetimeoffset-system-datetimeoffset)

Option Explicit

''
' The following example illustrates calls to the Equals(DateTimeOffset, DateTimeOffset)
' method to test various pairs of DateTimeOffset objects for equality.
''
Public Sub DateTimeOffsetEquals3()
    Dim firstTime As DotNetLib.DateTimeOffset
    Set firstTime = DateTimeOffset.CreateFromDateTimeParts(2007, 11, 15, 11, 35, 0, DateTimeOffset.Now.Offset)
    Dim secondTime As DotNetLib.DateTimeOffset
    Set secondTime = firstTime
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                firstTime, secondTime, _
                                DateTimeOffset.Equals(firstTime, secondTime))
    
    ' The value of firstTime remains unchanged
    Set secondTime = DateTimeOffset.CreateFromDateTime2(firstTime.DateTime, TimeSpan.FromHours(firstTime.Offset.Hours + 1))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                firstTime, secondTime, _
                                DateTimeOffset.Equals(firstTime, secondTime))
                  
    ' value of firstTime remains unchanged
    Set secondTime = DateTimeOffset.CreateFromDateTime2(DateTime.Addition(firstTime.DateTime, TimeSpan.FromHours(1)), TimeSpan.FromHours(firstTime.Offset.Hours + 1))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                firstTime, secondTime, _
                                DateTimeOffset.Equals(firstTime, secondTime))
End Sub

'  The example produces the following output:
'        11/15/2007 11:35:00 AM -07:00 = 11/15/2007 11:35:00 AM -07:00: True
'        11/15/2007 11:35:00 AM -07:00 = 11/15/2007 11:35:00 AM -06:00: False
'        11/15/2007 11:35:00 AM -07:00 = 11/15/2007 12:35:00 PM -06:00: True

