Attribute VB_Name = "DateTimeOffsetEqualsExample"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 20, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.equals?view=netframework-4.8.1#system-datetimeoffset-equals(system-datetimeoffset)

Option Explicit

''
' The following example illustrates calls to the Equals(DateTimeOffset) method to
' test DateTimeOffset objects for equality with the current DateTimeOffset object.
''
Public Sub DateTimeOffsetEquals()
    Dim firstTime As DotNetLib.DateTimeOffset
    Set firstTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, _
                                        TimeSpan.Create(-7, 0, 0))
    
    Dim secondTime As DotNetLib.DateTimeOffset
    Set secondTime = firstTime
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                firstTime, secondTime, _
                                firstTime.Equals(secondTime))
             
    Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, _
                                        TimeSpan.Create(-6, 0, 0))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                firstTime, secondTime, _
                                firstTime.Equals(secondTime))
    
    Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 8, 45, 0, _
                                        TimeSpan.Create(-5, 0, 0))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                                firstTime, secondTime, _
                                firstTime.Equals(secondTime))
End Sub

' The example displays the following output to the console:
'      9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM -07:00: True
'      9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM -06:00: False
'      9/1/2007 6:45:00 AM -07:00 = 9/1/2007 8:45:00 AM -05:00: True

