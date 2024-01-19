Attribute VB_Name = "DateTimeOffsetEquals2Example"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified January 10, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.equals?view=netframework-4.8.1#system-datetimeoffset-equals(system-object)

Option Explicit

''
' The following example indicates whether the current DateTimeOffset object is
' equal to several other DateTimeOffset objects, as well as to a null reference
' and a DateTime object.
''
Public Sub DateTimeOffsetEquals2()
    Dim firstTime As DotNetLib.DateTimeOffset
    Set firstTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, TimeSpan.Create(-7, 0, 0))
    
    Dim secondTime As Object
    Set secondTime = firstTime
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                     firstTime, secondTime, _
                     firstTime.Equals2(secondTime))
             
    Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, TimeSpan.Create(-6, 0, 0))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                     firstTime, secondTime, _
                     firstTime.Equals2(secondTime))
    
    Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 8, 45, 0, TimeSpan.Create(-5, 0, 0))
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                     firstTime, secondTime, _
                     firstTime.Equals2(secondTime))
             
    Set secondTime = Nothing
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                     firstTime, secondTime, _
                     firstTime.Equals2(secondTime))

    Set secondTime = DateTime.CreateFromDateTime(2007, 9, 1, 6, 45, 0)
    Debug.Print VBString.Format("{0} = {1}: {2}", _
                     firstTime, secondTime, _
                     firstTime.Equals2(secondTime))
End Sub

' The example displays the following output to the console:
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM -07:00: True
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM -06:00: False
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 8:45:00 AM -05:00: True
'       9/1/2007 6:45:00 AM -07:00 = : False
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM: False

