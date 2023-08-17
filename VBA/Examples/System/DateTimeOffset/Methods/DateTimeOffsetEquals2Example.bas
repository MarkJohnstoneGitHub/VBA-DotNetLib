Attribute VB_Name = "DateTimeOffsetEquals2Example"
'@Folder "Examples.System.DateTimeOffset.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 21, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.equals?view=netframework-4.8.1#system-datetimeoffset-equals(system-object)

Option Explicit

'@Description("The following example indicates whether the current DateTimeOffset object is equal to several other DateTimeOffset objects, as well as to a null reference and a DateTime object.")
Public Sub DateTimeOffsetEquals2()
Attribute DateTimeOffsetEquals2.VB_Description = "The following example indicates whether the current DateTimeOffset object is equal to several other DateTimeOffset objects, as well as to a null reference and a DateTime object."
    Dim firstTime As IDateTimeOffset
    Set firstTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, TimeSpan.Create(-7, 0, 0))
    
    Dim secondTime As Object
    Set secondTime = firstTime
    Debug.Print firstTime.ToString() & " = " & _
                secondTime.ToString() & ": " & _
                firstTime.Equals2(secondTime)
             
    Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 6, 45, 0, TimeSpan.Create(-6, 0, 0))
    Debug.Print firstTime.ToString() & " = " & _
                secondTime.ToString() & ": " & _
                firstTime.Equals2(secondTime)
    
    Set secondTime = DateTimeOffset.CreateFromDateTimeParts(2007, 9, 1, 8, 45, 0, TimeSpan.Create(-5, 0, 0))
    Debug.Print firstTime.ToString() & " = " & _
                secondTime.ToString() & ": " & _
                firstTime.Equals2(secondTime)
             
    Set secondTime = Nothing
    Dim outString As String
    outString = firstTime.ToString() & " = "
    outString = outString + IIf(Not secondTime Is Nothing, secondTime, VBA.vbNullString) & ": "
    outString = outString & firstTime.Equals2(secondTime)
    
    Debug.Print firstTime.ToString() & " = " & _
                IIf(Not secondTime Is Nothing, secondTime, VBA.vbNullString) & ": " & _
                firstTime.Equals2(secondTime)
    
    Set secondTime = DateTime.CreateFromDateTime(2007, 9, 1, 6, 45, 0)
    Debug.Print firstTime.ToString() & " = " & _
                secondTime.ToString() & ": " & _
                firstTime.Equals2(secondTime)
End Sub

' The example displays the following output to the console:
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM -07:00: True
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM -06:00: False
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 8:45:00 AM -05:00: True
'       9/1/2007 6:45:00 AM -07:00 = : False
'       9/1/2007 6:45:00 AM -07:00 = 9/1/2007 6:45:00 AM: False
