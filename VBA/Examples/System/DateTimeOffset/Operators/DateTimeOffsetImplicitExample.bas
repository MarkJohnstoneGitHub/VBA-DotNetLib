Attribute VB_Name = "DateTimeOffsetImplicitExample"
'@Folder "Examples.System.DateTimeOffset.Operators"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified January 11, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset.op_implicit?view=netframework-4.8.1#remarks

Option Explicit

''
' The Implicit method enables the compiler to automatically convert a DateTime object
' to a DateTimeOffset object without an explicit casting operator (in C#) or a call to
' a conversion function (in Visual Basic). It defines a widening conversion that does not
' involve data loss and does not throw an OverflowException.
' The Implicit method makes code such as the following possible:
''
Public Sub DateTimeOffsetImplicit()
    Dim timeWithOffset As DotNetLib.DateTimeOffset
    Set timeWithOffset = DateTimeOffset.Implicit(DateTime.CreateFromDateTime(2008, 7, 3, 18, 45, 0))
    Debug.Print timeWithOffset.ToString()

    Set timeWithOffset = DateTimeOffset.Implicit(DateTime.UtcNow)
    Debug.Print timeWithOffset.ToString()
    
    Set timeWithOffset = DateTimeOffset.Implicit(DateTime.SpecifyKind(DateTime.Now, DateTimeKind.DateTimeKind_Unspecified))
    Debug.Print timeWithOffset.ToString()
    
    Set timeWithOffset = DateTimeOffset.Implicit(DateTime.Addition(DateTime.CreateFromDateTime(2008, 7, 1, 2, 30, 0), TimeSpan.Create2(1, 0, 0, 0)))
    Debug.Print timeWithOffset.ToString()

    Set timeWithOffset = DateTimeOffset.Implicit(DateTime.CreateFromDateTime(2008, 1, 1, 2, 30, 0))
    Debug.Print timeWithOffset.ToString()
End Sub

' The example produces the following output if run on 3/20/2007
' at 6:25 PM on a computer in the U.S. Pacific Daylight Time zone:
'       7/3/2008 6:45:00 PM -07:00
'       3/21/2007 1:25:52 AM +00:00
'       3/20/2007 6:25:52 PM -07:00
'       7/2/2008 2:30:00 AM -07:00
'       1/1/2008 2:30:00 AM -08:00
'
' The last example shows automatic adaption to the U.S. Pacific Time
' for winter dates.
