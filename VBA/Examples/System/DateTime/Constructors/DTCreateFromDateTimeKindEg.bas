Attribute VB_Name = "DTCreateFromDateTimeKindEg"
'@Folder "Examples.System.DateTime.Constructors"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 3, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-datetimekind)

Option Explicit

''
'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, DateTimeKind) constructor to instantiate a DateTime value.")
' Initializes a new instance of the DateTime structure to the specified
' year, month, day, hour, minute, second, and Coordinated Universal Time (UTC)
' or local time.
''
Public Sub DateTimeCreateFromDateTimeKind()
Attribute DateTimeCreateFromDateTimeKind.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, DateTimeKind) constructor to instantiate a DateTime value."
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)
    Debug.Print VBString.Format("{0} {1}", date1, DateTimeKindHelper.ToString(date1.Kind))
End Sub

' The example displays the following output, in this case for en-us culture:
'      8/18/2010 4:32:00 PM Local
