Attribute VB_Name = "DateTimeConstructorExamples"
'Rubberduck annotations
'@Folder "VBADotNetLib.Examples.DateTime.Constructors"

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int64)

Option Explicit

'@Description("This example demonstrates the DateTime(Int64) constructor.")
Public Sub DateTimeCreateFromTicks()
Attribute DateTimeCreateFromTicks.VB_Description = "This example demonstrates the DateTime(Int64) constructor."
    ' Instead of using the implicit, default "G" date and time format string, we
    ' use a custom format string that aligns the results and inserts leading zeroes.
    Dim Format As String
    Format = "MM/dd/yyyy hh:mm:ss tt"
    
    'Create a DateTime for the maximum date and time using ticks.
    Dim dt1 As IDateTime
    Set dt1 = DateTime.CreateFromTicks(DateTime.MaxValue.Ticks)
    
    'Create a DateTime for the minimum date and time using ticks.
    Dim dt2 As IDateTime
    Set dt2 = DateTime.CreateFromTicks(DateTime.MinValue.Ticks)
    
    'Create a custom DateTime for 7/28/1979 at 10:35:05 PM
    Dim pvtTicks As LongLong
    pvtTicks = DateTime.CreateFromDateTime(1979, 7, 28, 22, 35, 5).Ticks
    Dim dt3 As IDateTime
    Set dt3 = DateTime.CreateFromTicks(pvtTicks)
    
    Debug.Print "1) The maximum date and time is " & dt1.ToString2(Format)
    Debug.Print "2) The minimum date and time is " & dt2.ToString2(Format)
    Debug.Print "3) The custom  date and time is " & dt3.ToString2(Format)
    Debug.Print "The custom date and time is created from " & pvtTicks & " ticks."
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, DateTimeKind) constructor to instantiate a DateTime value.")
Public Sub DateTimeCreateFromDateTimeKind()
Attribute DateTimeCreateFromDateTimeKind.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, DateTimeKind) constructor to instantiate a DateTime value."
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)
    
    Debug.Print date1.ToString() & " " & DateTimeKindHelper.ToString(date1.Kind)
    ' The example displays the following output, in this case for en-us culture:
    '      8/18/2010 4:32:00 PM Local
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Int32, DateTimeKind) constructor to instantiate a DateTime value.)"
'@Ignore UseMeaningfulName
Public Sub DateTimeCreateFromDateTimeKind2()
    '@Ignore UseMeaningfulName
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTimeKind2(2010, 8, 18, 16, 32, 18, 500, DateTimeKind.DateTimeKind_Local)
    
    Debug.Print date1.ToString2("M/dd/yyyy h:mm:ss.fff tt") & " " & DateTimeKindHelper.ToString(date1.Kind)
    ' The example displays the following output, in this case for en-us culture:
    ' 8/18/2010 4:32:18.500 PM Local
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value.")
Public Sub DateTimeCreateFromDateTime()
Attribute DateTimeCreateFromDateTime.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value."
    '@Ignore UseMeaningfulName
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 0)
    
    Debug.Print date1.ToString()
    ' The example displays the following output, in this case for en-us culture:
    '      8/18/2010 4:32:00 PM
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value.")
'@Ignore UseMeaningfulName
Public Sub DateTimeCreateFromDateTime2()
Attribute DateTimeCreateFromDateTime2.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32, Int32, Int32, Int32, Int32) constructor to instantiate a DateTime value."
    '@Ignore UseMeaningfulName
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 18, 500)
    
    Debug.Print date1.ToString2("M/dd/yyyy h:mm:ss.fff tt")
    ' The example displays the following output, in this case for en-us culture:
    ' 8/18/2010 4:32:18.500 PM
End Sub

'@Description("The following example uses the DateTime(Int32, Int32, Int32) constructor to instantiate a DateTime value. The example also illustrates that this overload creates a DateTime value whose time component equals midnight (or 0:00).")
Public Sub DateTimeCreateFromDate()
Attribute DateTimeCreateFromDate.VB_Description = "The following example uses the DateTime(Int32, Int32, Int32) constructor to instantiate a DateTime value. The example also illustrates that this overload creates a DateTime value whose time component equals midnight (or 0:00)."
    '@Ignore UseMeaningfulName
    Dim date1 As IDateTime
    Set date1 = DateTime.CreateFromDate(2010, 8, 18)
    
    Debug.Print date1.ToString()
    ' The example displays the following output:
    '      8/18/2010 12:00:00 AM
End Sub
