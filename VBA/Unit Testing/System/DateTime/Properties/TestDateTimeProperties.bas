Attribute VB_Name = "TestDateTimeProperties"
'@TestModule
'@Folder("Unit Testing.System.DateTime.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 30, 2023
'@LastModified October 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.date?view=netframework-4.8.1

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Private DaysToMonth365() As Long
Private DaysToMonth366() As Long

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    ArrayEx.CreateInitialize1D DaysToMonth365, 0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365
    ArrayEx.CreateInitialize1D DaysToMonth366, 0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 366
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("DateTime Properties")
Private Sub TestingDateTimeDate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 6, 1, 7, 47, 0)
    Dim expectedDateOnly As DotNetLib.DateTime
    Set expectedDateOnly = DateTime.CreateFromDate(2008, 6, 1)
    
    'Act:
    ' Get date-only portion of date, without its time.
    Dim testDateTime As DotNetLib.DateTime
    Set testDateTime = date1.Date()

    'Assert:
    Assert.areEqual testDateTime.Ticks, expectedDateOnly.Ticks
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.Day property")
Private Sub TestingDateTimeDay()
Attribute TestingDateTimeDay.VB_Description = "Testing the DateTime.Day property"
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedDay As Long = 13
    Dim moment  As DotNetLib.DateTime
    Set moment = DateTime.CreateFromDateTime(1999, 1, 13, 3, 57, 32, 11)
    
    'Act:
    Dim dayResult As Long
    dayResult = moment.Day

    'Assert:
    Assert.areEqual moment.Day, expectedDay
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.DayOfWeek property")
Private Sub TestingDateTimeDayOfWeek()
Attribute TestingDateTimeDayOfWeek.VB_Description = "Testing the DateTime.DayOfWeek property"
    On Error GoTo TestFail
    
    'Arrange:
    Dim dt As DotNetLib.DateTime
    Set dt = DateTime.CreateFromDate(2003, 5, 1)
    
    'Act:

    'Assert:
    Assert.areEqual dt.DayOfWeek, DayOfWeek.DayOfWeek_Thursday
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.DayOfYear property")
Private Sub TestingDateTimeDayOfYear()
Attribute TestingDateTimeDayOfYear.VB_Description = "Testing the DateTime.DayOfYear property"
    On Error GoTo TestFail
    
    'Arrange:
    Dim dec31 As DotNetLib.DateTime
    Set dec31 = DateTime.CreateFromDate(2010, 12, 31)
    Dim expectedResult1 As Long
    expectedResult1 = 365
    Dim expectedResult2 As Long
    expectedResult2 = 366
    
    'Act:

    'Assert:
    Assert.areEqual dec31.DayOfYear, expectedResult1
    Assert.areEqual dec31.AddYears(2).DayOfYear, expectedResult2 'Test leap year
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.Hour property")
Private Sub TestingDateTimeHour()
Attribute TestingDateTimeHour.VB_Description = "Testing the DateTime.Hour property"
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedHour As Long = 3
    Dim moment  As DotNetLib.DateTime
    Set moment = DateTime.CreateFromDateTime(1999, 1, 13, expectedHour, 57, 32, 11)
    
    'Act:

    'Assert:
    Assert.areEqual moment.Hour, expectedHour
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.Kind property")
Private Sub TestingDateTimeKind()
Attribute TestingDateTimeKind.VB_Description = "Testing the DateTime.Kind property"
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedMillisecond As Long = 11
    Dim dt  As DotNetLib.DateTime
    Set dt = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)
    
    'Act:

    'Assert:
    Assert.areEqual dt.Kind, DateTimeKind.DateTimeKind_Local
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.Millisecond property")
Private Sub TestingDateTimeMillisecond()
Attribute TestingDateTimeMillisecond.VB_Description = "Testing the DateTime.Millisecond property"
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedMillisecond As Long = 11
    Dim moment  As DotNetLib.DateTime
    Set moment = DateTime.CreateFromDateTime(1999, 1, 13, 3, 57, 32, expectedMillisecond)
    
    'Act:

    'Assert:
    Assert.areEqual moment.Millisecond, expectedMillisecond
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.Minute property")
Private Sub TestingDateTimeMinute()
Attribute TestingDateTimeMinute.VB_Description = "Testing the DateTime.Minute property"
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedMinute As Long = 57
    Dim moment  As DotNetLib.DateTime
    Set moment = DateTime.CreateFromDateTime(1999, 1, 13, 3, expectedMinute, 32, 11)
    
    'Act:

    'Assert:
    Assert.areEqual moment.Minute, expectedMinute
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.Second property")
Private Sub TestingDateTimeSecond()
Attribute TestingDateTimeSecond.VB_Description = "Testing the DateTime.Second property"
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedSecond As Long = 32
    Dim moment  As DotNetLib.DateTime
    Set moment = DateTime.CreateFromDateTime(1999, 1, 13, 3, 57, expectedSecond, 11)
    
    'Act:

    'Assert:
    Assert.areEqual moment.SECOND, expectedSecond
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.Ticks property")
Private Sub TestingDateTimeTicks()
Attribute TestingDateTimeTicks.VB_Description = "Testing the DateTime.Ticks property"
    On Error GoTo TestFail
    
    'Arrange:
    Const expectedTicks As LongLong = "638299008000000000"
    
    Dim inputYear As Long
    inputYear = 2023
    Dim inputMonth As Long
    inputMonth = 9
    Dim inputDay As Long
    inputDay = 10
    Dim inputHour As Long
    inputHour = 0
    Dim inputMinute As Long
    inputMinute = 0
    Dim inputSecond As Long
    inputSecond = 0
    Dim inputMillisecond As Long
    inputMillisecond = 0
    
    Dim testDateTime As DotNetLib.DateTime
    Set testDateTime = DateTime.CreateFromDateTime(inputYear, inputMonth, inputDay, inputHour, inputMinute, inputSecond, inputMillisecond)
        
    'Act:
    Dim expectedTicksCalculated As LongLong
    expectedTicksCalculated = DateToTicks(inputYear, inputMonth, inputDay) + TimeToTicks(inputHour, inputMinute, inputSecond)
    expectedTicksCalculated = expectedTicksCalculated + inputMillisecond * TimeSpan.TicksPerMillisecond
    
    'Assert:
    Assert.areEqual testDateTime.Ticks, expectedTicksCalculated
    Assert.areEqual testDateTime.Ticks, expectedTicks
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'https://referencesource.microsoft.com/#mscorlib/system/timespan.cs,262a2ffd9cece820
Private Function TimeToTicks(ByVal pHour As Long, ByVal pMinute As Long, ByVal pSecond As Long) As LongLong
    Dim totalSeconds As LongLong
    totalSeconds = CLngLng(pHour) * 3600 + pMinute * 60 + pSecond
End Function

'https://referencesource.microsoft.com/#mscorlib/system/datetime.cs,256
Private Function DateToTicks(ByVal pYear As Long, ByVal pMonth As Long, ByVal pDay As Long) As LongLong
    Dim days() As Long
    days = IIf(DateTime.IsLeapYear(pYear), DaysToMonth366, DaysToMonth365)
    Dim y As Long
    y = pYear - 1
    DateToTicks = CLngLng(y * 365 + y / 4 - y / 100 + y / 400 + days(pMonth - 1) + pDay - 1) * TimeSpan.TicksPerDay
End Function

'@TestMethod("DateTime Properties")
'@Description("Testing the DateTime.TimeOfDay property")
Private Sub TestingDateTimeTimeOfDay()
Attribute TestingDateTimeTimeOfDay.VB_Description = "Testing the DateTime.TimeOfDay property"
    On Error GoTo TestFail
    
    'Arrange:
    Dim pvtHours As Long
    pvtHours = 9
    Dim pvtMinutes As Long
    pvtMinutes = 28
    Dim pvtSeconds As Long
    pvtSeconds = 0
    
    Dim dt  As DotNetLib.DateTime
    Set dt = DateTime.CreateFromDateTime(2013, 9, 14, pvtHours, pvtMinutes, pvtSeconds)
    
    'Act:
    Dim ts As DotNetLib.TimeSpan
    Set ts = dt.TimeOfDay
    

    'Assert:
    Assert.areEqual ts.Hours, pvtHours
    Assert.areEqual ts.Minutes, pvtMinutes
    Assert.areEqual ts.Seconds, pvtSeconds
    Assert.areEqual ts.Milliseconds, CLng(0)
    Assert.areEqual ts.Ticks, dt.Ticks
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub
