Attribute VB_Name = "TestDateTimeProperties"
'@TestModule
'@Folder("Unit Testing.System.DateTime.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 30, 2023
'@LastModified September 10, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.date?view=netframework-4.8.1

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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
    Assert.areEqual expectedDateOnly.Ticks, testDateTime.Ticks
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Assert.areEqual expectedDay, moment.Day
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Assert.areEqual DayOfWeek.DayOfWeek_Thursday, dt.DayOfWeek
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Assert.areEqual expectedHour, moment.Hour
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Assert.areEqual expectedMinute, moment.Minute
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Assert.areEqual expectedSecond, moment.Second
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
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
    Assert.areEqual expectedMillisecond, moment.Millisecond
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
