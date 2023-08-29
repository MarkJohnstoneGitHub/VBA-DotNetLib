Attribute VB_Name = "TestDateTimeCreateFromDateTime"
Attribute VB_Description = "Unit testing for DateTime.CreateFromDateTime(Int32 year, Int32 month, Int32 day, Int32 hour, Int32 minute, Int32 second, Int32 [millisecond])"
'@IgnoreModule VariableNotUsed, EmptyMethod
'@TestModule
'@ModuleDescription "Unit testing for DateTime.CreateFromDateTime(Int32 year, Int32 month, Int32 day, Int32 hour, Int32 minute, Int32 second, Int32 [millisecond])"
'@Folder("Unit Testing.System.DateTime.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 29, 2023
'@LastModified August 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32)
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-int32)

'Initializes a new instance of the DateTime structure to the specified year, month, day, hour, minute, second, and/or millisecond.
'@Parameters
'   Year Int32
'       The year (1 through 9999).
'   Month Int32
'       The month (1 through 12).
'   Day Int32
'       The day (1 through the number of days in month).
'   Hour Int32
'       The hours (0 through 23).
'   Minute Int32
'       The minutes (0 through 59).
'   Second Int32
'       The seconds (0 through 59).
'   Millisecond Int32
'       The milliseconds (0 through 999).
'
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
'-or-
'   month is less than 1 or greater than 12.
'-or-
'   day is less than 1 or greater than the number of days in month.
'-or-
'   hour is less than 0 or greater than 23.
'-or-
'   minute is less than 0 or greater than 59.
'-or-
'   second is less than 0 or greater than 59.
'-or-
'   millisecond is less than 0 or greater than 999.

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

'@TestMethod("DateTime CreateFromDateTime")
Private Sub TestMethodDateTimeCreateFromDateTime()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    ' Initializes a new instance of the DateTime structure to the specified
    ' year, month, day, hour, minute, and second.
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 0)
    
    ' Initializes a new instance of the DateTime structure to the specified
    ' year, month, day, hour, minute, second, and millisecond.
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 18, 500)

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidYearMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidYear As Long = 10000
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(invalidYear, 8, 18, 16, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidYearMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidYear As Long = 0
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(invalidYear, 8, 18, 16, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateInvalidMonthMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMonth As Long = 13
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid month parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, invalidMonth, 18, 16, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateInvalidMonthMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMonth As Long = 0
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid month parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, invalidMonth, 18, 16, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   day is less than 1 or greater than the number of days in month.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidDayMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidDay As Long = 32
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid day parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, invalidDay, 16, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   day is less than 1 or greater than the number of days in month.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidDayMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidDay As Long = 0
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid day parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, invalidDay, 16, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   hour is less than 0 or greater than 23.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidHourMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidHour As Long = 25
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid hour parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, invalidHour, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   hour is less than 0 or greater than 23.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidHourMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidHour As Long = -1
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid hour parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, invalidHour, 32, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   minute is less than 0 or greater than 59
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidMinuteMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMinute As Long = 60
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid minute parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, invalidMinute, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   minute is less than 0 or greater than 59
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidMinuteMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMinute As Long = -1
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid minute parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, invalidMinute, 18, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   second is less than 0 or greater than 59.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidSecondMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidSecond As Long = 60
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid second parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, invalidSecond, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   second is less than 0 or greater than 59.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidSecondMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidSecond As Long = -1
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid second parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, invalidSecond, 500)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   millisecond is less than 0 or greater than 999.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidMillisecondMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMillisecond As Long = 1000
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid millisecond parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 18, invalidMillisecond)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod("DateTime CreateFromDateTime")
'@Exceptions
'ArgumentOutOfRangeException
'   millisecond is less than 0 or greater than 999.
Private Sub TestMethodDateTimeCreateFromDateTimeInvalidMillisecondMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMillisecond As Long = -1
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid millisecond parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 18, invalidMillisecond)

Assert:
    Assert.Fail "Expected error was not raised"

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub
