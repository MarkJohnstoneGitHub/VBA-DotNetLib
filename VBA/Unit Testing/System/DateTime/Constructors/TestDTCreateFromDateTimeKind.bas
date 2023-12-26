Attribute VB_Name = "TestDTCreateFromDateTimeKind"
'@IgnoreModule VariableNotUsed, EmptyMethod
'@TestModule
'@Folder("Unit Testing.System.DateTime.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 29, 2023
'@LastModified August 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-datetimekind)

'DateTime.CreateFromDateTimeKind(Int32 year, Int32 month, Int32 day, Int32 hour, Int32 minute, Int32 second, DateTimeKind kind)
'Initializes a new instance of the DateTime structure to the specified year, month,
'day, hour, minute, second, and Coordinated Universal Time (UTC) or local time.")
'
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
'   Kind DateTimeKind
'       One of the enumeration values that indicates whether year, month, day,
'       hour, minute and second specify a local time, Coordinated Universal Time
'       (UTC), or neither.
'
'@Returns DotNetLib.DateTime
'
'@Exceptions
'
'   ArgumentOutOfRangeException
'       year is less than 1 or greater than 9999.
'   -or-
'       month is less than 1 or greater than 12.
'   -or-
'       day is less than 1 or greater than the number of days in month.
'   -or-
'       hour is less than 0 or greater than 23.
'   -or-
'       minute is less than 0 or greater than 59.
'   -or-
'       second is less than 0 or greater than 59.
'
'   ArgumentException
'       kind is not one of the DateTimeKind values.
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

'@TestMethod("DateTime CreateFromDateTimeKind")
Private Sub TestMethodDateTimeCreateFromDateTimeKind()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    ' Initializes a new instance of the DateTime structure to the specified
    ' year, month, day, hour, minute, second, and Coordinated Universal Time (UTC)
    ' or local time.
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidYearMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidYear As Long = 10000
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(invalidYear, 8, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)
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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidYearMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidYear As Long = 0
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(invalidYear, 8, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateKindInvalidMonthMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMonth As Long = 13
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid month parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, invalidMonth, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)
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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateKindInvalidMonthMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMonth As Long = 0
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid month parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, invalidMonth, 18, 16, 32, 0, DateTimeKind.DateTimeKind_Local)
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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   day is less than 1 or greater than the number of days in month.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidDayMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidDay As Long = 32
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid day parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, invalidDay, 16, 32, 0, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   day is less than 1 or greater than the number of days in month.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidDayMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidDay As Long = 0
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid day parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, invalidDay, 16, 32, 0, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   hour is less than 0 or greater than 23.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidHourMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidHour As Long = 25
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid hour parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, invalidHour, 32, 0, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   hour is less than 0 or greater than 23.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidHourMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidHour As Long = -1
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid hour parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, invalidHour, 32, 0, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   minute is less than 0 or greater than 59
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidMinuteMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMinute As Long = 60
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid minute parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, invalidMinute, 0, DateTimeKind.DateTimeKind_Local)
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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   minute is less than 0 or greater than 59
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidMinuteMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidMinute As Long = -1
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid minute parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, invalidMinute, 0, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   second is less than 0 or greater than 59.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidSecondMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidSecond As Long = 60
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid second parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, invalidSecond, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions
'ArgumentOutOfRangeException
'   second is less than 0 or greater than 59.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidSecondMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidSecond As Long = -1
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Invalid second parameter raise ArgumentOutOfRangeException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, invalidSecond, DateTimeKind.DateTimeKind_Local)

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

'@TestMethod("DateTime CreateFromDateTimeKind")
'@Exceptions ArgumentException
'   kind is not one of the DateTimeKind values.
Private Sub TestMethodDateTimeCreateFromDateTimeKindInvalidKind()
    Const ExpectedError As Long = ArgumentException
    On Error GoTo TestFail
    
    'Arrange:
    Const invalidKind As Long = 3
    Dim testDateTime As DotNetLib.DateTime

    'Act:
    'Invalid kind parameter raise ArgumentException
    Set testDateTime = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, invalidKind)

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

