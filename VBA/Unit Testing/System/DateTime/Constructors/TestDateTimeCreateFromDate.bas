Attribute VB_Name = "TestDateTimeCreateFromDate"
'@TestModule
'@Folder("Unit Testing.System.DateTime.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 29, 2023
'@LastModified August 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32)

' Initializes a new instance of the DateTime structure to the specified year, month, and day.
' DateTime.CreateFromDate(Int32 year, Int32 month, Int32 day)
'@Parameters
'   Year Int32
'       The year (1 through 9999).
'   Month Int32
'       The month (1 through 12).
'   Day Int32
'       The day (1 through the number of days in month).
'
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
'
'-or-
'
'   month is less than 1 or greater than 12.
'
'-or-
'
'   day is less than 1 or greater than the number of days in month.

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

'@TestMethod("DateTime CreateFromDate")
Private Sub TestMethodDateTimeCreateFromDate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime
    
    'Act:
    Set TestDateTime = DateTime.CreateFromDate(2010, 8, 18)
    Set TestDateTime = DateTime.CreateFromDate(2023, 8, 29)
    Set TestDateTime = DateTime.CreateFromDate(1999, 1, 1)

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime CreateFromDate")
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
Private Sub TestMethodDateTimeCreateFromDateInvalidYearMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromDate(-1, 8, 18)

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

'@TestMethod("DateTime CreateFromDate")
'@Exceptions
'ArgumentOutOfRangeException
'   year is less than 1 or greater than 9999.
Private Sub TestMethodDateTimeCreateFromDateInvalidYearMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromDate(10000, 8, 18)

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

'@TestMethod("DateTime CreateFromDate")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateInvalidMonthMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromDate(2023, 0, 18)

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

'@TestMethod("DateTime CreateFromDate")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateInvalidMonthMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromDate(2023, 13, 18)

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

'@TestMethod("DateTime CreateFromDate")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateInvalidDayMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromDate(2023, 8, 0)

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

'@TestMethod("DateTime CreateFromDate")
'@Exceptions
'ArgumentOutOfRangeException
'   month is less than 1 or greater than 12.
Private Sub TestMethodDateTimeCreateFromDateInvalidDayMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Invalid year parameter raise ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromDate(2023, 8, 32)

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
