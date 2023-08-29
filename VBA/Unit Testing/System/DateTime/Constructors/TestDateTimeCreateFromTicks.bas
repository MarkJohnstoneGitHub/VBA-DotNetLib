Attribute VB_Name = "TestDateTimeCreateFromTicks"
Attribute VB_Description = "Unit testing for DateTime.CreateFromTicks(LongLong ticks, DateTimeKind kind)."
'@ModuleDescription "Unit testing for DateTime.CreateFromTicks(LongLong ticks, DateTimeKind kind)."
'@TestModule
'@Folder("Unit Testing.System.DateTime.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 29, 2023
'@LastModified August 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int64)
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int64-system-datetimekind)

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

'--------------------------------------------------------------------
'Constructors
'--------------------------------------------------------------------

'@TestMethod("DateTime CreateFromTicks")
Private Sub TestMethodDateTimeCreateFromTicks()
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime
    Dim pvtTicks As LongLong
    pvtTicks = DateTime.CreateFromDateTime(1979, 7, 28, 22, 35, 5).Ticks
    
    'Act:
    'Create a DateTime for the maximum date and time using ticks.
    Set TestDateTime = DateTime.CreateFromTicks(DateTime.MaxValue.Ticks)
    
    'Create a DateTime for the minimum date and time using ticks.
    Set TestDateTime = DateTime.CreateFromTicks(DateTime.MinValue.Ticks)
    
    'Create a custom DateTime for 7/28/1979 at 10:35:05 PM
    Set TestDateTime = DateTime.CreateFromTicks(pvtTicks)
    
    'Initializes a new instance of the DateTime structure to a specified number of ticks
    'and to Coordinated Universal Time (UTC) or local time.
    Set TestDateTime = DateTime.CreateFromTicks(DateTime.MaxValue.Ticks, DateTimeKind.DateTimeKind_Local)
    Set TestDateTime = DateTime.CreateFromTicks(pvtTicks, DateTimeKind.DateTimeKind_Utc)
    
    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("DateTime CreateFromTicks")
'@Exception ArgumentOutOfRangeException
'   ticks is less than DateTime.MinValue or greater than DateTime.MaxValue.
Private Sub TestMethodDateTimeCreateFromTicksInvalidTicksMax()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Create a DateTime for the maximum date and time using ticks and increament to cause ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromTicks(DateTime.MaxValue.Ticks + 1)

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

'@TestMethod("DateTime CreateFromTicks")
'@Exception ArgumentOutOfRangeException
'   ticks is less than DateTime.MinValue or greater than DateTime.MaxValue.
Private Sub TestMethodDateTimeCreateFromTicksInvalidTicksMin()
    Const ExpectedError As Long = ArgumentOutOfRangeException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime

    'Act:
    'Create a DateTime for the minimum date and time using ticks and decrement to raise ArgumentOutOfRangeException
    Set TestDateTime = DateTime.CreateFromTicks(DateTime.MinValue.Ticks - 1)

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

'@TestMethod("DateTime CreateFromTicks")
'@Exceptions ArgumentException
'   kind is not one of the DateTimeKind values.
Private Sub TestMethodDateTimeCreateFromTicksInvalidKind()
    Const ExpectedError As Long = ArgumentException
    On Error GoTo TestFail
    
    'Arrange:
    Dim TestDateTime As DotNetLib.DateTime
    Dim pvtTicks As LongLong
    pvtTicks = DateTime.CreateFromDateTime(1979, 7, 28, 22, 35, 5).Ticks

    'Act:
    'Invalid kind parameter raise ArgumentException
    Set TestDateTime = DateTime.CreateFromTicks(pvtTicks, 3)

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

