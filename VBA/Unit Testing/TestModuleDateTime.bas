Attribute VB_Name = "TestModuleDateTime"
'@ModuleDescription "Unit testing for DotNetLib.DateTime and DateTime factoy and static methods."
' @References
' https://github.com/rubberduck-vba/Rubberduck/wiki/Unit-Testing
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 29, 2023
'@LastModified August 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

Option Explicit
Option Private Module

'@TestModule
'@Folder "Unit Testing.System.DateTime"

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

'@TestMethod("DateTime Constructors")
Private Sub TestMethodDateTimeCreateFromDate()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    Set testDateTime = DateTime.CreateFromDate(2010, 8, 18)
    Set testDateTime = DateTime.CreateFromDate(2023, 8, 29)
    Set testDateTime = DateTime.CreateFromDate(1999, 1, 1)

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Constructors")
Private Sub TestMethodDateTimeCreateFromTicks()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    'Create a DateTime for the maximum date and time using ticks.
    Set testDateTime = DateTime.CreateFromTicks(DateTime.MaxValue.Ticks)
    
    'Create a DateTime for the minimum date and time using ticks.
    Set testDateTime = DateTime.CreateFromTicks(DateTime.MinValue.Ticks)
    
    'Create a custom DateTime for 7/28/1979 at 10:35:05 PM
    Dim pvtTicks As LongLong
    pvtTicks = DateTime.CreateFromDateTime(1979, 7, 28, 22, 35, 5).Ticks
    Set testDateTime = DateTime.CreateFromTicks(pvtTicks)

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


