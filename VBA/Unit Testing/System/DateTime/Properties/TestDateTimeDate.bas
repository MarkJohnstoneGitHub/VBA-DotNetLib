Attribute VB_Name = "TestDateTimeDate"
'@TestModule
'@Folder("Unit Testing.System.DateTime.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 30, 2023
'@LastModified August 30, 2023

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
Private Sub TestingDateTimeDateOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim date1 As DotNetLib.DateTime
    Set date1 = DateTime.CreateFromDateTime(2008, 6, 1, 7, 47, 0)
    Dim expectedDateOnly As DotNetLib.DateTime
    Set expectedDateOnly = DateTime.CreateFromDate(2008, 6, 1)
    
    'Act:
    ' Get date-only portion of date, without its time.
    Dim testDateTime As DotNetLib.DateTime
    Set testDateTime = date1.DateOnly()

    'Assert:
    Assert.areEqual expectedDateOnly.Ticks, testDateTime.Ticks
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
