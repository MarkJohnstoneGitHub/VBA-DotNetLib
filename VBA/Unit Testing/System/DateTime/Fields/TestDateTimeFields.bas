Attribute VB_Name = "TestDateTimeFields"
'@TestModule
'@Folder("Unit Testing.System.DateTime.Fields")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 30, 2023
'@LastModified August 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.maxvalue?view=netframework-4.8.1
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.minvalue?view=netframework-4.8.1

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

'@TestMethod("DateTime Fields")
Private Sub TestDateTimeMaxValue()
    On Error GoTo TestFail
    
    'Arrange:
    Const MaxTicks As LongLong = "3155378975999999999"
    Dim testDateTime As DotNetLib.DateTime
    'Act:
    Set testDateTime = DateTime.MaxValue

    'Assert:
    Assert.IsTrue testDateTime.Ticks = MaxTicks

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("DateTime Fields")
Private Sub TestDateTimeMinValue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim MinDateTime As DotNetLib.DateTime
    Set MinDateTime = New DotNetLib.DateTime
    Dim testDateTime As DotNetLib.DateTime
    'Act:
    Set testDateTime = DateTime.MinValue

    'Assert:
    Assert.IsTrue testDateTime.Ticks = MinDateTime.Ticks

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.number & " - " & Err.Description
    Resume TestExit
End Sub

