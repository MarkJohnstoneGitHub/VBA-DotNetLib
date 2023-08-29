Attribute VB_Name = "TestDTCreateFromDateTimeKind2"
'@IgnoreModule VariableNotUsed, EmptyMethod
'@TestModule
'@Folder("Unit Testing.System.DateTime.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 30, 2023
'@LastModified August 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References
' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.-ctor?view=netframework-4.8.1#system-datetime-ctor(system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-int32-system-datetimekind)

Option Compare Database
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

'@TestMethod("DateTime CreateFromDateTimeKind2")
Private Sub TestMethodDateTimeCreateFromDateTimeKind2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim testDateTime As DotNetLib.DateTime
    
    'Act:
    ' Initializes a new instance of the DateTime structure to the specified
    ' year, month, day, hour, minute, second, millisecond, and Coordinated
    ' Universal Time (UTC) or local time.
    Set testDateTime = DateTime.CreateFromDateTimeKind2(2010, 8, 18, 16, 32, 18, 500, DateTimeKind.DateTimeKind_Local)

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

