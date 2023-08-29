Attribute VB_Name = "TestDateTimeCreateFromDateTime"
'@TestModule
'@Folder("Unit Testing.System.DateTime.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 29, 2023
'@LastModified August 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@References

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
    Dim TestDateTime As DotNetLib.DateTime
    
    'Act:
    
    ' Initializes a new instance of the DateTime structure to the specified
    ' year, month, day, hour, minute, and second.
    Set TestDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 0)
    
    ' Initializes a new instance of the DateTime structure to the specified
    ' year, month, day, hour, minute, second, and millisecond.
    Set TestDateTime = DateTime.CreateFromDateTime(2010, 8, 18, 16, 32, 18, 500)

    'Assert:
    Assert.Succeed

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
