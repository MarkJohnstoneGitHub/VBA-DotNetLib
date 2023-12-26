VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface for Exception."
'@ModuleDescription("Interface for Exception.")
'@Folder "VBACorLib.ExceptionHandling"
'@Interface

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 24, 2023
'@LastModified August 24, 2023

Option Explicit

'@Description("Gets the error description.")
Public Property Get Description() As String
Attribute Description.VB_Description = "Gets the error description."
End Property

'@Description("Gets the error severity. I.e. warning or critical.")
Public Property Get ExceptionSeverity() As ExceptionSeverity
Attribute ExceptionSeverity.VB_Description = "Gets the error severity. I.e. warning or critical."
End Property

'@Description("Gets the error number.")
Public Property Get Number() As Long
Attribute Number.VB_Description = "Gets the error number."
End Property

'@Description("Gets the error source.")
Public Property Get source() As String
Attribute source.VB_Description = "Gets the error source."
End Property

'@Description("Gets the time stamp when the error was generated.")
Public Property Get TimeStamp() As Date
Attribute TimeStamp.VB_Description = "Gets the time stamp when the error was generated."
End Property
