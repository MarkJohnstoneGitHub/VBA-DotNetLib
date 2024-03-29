VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Exception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents errors that occur during application execution."
'@ModuleDescription("Represents errors that occur during application execution.")
'@Folder "VBACorLib.ExceptionHandling"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 24, 2023
'@LastModified August 24, 2023

'@Dependenicies
'   IException.cls
'   ExceptionSeverityEnum.bas

Option Explicit

Implements IException

Private Type TException
    Number As Long
    source As String
    Description As String
    ExceptionSeverity As ExceptionSeverity
    TimeStamp As Date
End Type

Private this As TException

Friend Property Get Self() As Exception
   Set Self = Me
End Property

'@Description("Initializes a new instance of an exception.")
Public Function Create(ByVal errNumber As Long, ByVal errSource As String, ByVal errDescription As String, Optional ByVal errType As ExceptionSeverity = Unspecified) As IException
Attribute Create.VB_Description = "Initializes a new instance of an exception."
   With New Exception
      .Number = errNumber
      .source = errSource
      .Description = errDescription
      .TimeStamp = VBA.Now()
      .ExceptionSeverity = errType
      Set Create = .Self
   End With
End Function

'----------------------------------------------------------------------
' Properties
'----------------------------------------------------------------------
'@Description("Gets the error description.")
Public Property Get Description() As String
Attribute Description.VB_Description = "Gets the error description."
    Description = this.Description
End Property

Friend Property Let Description(ByVal value As String)
    this.Description = value
End Property

'@Description("Gets the error severity. I.e. warning or critical.")
Public Property Get ExceptionSeverity() As ExceptionSeverity
Attribute ExceptionSeverity.VB_Description = "Gets the error severity. I.e. warning or critical."
    ExceptionSeverity = this.ExceptionSeverity
End Property

Friend Property Let ExceptionSeverity(ByVal value As ExceptionSeverity)
    this.ExceptionSeverity = value
End Property


'@Description("Gets the error number.")
Public Property Get Number() As Long
Attribute Number.VB_Description = "Gets the error number."
    Number = this.Number
End Property

Friend Property Let Number(ByVal value As Long)
    this.Number = value
End Property

'@Description("Gets the error source.")
Public Property Get source() As String
Attribute source.VB_Description = "Gets the error source."
    source = this.source
End Property

Friend Property Let source(ByVal value As String)
    this.source = value
End Property

'@Description("Gets the time stamp when the error was generated.")
Public Property Get TimeStamp() As Date
Attribute TimeStamp.VB_Description = "Gets the time stamp when the error was generated."
    TimeStamp = this.TimeStamp
End Property

Friend Property Let TimeStamp(ByVal value As Date)
    this.TimeStamp = value
End Property

'----------------------------------------------------------------------
' Methods
'----------------------------------------------------------------------

'@Static
'@Description("Static. Throws the excpetion or raises the current VBA Err object.")
'@Reference https://stackoverflow.com/questions/730250/is-there-a-difference-between-throw-and-throw-ex
Public Sub Throw(Optional ByVal ex As Exception = Nothing)
Attribute Throw.VB_Description = "Static. Throws the excpetion or raises the current VBA Err object."
    If ex Is Nothing Then
        Err.Raise Err.Number, Err.source, Err.Description
    Else
        Err.Raise ex.Number, ex.source, ex.Description
    End If
End Sub

'@Static
'@Description("Static. Returns true for the error number attempting to catch or if not provided error number is not zero.")
'@TODO Overloads? For ErrorObj or IException?
Public Function Catch(Optional ByVal errNumber As Variant) As Boolean
Attribute Catch.VB_Description = "Static. Returns true for the error number attempting to catch or if not provided error number is not zero."
    If IsMissing(errNumber) Then
        If Err.Number <> 0 Then
            Catch = True
        End If
    ElseIf Err.Number = errNumber Then
        Catch = True
    End If
End Function

'@Static
'@Description("Static. Returns true if error number is 0.")
Public Function Try() As Boolean
Attribute Try.VB_Description = "Static. Returns true if error number is 0."
   If Err.Number = 0 Then
      Try = True
   End If
End Function

'----------------------------------------------------------------------
' Interface IException
'----------------------------------------------------------------------

Private Property Get IException_Number() As Long
    IException_Number = Number
End Property

Private Property Get IException_Source() As String
    IException_Source = source
End Property

Private Property Get IException_Description() As String
    IException_Description = Description
End Property

Private Property Get IException_TimeStamp() As Date
    IException_TimeStamp = TimeStamp
End Property

Private Property Get IException_ExceptionSeverity() As ExceptionSeverity
    IException_ExceptionSeverity = ExceptionSeverity
End Property
