VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Exceptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a list of errors that occur during application execution used for logging errors."
'@ModuleDescription("Represents a list of errors that occur during application execution used for logging errors.")
'@Folder "VBACorLib.ExceptionHandling"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 23, 2023
'@LastModified August 24, 2023

'@Dependenicies
'   ExceptionSeverityEnum.bas
'   Exception.cls
'   IException.cls

'@References
'   https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
'   https://stackoverflow.com/a/65736562/10759363

'@TODO Possible keep a running total of error types and properties?

Option Explicit

Private Type TExceptions
    ExceptionList As Collection
End Type

Private this As TExceptions

'----------------------------------------------------------------------
'Constructors and destructors
'----------------------------------------------------------------------
Private Sub Class_Initialize()
    Set this.ExceptionList = New VBA.Collection
End Sub

''
'@Description("Returns the number of exceptions in the exception list.")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of exceptions in the exception list."
    Count = this.ExceptionList.Count
End Property

'@Enumerator
'@Description("Enumerator for exception list which returns variant keys.")
'@Remarks Attribute NewEnum.VB_UserMemId = -4 declares NewEnum as the enumerator
'         Attribute NewEnum.VB_MemberFlags = "40" declares NewEnum as a hidden property.
'@MemberAttribute VB_MemberFlags, "40"
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_Description = "Enumerator for exception list which returns variant keys."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = this.ExceptionList.[_NewEnum]
End Property

''
'@Description("Adds a new exception to the exception list.")
'@param Item The new Error item being added.
Public Sub Add(ByVal Item As Exception)
Attribute Add.VB_Description = "Adds a new exception to the exception list."
    this.ExceptionList.Add Item:=Item, Key:=str$(this.ExceptionList.Count + 1)
End Sub

'@Description("List of all warnings and critical errrors when exported.")
'@Reference https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
Public Function GetErrorListForType(ByVal errorType As ExceptionSeverity) As Exceptions
Attribute GetErrorListForType.VB_Description = "List of all warnings and critical errrors when exported."
    Dim pvtErrorList As Exceptions
    Set pvtErrorList = New Exceptions
    Dim varError As Variant
    Set varError = Nothing ' VBA Win 64 crashes without this
    For Each varError In this.ExceptionList
        Dim pvtError As IException
        Set pvtError = varError
        If pvtError.ExceptionSeverity = errorType Then
            pvtErrorList.Add pvtError
        End If
    Next
    Set GetErrorListForType = pvtErrorList
End Function

'@Description("Total warning errors.")
Public Function TotalWarnings() As Long
Attribute TotalWarnings.VB_Description = "Total warning errors."
    Dim warningsCount As Long
    Dim varError As Variant
    For Each varError In this.ExceptionList
        Dim pvtError As IException
        Set pvtError = varError
        If pvtError.ExceptionSeverity = ExceptionSeverity.Warning Then
            warningsCount = warningsCount + 1
        End If
    Next
    TotalWarnings = warningsCount
End Function

'@Description("Total critical errors.")
Public Function TotalCritical() As Long
Attribute TotalCritical.VB_Description = "Total critical errors."
    Dim criticalCount As Long
    Dim varError As Variant
    For Each varError In this.ExceptionList
        Dim pvtError As IException
        Set pvtError = varError
        If pvtError.ExceptionSeverity = ExceptionSeverity.Critical Then
            criticalCount = criticalCount + 1
        End If
    Next
    TotalCritical = criticalCount
End Function

'@Description("Output exception list to immediate window.")
Public Sub ErrorReport()
Attribute ErrorReport.VB_Description = "Output exception list to immediate window."
    Dim varError As Variant
    Dim pvtError As IException
    For Each varError In this.ExceptionList
        Set pvtError = varError
        Debug.Print pvtError.TimeStamp; " "; pvtError.Number; " "; pvtError.source; ", "; pvtError.Description
    Next
End Sub
