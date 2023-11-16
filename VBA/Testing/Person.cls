VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Testing")
'@PredeclaredId

'References
'https://stackoverflow.com/questions/56151269/is-there-an-implemented-function-similar-to-the-tostring-function-in-vba
'https://learn.microsoft.com/en-us/uwp/api/windows.foundation.istringable?view=winrt-22621
'
Option Explicit
Const PersonFormat  As String = "{4} {3}FirstName{3}: {0}, {3}LastName{3}: {1}, {3}DateOfBirth{3}: {2:s} {5}"
Const Quote As String = """"
Const BraceOpen As String = "{"
Const BraceClose As String = "}"

Private Type TPerson
   FirstName As String
   LastName As String
   DateOfBirth As DotNetLib.DateTime
End Type

Private this As TPerson

Implements IStringable

Friend Property Get Self() As Person
   Set Self = Me
End Property

Public Function Create(ByVal pFirstName As String, ByVal pLastName As String, ByVal pDateOfBirth As DotNetLib.DateTime) As Person
    With New Person
        .FirstName = pFirstName
        .LastName = pLastName
        Set .DateOfBirth = pDateOfBirth
        Set Create = .Self
    End With
End Function

Public Property Get FirstName() As String
    FirstName = this.FirstName
End Property

Friend Property Let FirstName(ByVal value As String)
    this.FirstName = value
End Property

Public Property Get LastName() As String
    LastName = this.LastName
End Property

Friend Property Let LastName(ByVal value As String)
    this.LastName = value
End Property

Public Property Get DateOfBirth() As DotNetLib.DateTime
    Set DateOfBirth = this.DateOfBirth
End Property

Friend Property Set DateOfBirth(ByVal value As DotNetLib.DateTime)
    Set this.DateOfBirth = value
End Property

'@DefaultMember
Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
    ToString = VBAString.Format(PersonFormat, this.FirstName, this.LastName, this.DateOfBirth, Quote, BraceOpen, BraceClose)
End Property

Public Function GetType() As DotNetLib.Type
    Dim obj As DotNetLib.Object
    Set obj = Objects.Create(Me)
    Set GetType = obj.GetType()
End Function

Private Property Get IStringable_ToString() As String
   IStringable_ToString = ToString()
End Property
