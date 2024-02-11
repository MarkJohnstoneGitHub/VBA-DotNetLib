VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestExamples.DataTypes")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 11, 2024
'@LastModified February 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'References
'https://stackoverflow.com/questions/56151269/is-there-an-implemented-function-similar-to-the-tostring-function-in-vba
'https://learn.microsoft.com/en-us/uwp/api/windows.foundation.istringable?view=winrt-22621
'
Option Explicit

Public Enum PersonStatus
    DeadOrAlive = 0
    Alive = 1
    Deceased = 2
End Enum


'Const PersonFormat  As String = "{4} {3}FirstName{3}: {0}, {3}LastName{3}: {1}, {3}DateOfBirth{3}: {2:s} {5}"
Const PersonFormat  As String = "{1} {0}FirstName{0}: {3}, {0}LastName{0}: {4}, {0}DateOfBirth{0}: {5:s}, {0}DateOfDeath{0}: {6:s} {2}"

Const Quote As String = """"
Const BraceOpen As String = "{"
Const BraceClose As String = "}"

Private Type TPerson
   FirstName As String
   LastName As String
   DateOfBirth As DotNetLib.DateTime
   DateOfDeath As DotNetLib.DateTime
End Type

Private this As TPerson

Implements IStringable

Friend Property Get Self() As Person
   Set Self = Me
End Property

Public Function Create(ByVal pFirstName As String, ByVal pLastName As String, ByVal pDateOfBirth As DotNetLib.DateTime, Optional ByVal pDateOfDeath As DotNetLib.DateTime) As Person
    With New Person
        .FirstName = pFirstName
        .LastName = pLastName
        Set .DateOfBirth = pDateOfBirth
        Set .DateOfDeath = pDateOfDeath
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

Public Property Get DateOfDeath() As DotNetLib.DateTime
    Set DateOfDeath = this.DateOfDeath
End Property

Friend Property Set DateOfDeath(ByVal value As DotNetLib.DateTime)
    Set this.DateOfDeath = value
End Property

'@DefaultMember
Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
    ToString = VBString.Format(PersonFormat, Quote, BraceOpen, BraceClose, _
                                this.FirstName, this.LastName, this.DateOfBirth, _
                                IIf(this.DateOfDeath Is Nothing, VBA.vbNullString, this.DateOfDeath))
End Property

Public Property Get IsAlive() As Boolean
    IsAlive = this.DateOfDeath Is Nothing
End Property

Public Function GetType() As DotNetLib.Type
    Dim obj As DotNetLib.Object
    Set obj = Objects.Create(Me)
    Set GetType = obj.GetType()
End Function

Private Property Get IStringable_ToString() As String
   IStringable_ToString = ToString()
End Property
