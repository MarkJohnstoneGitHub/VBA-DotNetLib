VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersonFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestExamples.System.Array.Methods")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 11, 2024
'@LastModified February 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

Option Explicit

Implements DotNetLib.IPredicate

Private Type TPersonFilter
    mFirstName As String
    mLastName As String
    mDateOfBirth As DotNetLib.DateTime
    mAge As Long
    mDateTimeNow As DotNetLib.DateTime
    mStatus As PersonStatus
End Type

Private mPredicate As DotNetLib.Predicate

Private this As TPersonFilter

Private Sub Class_Initialize()
    Set this.mDateTimeNow = DateTime.Now().Date
End Sub

'----------------------------------------------------------------------
' Factory Methods
'----------------------------------------------------------------------

''
' Creates a predicate for a Person with filters for a persons containing first name or last name,
' date of birth greater then a date or greater than a age.
' String comparisions for first and last name case is ignored by converting to lower case.
''
Public Function Create(Optional ByVal pFirstName As String, Optional ByVal pLastName As String, Optional ByVal DOB As DotNetLib.DateTime, Optional ByVal pAge As Long, Optional ByVal pStatus As PersonStatus) As PersonFilter
    With New PersonFilter
        .FirstName = LCase$(pFirstName)
        .LastName = LCase$(pLastName)
        Set .DateOfBirth = DOB
        .Age = pAge
        .Status = pStatus
        Set .CallBack = Predicate.Create(.Self)
        Set Create = .Self
    End With
End Function

'----------------------------------------------------------------------
' Properties
'----------------------------------------------------------------------

Friend Property Get Self() As PersonFilter
    Set Self = Me
End Property

'Returns the predicate callback
Public Property Get CallBack() As DotNetLib.Predicate
    Set CallBack = mPredicate
End Property

Friend Property Set CallBack(ByVal pPredicate As DotNetLib.Predicate)
    Set mPredicate = pPredicate
End Property

Public Property Get FirstName() As String
    FirstName = this.mFirstName
End Property

Public Property Let FirstName(ByVal pFirstName As String)
    this.mFirstName = pFirstName
End Property

Public Property Get LastName() As String
    LastName = this.mLastName
End Property

Public Property Let LastName(ByVal pLastName As String)
    this.mLastName = pLastName
End Property

Public Property Get DateOfBirth() As DotNetLib.DateTime
    Set DateOfBirth = this.mDateOfBirth
End Property

Public Property Set DateOfBirth(ByVal pDOB As DotNetLib.DateTime)
    Set this.mDateOfBirth = pDOB
End Property

Public Property Get Age() As Long
    Age = this.mAge
End Property

Public Property Let Age(ByVal pAgeFilter As Long)
    this.mAge = pAgeFilter
    Set this.mDateTimeNow = DateTime.Now().Date
End Property

Public Property Get Status() As PersonStatus
    Status = this.mStatus
End Property

Public Property Let Status(ByVal pPersonStatus As PersonStatus)
    this.mStatus = pPersonStatus
End Property

'----------------------------------------------------------------------
' Methods
'----------------------------------------------------------------------

Public Sub RefreshNow()
    Set this.mDateTimeNow = DateTime.Now().Date
End Sub

Public Sub Reset()
    this.mFirstName = VBA.vbNullString
    this.mLastName = VBA.vbNullString
    Set this.mDateOfBirth = Nothing
    this.mAge = 0
    this.mStatus = PersonStatus.DeadOrAlive
    Set this.mDateTimeNow = DateTime.Now().Date
End Sub

''
' Finds a person match according to the person filters specified.
' Person filter matches if a person which contains first name, or last name ignoring case.
' Also if date of birth is greater than or equal to DOB filter specified or their age is
' greater than or equal to age filter specified.
' Strings comparision ignore case by converting to lower
' Date,Age comparison is greater then or equal to.
''
Public Function IsPersonMatch(ByVal pMatch As Person) As Boolean
    Select Case True
        Case this.mFirstName <> VBA.vbNullString Or this.mLastName <> VBA.vbNullString
            IsPersonMatch = InStr(LCase$(pMatch.FirstName), this.mFirstName) And InStr(LCase$(pMatch.LastName), this.mLastName)
            If Not IsPersonMatch Then
                Exit Function
            End If
            
            Select Case this.mStatus
                Case PersonStatus.Alive
                    IsPersonMatch = pMatch.IsAlive
                Case PersonStatus.Deceased
                    IsPersonMatch = Not pMatch.IsAlive
            End Select
            If Not IsPersonMatch Then
                Exit Function
            End If
            
            If Not this.mDateOfBirth Is Nothing Then
                IsPersonMatch = pMatch.DateOfBirth.Ticks >= this.mDateOfBirth.Ticks
            ElseIf this.mAge > 0 Then
                IsPersonMatch = DateTime.TotalYears(pMatch.DateOfBirth, this.mDateTimeNow) >= this.mAge
            End If
        Case Not this.mDateOfBirth Is Nothing
            IsPersonMatch = pMatch.DateOfBirth.Ticks >= this.mDateOfBirth.Ticks
        Case this.mAge > 0
            IsPersonMatch = DateTime.TotalYears(pMatch.DateOfBirth, this.mDateTimeNow) >= this.mAge And (pMatch.IsAlive)
'        Case this.mAge > 0
'            IsPersonMatch = (DateTime.TotalYears(pMatch.DateOfBirth, this.mDateTimeNow) >= this.mAge) And (pMatch.IsAlive)
        
        Case this.mStatus = PersonStatus.Alive
            IsPersonMatch = pMatch.IsAlive
        Case this.mStatus = PersonStatus.Deceased
            IsPersonMatch = Not pMatch.IsAlive
        Case Else
            IsPersonMatch = True 'i.e. no filter applied
    End Select
End Function

'----------------------------------------------------------------------
' Interface IPredicate
'----------------------------------------------------------------------
Private Function IPredicate_CallBack(ByVal pMatch As Variant) As Boolean
    IPredicate_CallBack = IsPersonMatch(pMatch)
End Function
