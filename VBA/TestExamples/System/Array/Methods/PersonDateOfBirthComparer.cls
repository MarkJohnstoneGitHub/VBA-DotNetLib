VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PersonDateOfBirthComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("TestExamples.System.Array.Methods")
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 13, 2023
'@LastModified February 13, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.sort?view=netframework-4.8.1#system-array-sort(system-array)

Option Explicit

Implements mscorlib.IComparer

Private Type TPersonDateOfBirthComparer
   CaseInsensitiveComparer As DotNetLib.CaseInsensitiveComparer
End Type

Private this As TPersonDateOfBirthComparer

Private Sub Class_Initialize()
   Set this.CaseInsensitiveComparer = CaseInsensitiveComparer.Create()
End Sub

Public Function CompareDateOfBirth(ByVal x As Person, ByVal y As Person) As Long
    If (x = y) Then
        CompareDateOfBirth = 0
        Exit Function
    End If
    ' they can't both be null;
    If (x Is Nothing) Then
        CompareDateOfBirth = -1
        Exit Function
    End If
    If (y Is Nothing) Then
        CompareDateOfBirth = 1
        Exit Function
    End If
    
    If x.DateOfBirth.Ticks < y.DateOfBirth.Ticks Then
        CompareDateOfBirth = -1
        Exit Function
    End If
    If x.DateOfBirth.Ticks > y.DateOfBirth.Ticks Then
        CompareDateOfBirth = 1
        Exit Function
    End If
    CompareDateOfBirth = 0
End Function

' ------------------------------------------------------------'
'mscorlib.IComparer Methods
' ------------------------------------------------------------'
Private Function IComparer_Compare(ByVal x As Variant, ByVal y As Variant) As Long
    IComparer_Compare = Me.CompareDateOfBirth(x, y)
End Function
