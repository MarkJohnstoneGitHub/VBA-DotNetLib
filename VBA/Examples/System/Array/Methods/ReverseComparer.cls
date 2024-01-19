VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReverseComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Examples.System.Array.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 29, 2023
'@LastModified October 29, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.sort?view=netframework-4.8.1#system-array-sort(system-array)

Option Explicit

Implements mscorlib.IComparer

Private Type TReverseComparer
   CaseInsensitiveComparer As DotNetLib.CaseInsensitiveComparer
End Type

Private this As TReverseComparer

Private Sub Class_Initialize()
   Set this.CaseInsensitiveComparer = CaseInsensitiveComparer.Create()
End Sub


Public Function Compare(ByVal x As Variant, ByVal y As Variant) As Long
    Compare = this.CaseInsensitiveComparer.Compare(y, x)
End Function

' ------------------------------------------------------------'
'mscorlib.IComparer Methods
' ------------------------------------------------------------'
Private Function IComparer_Compare(ByVal x As Variant, ByVal y As Variant) As Long
    ' Call CaseInsensitiveComparer.Compare with the parameters reversed.
    IComparer_Compare = this.CaseInsensitiveComparer.Compare(y, x)
End Function
