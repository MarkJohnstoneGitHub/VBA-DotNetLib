VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyReverserClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Examples.System.Collections.ArrayList.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 6, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.sort?view=netframework-4.8.1#system-collections-arraylist-sort(system-collections-icomparer)

Option Explicit

Implements mscorlib.IComparer

Public Function Compare(ByVal x As String, ByVal y As String) As Long
    Compare = CaseInsensitiveComparer.Create().Compare(y, x)
End Function

' ------------------------------------------------------------'
'IScriptingDictionary Methods
' ------------------------------------------------------------'
Private Function IComparer_Compare(ByVal x As Variant, ByVal y As Variant) As Long
    IComparer_Compare = Compare(x, y)
End Function

