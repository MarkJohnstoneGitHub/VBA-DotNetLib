VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SimpleStringComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Examples.System.Collections.ArrayList.Properties")

Option Explicit

Implements mscorlib.IComparer

Public Function Compare(ByVal x As String, y As String) As Long
    Compare = Strings.Compare(x, y)
End Function

'public class SimpleStringComparer : IComparer
'{
'    int IComparer.Compare(object x, object y)
'    {
'        string cmpstr = (string)x;
'        return cmpstr.CompareTo((string)y);
'    }
'}

' ------------------------------------------------------------'
'IScriptingDictionary Methods
' ------------------------------------------------------------'
Private Function IComparer_Compare(ByVal x As Variant, ByVal y As Variant) As Long
'    Dim cmpstr As String
'    cmpstr = CStr(x)
    IComparer_Compare = Strings.Compare(x, y)
End Function
