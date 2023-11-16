VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Examples.System.Collections.Hashtable.Constructors")

Option Explicit

Implements mscorlib.IEqualityComparer

'class myComparer: IEqualityComparer
'{
'    public new bool Equals(object x, object y)
'    {
'        return x.Equals(y);
'    }
'
'    public int GetHashCode(object obj)
'    {
'        return obj.ToString().ToLower().GetHashCode();
'    }
'}

Private Function IEqualityComparer_Equals(ByVal x As Variant, ByVal y As Variant) As Boolean
    IEqualityComparer_Equals = Objects.Equals(x, y)
End Function

Private Function IEqualityComparer_GetHashCode(ByVal obj As Variant) As Long
    Dim objDNL As DotNetLib.Object
    Set objDNL = Objects.Create(obj)
    IEqualityComparer_GetHashCode = objDNL.GetHashCode()
End Function
