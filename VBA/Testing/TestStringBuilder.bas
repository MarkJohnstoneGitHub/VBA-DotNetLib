Attribute VB_Name = "TestStringBuilder"
'@Folder("Testing")

Option Explicit


'@Reference https://stackoverflow.com/questions/76347710/excel-vba-getting-automation-error-430-using-stringbuilder-class-from-net-3
Public Function StringBuilderTest() As String
    Dim sb As mscorlib.StringBuilder
    Dim i As Integer
    
    On Error GoTo ErrorLabel
    
    Set sb = New mscorlib.StringBuilder
    
    For i = 1 To 100
        sb.Append_3 "text" 'Fixed changed from sb.Append "text" which was causing Automation Error 430
    Next i
    
    StringBuilderTest = sb.ToString
    Exit Function
    
ErrorLabel:
    Debug.Print Err.Description
    Debug.Print Err.number
    Debug.Print Err.Source
    
    StringBuilderTest = "Error"
    Exit Function
    
End Function

Public Sub DisplayStringBuilderMembers()
    Dim sb As mscorlib.StringBuilder
    Set sb = New mscorlib.StringBuilder
    Dim stringBuilderType As mscorlib.Type
    Set stringBuilderType = sb.GetType()
    Debug.Print stringBuilderType.FullName; " members"
    Debug.Print
    Dim Members() As mscorlib.MemberInfo
    Members = stringBuilderType.GetMembers_2
    Dim i As Long
    For i = 0 To UBound(Members)
        Dim pvtMemberInfo As mscorlib.MemberInfo
        Set pvtMemberInfo = Members(i)
        Debug.Print pvtMemberInfo.ToString
    Next i
End Sub

Public Sub DisplayMembers(ByVal pType As mscorlib.Type)
    Dim pvtMembers() As mscorlib.MemberInfo
    pvtMembers = pType.GetMembers_2
    Dim i As Long
    For i = 0 To UBound(pvtMembers)
        Dim pvtMemberInfo As mscorlib.MemberInfo
        Set pvtMemberInfo = pvtMembers(i)
        Debug.Print pvtMemberInfo.ToString
    Next i
End Sub

Public Sub DisplayStringMembers()
    Dim str As DotNetLib.String
    Set str = Strings.Create("abc")
    Dim obj As DotNetLib.Object
    Set obj = Objects.Create(str.WrappedString)
    Dim objType As mscorlib.Type
    Set objType = obj.GetType.WrappedType
    
    Debug.Print objType.FullName; " members"
    Debug.Print
    Dim Members() As mscorlib.MemberInfo
    Members = objType.GetMembers_2
    Dim i As Long
    For i = 0 To UBound(Members)
        Dim pvtMemberInfo As mscorlib.MemberInfo
        Set pvtMemberInfo = Members(i)
        Debug.Print pvtMemberInfo.ToString
    Next i
End Sub

