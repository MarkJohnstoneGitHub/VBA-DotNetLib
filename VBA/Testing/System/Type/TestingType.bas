Attribute VB_Name = "TestingType"
'@Folder("Testing.System.Type")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 8, 2023
'@LastModified October 11, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1

Option Explicit

''
' Testing obtaining Type
''
Public Sub TestType()
    Dim myDateTime As DotNetLib.DateTime
    Set myDateTime = DateTime.CreateFromDate(2023, 1, 1)
    
    Dim DateTimeType As DotNetLib.Type
    Set DateTimeType = myDateTime.GetType()
    Debug.Print DateTimeType.FullName
    
    Dim pvtMembers() As mscorlib.MemberInfo
    pvtMembers = DateTimeType.GetMembers()
    Dim pvtMemberInfo As mscorlib.MemberInfo
    
    Dim i As Long
    For i = 0 To UBound(pvtMembers)
        Set pvtMemberInfo = pvtMembers(i)
        Debug.Print pvtMemberInfo.ToString
    Next i
    Debug.Print
    Dim myPerson As Person
    Set myPerson = Person.Create("Mary", "Smith", DateTime.CreateFromDate(2000, 1, 1))
    
    Dim personType As DotNetLib.Type
    Set personType = myPerson.GetType()
    Debug.Print personType.FullName
    Debug.Print personType.IsCOMObject
    pvtMembers = personType.GetMembers()
    
    For i = 0 To UBound(pvtMembers)
        Set pvtMemberInfo = pvtMembers(i)
        Debug.Print pvtMemberInfo.ToString
    Next i
    
End Sub
