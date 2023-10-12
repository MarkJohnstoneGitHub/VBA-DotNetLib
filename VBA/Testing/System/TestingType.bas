Attribute VB_Name = "TestingType"
'@Folder("Testing.System")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 8, 2023
'@LastModified October 8, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1

Option Explicit

'
Public Sub TestType()
    Dim myDateTime As DotNetLib.DateTime
    Set myDateTime = DateTime.CreateFromDate(2023, 1, 1)
    
    Dim dateTimeType As DotNetLib.Type
    Set dateTimeType = myDateTime.GetType()
    Debug.Print dateTimeType.FullName
    
    Dim myPerson As Person
    Set myPerson = Person.Create("Mary", DateTime.CreateFromDate(2000, 1, 1))
    
    Dim personType As DotNetLib.Type
    Set personType = myPerson.GetType()
    Debug.Print personType.FullName
    
End Sub
