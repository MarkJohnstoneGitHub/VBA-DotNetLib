Attribute VB_Name = "ListObjectExamples"
'@Folder("Examples.System.Collections.Generic.ListObject")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 5, 2023
'@LastModified October 16, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1.contains?view=netframework-4.8.1

Option Explicit

'Test example using ListString i.e. List<String>
Public Sub ListObjectExample1()
    Dim names As DotNetLib.ListObject
    Set names = ListObject.Create()
    names.Add "Bob"
    names.Add "Mary"
    names.Add "James"
    names.Add "Mark"
    names.Add "Brian"
    names.Add "susan"
    names.Add "Michael"
    names.Add DateTime.Now
    names.Add Person.Create("Martha", "Jones", DateTime.CreateFromDate(2000, 1, 1))
    Debug.Print Strings.Format("Initial list: names.Count {0}", names.Count)
    DisplayList names
    Debug.Print
    names.Reverse
    DisplayList names
End Sub

Private Sub DisplayList(ByVal pList As DotNetLib.ListObject)
    Dim pvtIndex As Long
    pvtIndex = 0
    Dim varObject As Variant
    For Each varObject In pList
        If TypeOf varObject Is IStringable Then
            Dim stringableObj As IStringable
            Set stringableObj = varObject
            varObject = stringableObj.ToString
        End If
        Debug.Print Strings.Format("[{0}] {1}", pvtIndex, varObject)
        pvtIndex = pvtIndex + 1
    Next
End Sub
