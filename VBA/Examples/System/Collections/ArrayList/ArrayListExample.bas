Attribute VB_Name = "ArrayListExample"
'@Folder("Examples.System.Collections.ArrayList")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 5, 2023
'@LastModified October 5, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to create and initialize an ArrayList and
' how to display its values.
''
Public Sub ArrayListExample1()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    myAL.Add "Hello"
    myAL.Add "World"
    myAL.Add "!"
    
    ' Displays the properties and values of the ArrayList.
    Debug.Print "myAL"
    Debug.Print VBAString.Format("    Count:    {0}", myAL.Count)
    Debug.Print VBAString.Format("    Capacity: {0}", myAL.Capacity)
    Debug.Print "    Values:";
    PrintValues myAL
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBAString.Format("    {0}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces output similar to the following:
'
'myAL
'    Count:    3
'    Capacity: 4
'    Values:   Hello   World   !
'
'*/
