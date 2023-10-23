Attribute VB_Name = "ArrayListBinarySearch2Example"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.binarysearch?view=netframework-4.8.1#system-collections-arraylist-binarysearch(system-object-system-collections-icomparer)

'@Dependencies SimpleStringComparer.cls

Option Explicit

''
' The following example creates an ArrayList of colored animals. The provided
' IComparer performs the string comparison for the binary search. The results
' of both an iterative search and a binary search are displayed.
''
Public Sub ArrayListBinarySearch2()
    Dim coloredAnimals As DotNetLib.ArrayList
    Set coloredAnimals = ArrayList.Create()
    coloredAnimals.Add ("White Tiger")
    coloredAnimals.Add ("Pink Bunny")
    coloredAnimals.Add ("Red Dragon")
    coloredAnimals.Add ("Green Frog")
    coloredAnimals.Add ("Blue Whale")
    coloredAnimals.Add ("Black Cat")
    coloredAnimals.Add ("Yellow Lion")

    ' BinarySearch requires a sorted ArrayList.
    coloredAnimals.Sort
    
    ' Compare results of an iterative search with a binary search
    Dim index As Long
    index = IterativeSearch(coloredAnimals, "White Tiger")
    Debug.Print Strings.Format("Iterative search, item found at index: {0}", index)
    
    index = coloredAnimals.BinarySearch2("White Tiger", New SimpleStringComparer)
    Debug.Print Strings.Format("Binary search, item found at index:    {0}", index)

End Sub

Private Function IterativeSearch(ByVal pList As DotNetLib.ArrayList, ByVal finditem As String) As Long
    Dim index As Long
    index = -1
    Dim i As Long
    For i = 0 To pList.Count - 1
        If finditem = pList(i) Then
            index = i
            Exit For
        End If
    Next i
    IterativeSearch = index
End Function

' This code produces the following output.
'
' Iterative search, item found at index: 5
' Binary search, item found at index:    5
'
