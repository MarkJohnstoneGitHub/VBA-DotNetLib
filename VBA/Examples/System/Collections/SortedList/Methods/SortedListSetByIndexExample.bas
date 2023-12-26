Attribute VB_Name = "SortedListSetByIndexExample"
'@Folder("Examples.System.Collections.SortedList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 18, 2023
'@LastModified October 18, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.setbyindex?view=netframework-4.8.1#examples

Option Explicit

''
' Replaces the value at a specific index in a SortedList object.
''
Public Sub SortedListSetByIndex()
    ' Creates and initializes a new SortedList.
    Dim mySL As DotNetLib.SortedList
    Set mySL = SortedList.Create()
    Call mySL.Add(2, "two")
    Call mySL.Add(3, "three")
    Call mySL.Add(1, "one")
    Call mySL.Add(0, "zero")
    Call mySL.Add(4, "four")
    
    ' Displays the values of the SortedList.
    Debug.Print "The SortedList contains the following values:"
    Call PrintIndexAndKeysAndValues(mySL)
    
    ' Replaces the values at index 3 and index 4.
    Call mySL.SetByIndex(3, "III")
    Call mySL.SetByIndex(4, "IV")
    
    ' Displays the updated values of the SortedList.
    Debug.Print "After replacing the value at index 3 and index 4,"
    Call PrintIndexAndKeysAndValues(mySL)
End Sub

Private Sub PrintIndexAndKeysAndValues(ByVal myList As DotNetLib.SortedList)
    Debug.Print Regex.Unescape("\t-INDEX-\t-KEY-\t-VALUE-")
    Dim formatString As String
    formatString = Regex.Unescape("\t[{0}]:\t{1}\t{2}")
    Dim i As Long
    For i = 0 To myList.Count - 1
        Debug.Print VBAString.Format(formatString, i, myList.GetKey(i), myList.GetByIndex(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The SortedList contains the following values:
'    -INDEX-    -KEY-    -VALUE-
'    [0]:    0    zero
'    [1]:    1    one
'    [2]:    2    two
'    [3]:    3    three
'    [4]:    4    four
'
'After replacing the value at index 3 and index 4,
'    -INDEX-    -KEY-    -VALUE-
'    [0]:    0    zero
'    [1]:    1    one
'    [2]:    2    two
'    [3]:    3    III
'    [4]:    4    IV
'*/
