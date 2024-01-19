Attribute VB_Name = "SortedListAddExample"
'@Folder "Examples.System.Collections.SortedList.Methods"
'
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 17, 2023
'@LastModified October 17, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.add?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to add elements to a SortedList object.
''
Public Sub SortedListAdd()
    ' Creates and initializes a new SortedList.
    Dim mySL As DotNetLib.SortedList
    Set mySL = SortedList.Create()
    Call mySL.Add("one", "The")
    Call mySL.Add("two", "quick")
    Call mySL.Add("three", "brown")
    Call mySL.Add("four", "fox")
    
    ' Displays the SortedList.
    Debug.Print "The SortedList contains the following:"
    Call PrintKeysAndValues(mySL)
End Sub

Private Sub PrintKeysAndValues(ByVal myList As DotNetLib.SortedList)
    Debug.Print Regex.Unescape("\t-KEY-\t-VALUE-")
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}:\t{1}")
    Dim i As Long
    For i = 0 To myList.Count - 1
        Debug.Print VBString.Format(formatString, myList.GetKey(i), myList.GetByIndex(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The SortedList contains the following:
'    -KEY-    -VALUE-
'    four:    fox
'    one:     The
'    three:   brown
'    two:     quick
'*/
