Attribute VB_Name = "SortedListGetValueListExample"
'@Folder("Examples.System.Collections.SortedList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 18, 2023
'@LastModified October 18, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.getvaluelist?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to get one or all the keys or values in
' a SortedList object.
''
Public Sub SortedListGetValueList()
    ' Creates and initializes a new SortedList.
    Dim mySL As DotNetLib.SortedList
    Set mySL = SortedList.Create()
    Call mySL.Add(1.3, "fox")
    Call mySL.Add(1.4, "jumps")
    Call mySL.Add(1.5, "over")
    Call mySL.Add(1.2, "brown")
    Call mySL.Add(1.1, "quick")
    Call mySL.Add(1#, "The")
    Call mySL.Add(1.6, "the")
    Call mySL.Add(1.8, "dog")
    Call mySL.Add(1.7, "lazy")
    
    ' Gets the key and the value based on the index.
    Dim myIndex As Long
    myIndex = 3
    Debug.Print VBAString.Format("The key   at index {0} is {1}.", myIndex, mySL.GetKey(myIndex))
    Debug.Print VBAString.Format("The value at index {0} is {1}.", myIndex, mySL.GetByIndex(myIndex))
    
    ' Gets the list of keys and the list of values.
    Dim myKeyList As mscorlib.IList
    Set myKeyList = mySL.GetKeyList()
    Dim myValueList As mscorlib.IList
    Set myValueList = mySL.GetValueList()
    
    ' Prints the keys in the first column and the values in the second column.
    Debug.Print Regex.Unescape("\t-KEY-\t-VALUE-")
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}\t{1}")
    Dim i As Long
    For i = 0 To mySL.count - 1
        Debug.Print VBAString.Format(formatString, myKeyList(i), myValueList(i))
    Next i
End Sub

'/*
'This code produces the following output.
'
'The key   at index 3 is 1.3.
'The value at index 3 is fox.
'    -KEY-    -VALUE-
'    1      The
'    1.1    quick
'    1.2    brown
'    1.3    fox
'    1.4    jumps
'    1.5    over
'    1.6    the
'    1.7    lazy
'    1.8    dog
'*/
