Attribute VB_Name = "SortedListRemoveExample"
'@Folder("Examples.System.Collections.SortedList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 18, 2023
'@LastModified October 18, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.remove?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to remove elements from a SortedList object.
''
Public Sub SortedListRemove()
    ' Creates and initializes a new SortedList.
    Dim mySL As DotNetLib.SortedList
    Set mySL = SortedList.Create()
    Call mySL.Add("3c", "dog")
    Call mySL.Add("2c", "over")
    Call mySL.Add("1c", "brown")
    Call mySL.Add("1a", "The")
    Call mySL.Add("1b", "quick")
    Call mySL.Add("3a", "the")
    Call mySL.Add("3b", "lazy")
    Call mySL.Add("2a", "fox")
    Call mySL.Add("2b", "jumps")
    
    ' Displays the SortedList.
    Debug.Print "The SortedList initially contains the following:"
    Call PrintKeysAndValues(mySL)
    
    ' Removes the element with the key "3b".
    Call mySL.Remove("3b")
    
    ' Displays the current state of the SortedList.
    Debug.Print "After removing ""lazy"":"
    Call PrintKeysAndValues(mySL)

    ' Removes the element at index 5.
    Call mySL.RemoveAt(5)
    
    ' Displays the current state of the SortedList.
    Debug.Print "After removing the element at index 5:"
    Call PrintKeysAndValues(mySL)

End Sub

Private Sub PrintKeysAndValues(ByVal myList As DotNetLib.SortedList)
    Debug.Print Regex.Unescape("\t-KEY-\t-VALUE-")
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}:\t{1}")
    Dim i As Long
    For i = 0 To myList.Count - 1
        Debug.Print Strings.Format(formatString, myList.GetKey(i), myList.GetByIndex(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The SortedList initially contains the following:
'    -KEY-    -VALUE-
'1 a:       the
'1 b:       quick
'1 c:       brown
'2 a:       fox
'2 b:       jumps
'2 c:       over
'3 a:       the
'3 b:       lazy
'3 c:       dog
'
'After removing "lazy":
'    -KEY-    -VALUE-
'1 a:       the
'1 b:       quick
'1 c:       brown
'2 a:       fox
'2 b:       jumps
'2 c:       over
'3 a:       the
'3 c:       dog
'
'After removing the element at index 5:
'    -KEY-    -VALUE-
'1 a:       the
'1 b:       quick
'1 c:       brown
'2 a:       fox
'2 b:       jumps
'3 a:       the
'3 c:       dog
'*/
