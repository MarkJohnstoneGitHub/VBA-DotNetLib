Attribute VB_Name = "SortedListClearExample"
'@Folder("Examples.System.Collections.SortedList.Methods")
'
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 17, 2023
'@LastModified October 17, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.clear?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to trim the unused portions of a
' SortedList object and how to clear the values of the SortedList.
''
Public Sub SortedListClear()
    Dim mySL As DotNetLib.SortedList
    Set mySL = SortedList.Create()
    Call mySL.Add("one", "The")
    Call mySL.Add("two", "quick")
    Call mySL.Add("three", "brown")
    Call mySL.Add("four", "fox")
    Call mySL.Add("five", "jumps")
    
    ' Displays the count, capacity and values of the SortedList.
    Debug.Print "Initially,"
    Debug.Print VBAString.Format("   Count    : {0}", mySL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", mySL.Capacity)
    Debug.Print "   Values:"
    Call PrintKeysAndValues(mySL)
    
    '  Trims the SortedList.
    Call mySL.TrimToSize
    
    ' Displays the count, capacity and values of the SortedList.
    Debug.Print "After the second TrimToSize,"
    Debug.Print VBAString.Format("   Count    : {0}", mySL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", mySL.Capacity)
    Debug.Print "   Values:"
    Call PrintKeysAndValues(mySL)
    
    ' Clears the SortedList.
    Call mySL.Clear
    
    ' Displays the count, capacity and values of the SortedList.
    Debug.Print "After Clear,"
    Debug.Print VBAString.Format("   Count    : {0}", mySL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", mySL.Capacity)
    Debug.Print "   Values:"
    Call PrintKeysAndValues(mySL)
        
    ' Trims the SortedList again.
    Call mySL.TrimToSize
    
    ' Displays the count, capacity and values of the SortedList.
    Debug.Print "After the second TrimToSize,"
    Debug.Print VBAString.Format("   Count    : {0}", mySL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", mySL.Capacity)
    Debug.Print "   Values:"
    Call PrintKeysAndValues(mySL)
End Sub

Private Sub PrintKeysAndValues(ByVal myList As DotNetLib.SortedList)
    Debug.Print Regex.Unescape("\t-KEY-\t-VALUE-")
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}:\t{1}")
    Dim i As Long
    For i = 0 To myList.count - 1
        Debug.Print VBAString.Format(formatString, myList.GetKey(i), myList.GetByIndex(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'Initially,
'   Count    : 5
'   Capacity : 16
'   Values:
'    -KEY-    -VALUE-
'    five:    jumps
'    four:    fox
'    one:     the
'    three:   brown
'    two:     quick
'
'After TrimToSize,
'   Count    : 5
'   Capacity : 5
'Values:
'    -KEY-    -VALUE-
'    five:    jumps
'    four:    fox
'    one:     the
'    three:   brown
'    two:     quick
'
'After Clear,
'   Count    : 0
'   Capacity : 5
'   Values:
'    -KEY-    -VALUE-
'
'After the second TrimToSize,
'   Count    : 0
'   Capacity : 0
'   Values:
'    -KEY-    -VALUE-
'*/
