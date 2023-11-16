Attribute VB_Name = "SortedListContainsExample"
'@Folder("Examples.System.Collections.SortedList.Methods")
'
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 17, 2023
'@LastModified October 27, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.contains?view=netframework-4.8.1#examples

'@Remarks
' For this example Key is required to be converted to Long type.

Option Explicit

''
' The following code example shows how to determine whether a SortedList object
' contains a specific element.
''
Public Sub SortedListContains()
    ' Creates and initializes a new SortedList.
    Dim mySL As DotNetLib.SortedList
    Set mySL = SortedList.Create()
    Call mySL.Add(CLng(2), "two")
    Call mySL.Add(CLng(4), "four")
    Call mySL.Add(CLng(1), "one")
    Call mySL.Add(CLng(3), "three")
    Call mySL.Add(CLng(0), "zero")

    ' Displays the values of the SortedList.
    Debug.Print "The SortedList contains the following values:"
    Call PrintIndexAndKeysAndValues(mySL)
    
    ' Searches for a specific key.
    Dim myKey As Long
    myKey = 2
    Debug.Print VBAString.Format("The key ""{0}"" is {1}.", myKey, IIf(mySL.ContainsKey(myKey), "in the SortedList", "NOT in the SortedList"))
    myKey = 6
    Debug.Print VBAString.Format("The key ""{0}"" is {1}.", myKey, IIf(mySL.ContainsKey(myKey), "in the SortedList", "NOT in the SortedList"))
    
    ' Searches for a specific value.
    Dim myValue As String
    myValue = "three"
    Debug.Print VBAString.Format("The value ""{0}"" is {1}.", myValue, IIf(mySL.ContainsValue(myValue), "in the SortedList", "NOT in the SortedList"))
    myValue = "nine"
    Debug.Print VBAString.Format("The value ""{0}"" is {1}.", myValue, IIf(mySL.ContainsValue(myValue), "in the SortedList", "NOT in the SortedList"))
End Sub

Private Sub PrintIndexAndKeysAndValues(ByVal myList As DotNetLib.SortedList)
    Debug.Print Regex.Unescape("\t-INDEX-\t-KEY-\t-VALUE-")
    Dim formatString As String
    formatString = Regex.Unescape("\t[{0}]:\t{1}\t{2}")
    Dim i As Long
    For i = 0 To myList.count - 1
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
'The key "2" is in the SortedList.
'The key "6" is NOT in the SortedList.
'The value "three" is in the SortedList.
'The value "nine" is NOT in the SortedList.
'*/
