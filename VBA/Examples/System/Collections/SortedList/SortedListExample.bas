Attribute VB_Name = "SortedListExample"
'@Folder("Examples.System.Collections.SortedList")
'
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 16, 2023
'@LastModified October 16, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist?view=netframework-4.8.1

Option Explicit

''
' The following code example shows how to create and initialize a SortedList
' object and how to print out its keys and values.
''
Public Sub SortedListExample()
    ' Creates and initializes a new SortedList.
    Dim mySL As DotNetLib.SortedList
    Set mySL = SortedList.Create()
    Call mySL.Add("Third", "!")
    Call mySL.Add("Second", "World")
    Call mySL.Add("First", "Hello")
    
    ' Displays the properties and values of the SortedList.
    Debug.Print "mySL"
    Debug.Print Strings.Format("  Count:    {0}", mySL.Count)
    Debug.Print Strings.Format("  Capacity: {0}", mySL.Capacity)
    Debug.Print "  Keys and Values:"
    Call PrintKeysAndValues(mySL)
End Sub

Private Sub PrintKeysAndValues(ByVal myList As DotNetLib.SortedList)
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}:\t{1}")

    Debug.Print Regex.Unescape("\t-KEY-\t-VALUE-")
    Dim i As Long
    For i = 0 To myList.Count - 1
         Debug.Print Strings.Format(formatString, myList.GetKey(i), myList.GetByIndex(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'mySL
'  Count:    3
'  Capacity: 16
'  Keys and Values:
'    -KEY-    -VALUE-
'    First:        Hello
'    Second:        World
'    Third:    !
'*/
