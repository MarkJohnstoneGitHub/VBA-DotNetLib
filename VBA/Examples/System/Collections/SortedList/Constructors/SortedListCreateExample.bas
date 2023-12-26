Attribute VB_Name = "SortedListCreateExample"
'@Folder("Examples.System.Collections.SortedList.Constructors")
'
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 16, 2023
'@LastModified October 17, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.-ctor?view=netframework-4.8.1#system-collections-sortedlist-ctor

Option Explicit

''
' The following code example creates collections using different SortedList
' constructors and demonstrates the differences in the behavior of the collections.
''
Public Sub SortedListCreate()
    ' Create a SortedList using the default comparer.
    Dim mySL1 As DotNetLib.SortedList
    Set mySL1 = SortedList.Create()
    Debug.Print "mySL1 (default):"
    Call mySL1.Add("FIRST", "Hello")
    Call mySL1.Add("SECOND", "World")
    Call mySL1.Add("THIRD", "!")
    
    On Error Resume Next
    Call mySL1.Add("first", "Ola!")
    If Catch(ArgumentException) Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    Call PrintKeysAndValues(mySL1)
    
    ' Create a SortedList using the specified case-insensitive comparer.
    Dim mySL2 As DotNetLib.SortedList
    Set mySL2 = SortedList.Create2(CaseInsensitiveComparer.Create())
    Debug.Print "mySL2 (case-insensitive comparer):"
    Call mySL2.Add("FIRST", "Hello")
    Call mySL2.Add("SECOND", "World")
    Call mySL2.Add("THIRD", "!")
    On Error Resume Next
    Call mySL2.Add("first", "Ola!")
    If Catch(ArgumentException) Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    Call PrintKeysAndValues(mySL2)
    
    ' Create a SortedList using the specified CaseInsensitiveComparer,
    ' which is based on the Turkish culture (tr-TR), where "I" is not
    ' the uppercase version of "i".
    Dim myCul As DotNetLib.CultureInfo
    Set myCul = CultureInfo.CreateFromName("tr-TR")
    Dim mySL3 As DotNetLib.SortedList
    Set mySL3 = SortedList.Create2(CaseInsensitiveComparer.Create2(myCul))
    Debug.Print "mySL3 (case-insensitive comparer, Turkish culture):"
    Call mySL3.Add("FIRST", "Hello")
    Call mySL3.Add("SECOND", "World")
    Call mySL3.Add("THIRD", "!")
    On Error Resume Next
    Call mySL3.Add("first", "Ola!")
    If Catch(ArgumentException) Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    Call PrintKeysAndValues(mySL3)
    
    ' Create a SortedList using the
    ' StringComparer.InvariantCultureIgnoreCase value.
    Dim mySL4 As DotNetLib.SortedList
    Set mySL4 = SortedList.Create2(StringComparer.InvariantCultureIgnoreCase)
    Debug.Print "mySL4 (InvariantCultureIgnoreCase):"
    Call mySL4.Add("FIRST", "Hello")
    Call mySL4.Add("SECOND", "World")
    Call mySL4.Add("THIRD", "!")
    On Error Resume Next
    Call mySL4.Add("first", "Ola!")
    If Catch(ArgumentException) Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    Call PrintKeysAndValues(mySL4)
End Sub

Private Sub PrintKeysAndValues(ByVal myList As DotNetLib.SortedList)
    Debug.Print "        -KEY-   -VALUE-"
    Dim i As Long
    For i = 0 To myList.Count - 1
        Debug.Print VBAString.Format("        {0,-6}: {1}", myList.GetKey(i), myList.GetByIndex(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'Results vary depending on the system's culture settings.
'
'mySL1 (Default):
'        -KEY-   -VALUE-
'        FIRST:  Ola!
'        FIRST:  Hello
'        SECOND: World
'        THIRD : !
'
'mySL2 (case-insensitive comparer):
'System.ArgumentException: Item has already been added.  Key in dictionary: 'FIRST'  Key being added: 'first'
'   at System.Collections.SortedList.Add(Object key, Object value)
'   at SamplesSortedList.Main()
'        -KEY-   -VALUE-
'        FIRST:  Hello
'        SECOND: World
'        THIRD : !
'
'mySL3 (case-insensitive comparer, Turkish culture):
'        -KEY-   -VALUE-
'        FIRST:  Hello
'        FIRST:  Ola!
'        SECOND: World
'        THIRD : !
'
'mySL4 (InvariantCultureIgnoreCase):
'System.ArgumentException: Item has already been added.  Key in dictionary: 'FIRST'  Key being added: 'first'
'   at System.Collections.SortedList.Add(Object key, Object value)
'   at SamplesSortedList.Main()
'        -KEY-   -VALUE-
'        FIRST:  Hello
'        SECOND: World
'        THIRD : !
'
'*/
