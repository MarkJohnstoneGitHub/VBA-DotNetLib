Attribute VB_Name = "ArrayListIsReadOnlyExample"
'@Folder("Examples.System.Collections.ArrayList.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 6, 2023

'Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.isreadonly?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to create a read-only wrapper around an
' ArrayList and how to determine if an ArrayList is read-only.
''
Public Sub ArrayListIsReadOnly()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    myAL.Add "red"
    myAL.Add "orange"
    myAL.Add "yellow"

    ' Creates a read-only copy of the ArrayList.
    Dim myReadOnlyAL As DotNetLib.ArrayList
    Set myReadOnlyAL = ArrayList.ReadOnly(myAL)
    
    '  Displays whether the ArrayList is read-only or writable.
    Debug.Print Strings.Format("myAL is {0}.", IIf(myAL.IsReadOnly, "read-only", "writable"))
    Debug.Print Strings.Format("myReadOnlyAL is {0}.", IIf(myReadOnlyAL.IsReadOnly, "read-only", "writable"))
    
    ' Displays the contents of both collections.
    Debug.Print Regex.Unescape("\nInitially,")
    Debug.Print "The original ArrayList myAL contains:"
    Dim myStr As Variant
    For Each myStr In myAL
        Debug.Print Strings.Format("   {0}", myStr)
    Next

    Debug.Print "The read-only ArrayList myReadOnlyAL contains:"
    For Each myStr In myReadOnlyAL
        Debug.Print Strings.Format("   {0}", myStr)
    Next
    
    ' Adding an element to a read-only ArrayList throws an exception.
    Debug.Print Regex.Unescape("\nTrying to add a new element to the read-only ArrayList:")
    On Error Resume Next
    myReadOnlyAL.Add "green"
    If Catch() Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
    
    ' Adding an element to the original ArrayList affects the read-only ArrayList.
    myAL.Add "blue"
    
    ' Displays the contents of both collections again.
    Debug.Print Regex.Unescape("\nAfter adding a new element to the original ArrayList,")
    Debug.Print "The original ArrayList myAL contains:"
    For Each myStr In myAL
        Debug.Print Strings.Format("   {0}", myStr)
    Next
    Debug.Print "The read-only ArrayList myReadOnlyAL contains:"
    For Each myStr In myReadOnlyAL
        Debug.Print Strings.Format("   {0}", myStr)
    Next
End Sub

'/*
'This code produces the following output.
'
'myAL is writable.
'myReadOnlyAL is read-only.
'
'Initially,
'The original ArrayList myAL contains:
'   red
'   orange
'   yellow
'The read-only ArrayList myReadOnlyAL contains:
'   red
'   orange
'   yellow
'
'Trying to add a new element to the read-only ArrayList:
'Exception: System.NotSupportedException: Collection is read-only.
'   at System.Collections.ReadOnlyArrayList.Add(Object obj)
'   at SamplesArrayList.Main()
'
'After adding a new element to the original ArrayList,
'The original ArrayList myAL contains:
'   red
'   orange
'   yellow
'   blue
'The read-only ArrayList myReadOnlyAL contains:
'   red
'   orange
'   yellow
'   blue
'
'*/
