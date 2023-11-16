Attribute VB_Name = "ArrayListReverseExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.reverse?view=netframework-4.8.1#system-collections-arraylist-reverse

Option Explicit

''
' The following code example shows how to reverse the sort order of the values
' in an ArrayList.
''
Public Sub ArrayListReverse()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    myAL.Add "The"
    myAL.Add "quick"
    myAL.Add "brown"
    myAL.Add "fox"
    myAL.Add "jumps"
    myAL.Add "over"
    myAL.Add "the"
    myAL.Add "lazy"
    myAL.Add "dog"
    
    ' Displays the values of the ArrayList.
    Debug.Print "The ArrayList initially contains the following values:"
    Call PrintValues(myAL)

    ' Reverses the sort order of the values of the ArrayList.
    Call myAL.Reverse
    
    ' Displays the values of the ArrayList.
    Debug.Print "After reversing:"
    Call PrintValues(myAL)
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBAString.Format("   {0}", obj)
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The ArrayList initially contains the following values:
'   The
'   quick
'   brown
'   fox
'   jumps
'   over
'   The
'   lazy
'   dog
'
'After reversing:
'   dog
'   lazy
'   The
'   over
'   jumps
'   fox
'   brown
'   quick
'   The
'*/
