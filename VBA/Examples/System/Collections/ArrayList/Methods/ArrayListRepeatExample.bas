Attribute VB_Name = "ArrayListRepeatExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.repeat?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to create and initialize a new ArrayList
' with the same value.
''
Public Sub ArrayListRepeat()
    ' Creates a new ArrayList with five elements and initialize each element with a null value.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Repeat(Null, 5)

    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "ArrayList with five elements with a null value"
    Debug.Print VBAString.Format("   Count    : {0}", myAL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
    
    ' Creates a new ArrayList with seven elements and initialize each element with the string "abc".
    Set myAL = ArrayList.Repeat("abc", 7)

    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "ArrayList with seven elements with a string value"
    Debug.Print VBAString.Format("   Count    : {0}", myAL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBAString.Format("   {0}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'ArrayList with five elements with a null value
'   Count    : 5
'   Capacity : 5
'Values:
'ArrayList with seven elements with a string value
'   Count    : 7
'   Capacity : 7
'   Values:   abc   abc   abc   abc   abc   abc   abc
'
'*/
