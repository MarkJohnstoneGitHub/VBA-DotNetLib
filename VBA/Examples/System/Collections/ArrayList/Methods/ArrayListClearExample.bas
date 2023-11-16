Attribute VB_Name = "ArrayListClearExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 24, 2023
'@LastModified October 24, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.clear?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to trim the unused portions of the
' ArrayList and how to clear the values of the ArrayList.
''
Public Sub ArrayListClear()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    Call myAL.Add("The")
    Call myAL.Add("quick")
    Call myAL.Add("brown")
    Call myAL.Add("fox")
    Call myAL.Add("jumps")
    
    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "Initially,"
    Debug.Print VBAString.Format("   Count    : {0}", myAL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
    
    ' Trim the ArrayList.
    Call myAL.TrimToSize

    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "After TrimToSize,"
    Debug.Print VBAString.Format("   Count    : {0}", myAL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
    
    ' Clear the ArrayList.
    Call myAL.Clear
    
    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "After Clear,"
    Debug.Print VBAString.Format("   Count    : {0}", myAL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
    
    ' Trim the ArrayList again.
    Call myAL.TrimToSize
    
    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "After the second TrimToSize,"
    Debug.Print VBAString.Format("   Count    : {0}", myAL.count)
    Debug.Print VBAString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
End Sub

Public Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBAString.Format("   {0}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'Initially,
'   Count    : 5
'   Capacity : 8
'   Values:    The    quick    brown    fox    jumps
'After TrimToSize,
'   Count    : 5
'   Capacity : 5
'   Values:    The    quick    brown    fox    jumps
'After Clear,
'   Count    : 0
'   Capacity : 5
'Values:
'After the second TrimToSize,
'   Count    : 0
'   Capacity : 4
'Values:
'*/
