Attribute VB_Name = "ArrayListTrimToSizeExample"
'@Folder "Examples.System.Collections.ArrayList.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 26, 2023
'@LastModified October 26, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.trimtosize?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to trim the unused portions of the
' ArrayList and how to clear the values of the ArrayList.
''
Public Sub ArrayListTrimToSize()
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
    Debug.Print VBString.Format("   Count    : {0}", myAL.Count)
    Debug.Print VBString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)

    ' Trim the ArrayList.
    Call myAL.TrimToSize
    
    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "After TrimToSize,"
    Debug.Print VBString.Format("   Count    : {0}", myAL.Count)
    Debug.Print VBString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)

    ' Clear the ArrayList.
    Call myAL.Clear

    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "After Clear,"
    Debug.Print VBString.Format("   Count    : {0}", myAL.Count)
    Debug.Print VBString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
    
    ' Trim the ArrayList again.
    Call myAL.TrimToSize

    ' Displays the count, capacity and values of the ArrayList.
    Debug.Print "After the second TrimToSize,"
    Debug.Print VBString.Format("   Count    : {0}", myAL.Count)
    Debug.Print VBString.Format("   Capacity : {0}", myAL.Capacity)
    Debug.Print "   Values:";
    Call PrintValues(myAL)
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBString.Format("   {0}", obj);
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
