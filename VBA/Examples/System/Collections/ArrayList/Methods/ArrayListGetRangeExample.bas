Attribute VB_Name = "ArrayListGetRangeExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.getrange?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to set and get a range of elements in
' the ArrayList.
''
Public Sub ArrayListGetRange()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    Call myAL.Add("The")
    Call myAL.Add("quick")
    Call myAL.Add("brown")
    Call myAL.Add("fox")
    Call myAL.Add("jumps")
    Call myAL.Add("over")
    Call myAL.Add("the")
    Call myAL.Add("lazy")
    Call myAL.Add("dog")
    
    ' Creates and initializes the source ICollection.
    Dim mySourceList As DotNetLib.Queue
    Set mySourceList = Queue.Create()
    Call mySourceList.Enqueue("big")
    Call mySourceList.Enqueue("gray")
    Call mySourceList.Enqueue("wolf")

    ' Displays the values of five elements starting at index 0.
    Dim mySubAL As DotNetLib.ArrayList
    Set mySubAL = myAL.GetRange(0, 5)
    Debug.Print "Index 0 through 4 contains:"
    Dim mySeparator As String
    mySeparator = Regex.Unescape("\t")
    Call PrintValues(mySubAL, mySeparator)
    
    ' Replaces the values of five elements starting at index 1 with the values in the ICollection.
    Call myAL.SetRange(1, mySourceList)
    
    ' Displays the values of five elements starting at index 0.
    Set mySubAL = myAL.GetRange(0, 5)
    Debug.Print "Index 0 through 4 contains:"
    Call PrintValues(mySubAL, mySeparator)
End Sub

Public Sub PrintValues(ByVal myList As mscorlib.IEnumerable, ByVal mySeparator As String)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBAString.Format("{0}{1}", mySeparator, obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'Index 0 through 4 contains:
'        The     quick   brown   fox     jumps
'Index 0 through 4 now contains:
'        The     big     gray    wolf    jumps
'*/
