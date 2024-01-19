Attribute VB_Name = "ArrayListInsertExample"
'@Folder "Examples.System.Collections.ArrayList.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.insert?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to insert elements into the ArrayList.
''
Public Sub ArrayListInsert()
    ' Creates and initializes a new ArrayList using Insert instead of Add.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    Call myAL.Insert(0, "The")
    Call myAL.Insert(1, "fox")
    Call myAL.Insert(2, "jumps")
    Call myAL.Insert(3, "over")
    Call myAL.Insert(4, "the")
    Call myAL.Insert(5, "dog")
    
    ' Creates and initializes a new Queue.
    Dim myQueue As DotNetLib.Queue
    Set myQueue = Queue.Create()
    Call myQueue.Enqueue("quick")
    Call myQueue.Enqueue("brown")
    
    ' Displays the ArrayList and the Queue.
    Debug.Print "The ArrayList initially contains the following:"
    Call PrintValues(myAL)
    Debug.Print "The Queue initially contains the following:"
    Call PrintValues(myQueue)
    
    ' Copies the Queue elements to the ArrayList at index 1.
    Call myAL.InsertRange(1, myQueue)
    
    ' Displays the ArrayList.
    Debug.Print "After adding the Queue, the ArrayList now contains:"
    Call PrintValues(myAL)
    
    ' Search for "dog" and add "lazy" before it.
    Call myAL.Insert(myAL.IndexOf("dog"), "lazy")

    ' Displays the ArrayList.
    Debug.Print "After adding ""lazy"", the ArrayList now contains:"
    Call PrintValues(myAL)
    
    ' Add "!!!" at the end.
    Call myAL.Insert(myAL.Count, "!!!")

    ' Displays the ArrayList.
    Debug.Print "After adding ""!!!"", the ArrayList now contains:"
    Call PrintValues(myAL)
    
    ' Inserting an element beyond Count throws an exception.
    On Error Resume Next
    Call myAL.Insert(myAL.Count + 1, "anystring")
    If Err.Number Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

Public Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBString.Format("   {0}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The ArrayList initially contains the following:
'   The   fox   jumps   over   the   dog
'The Queue initially contains the following:
'   quick brown
'After adding the Queue, the ArrayList now contains:
'   The   quick   brown   fox   jumps   over   the   dog
'After adding "lazy", the ArrayList now contains:
'   The   quick   brown   fox   jumps   over   the   lazy   dog
'After adding "!!!", the ArrayList now contains:
'   The   quick   brown   fox   jumps   over   the   lazy   dog   !!!
'Exception: System.ArgumentOutOfRangeException: Insertion index was out of range.  Must be non-negative and less than or equal to size.
'Parameter Name: index
'   at System.Collections.ArrayList.Insert(int index, Object value)
'   at SamplesArrayList.Main()
'*/
