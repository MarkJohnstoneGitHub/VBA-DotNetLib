Attribute VB_Name = "QueueDequeueExample"
'@Folder("Examples.System.Collections.Queue.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue.dequeue?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to add elements to the Queue, remove elements
' from the Queue, or view the element at the beginning of the Queue.
''
Public Sub QueueDequeue()
    ' Creates and initializes a new Queue.
    Dim myQ As DotNetLib.Queue
    Set myQ = Queue.Create()
    myQ.Enqueue "The"
    myQ.Enqueue "quick"
    myQ.Enqueue "brown"
    myQ.Enqueue "fox"
    
    ' Displays the Queue.
    Debug.Print "Queue values:";
    PrintValues myQ
    
    ' Removes an element from the Queue.
    Debug.Print Strings.Format(Regex.Unescape("(Dequeue)\t{0}"), myQ.Dequeue())
    
    ' Displays the Queue.
    Debug.Print "Queue values:";
    PrintValues myQ
    
    ' Removes another element from the Queue.
    Debug.Print Strings.Format(Regex.Unescape("(Dequeue)\t{0}"), myQ.Dequeue())
    
    ' Displays the Queue.
    Debug.Print "Queue values:";
    PrintValues myQ
    
    ' Views the first element in the Queue but does not remove it.
    Debug.Print Strings.Format(Regex.Unescape("(Peek)   \t{0}"), myQ.Peek())

    ' Displays the Queue.
    Debug.Print "Queue values:";
    PrintValues myQ
End Sub

Private Sub PrintValues(ByVal myQ As DotNetLib.Queue)
    Dim myObj As Variant
    For Each myObj In myQ
        Debug.Print Strings.Format("    {0}", myObj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'Queue values:    The    quick    brown    fox
'(Dequeue)       The
'Queue values:    quick    brown    fox
'(Dequeue)       quick
'Queue Values:    brown fox
'(Peek)          brown
'Queue Values:    brown fox
'
'*/
