Attribute VB_Name = "QueueClearExample"
'@Folder("Examples.System.Collections.Queue.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue.clear?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to clear the values of the Queue.
''
Public Sub QueueClear()
    ' Creates and initializes a new Queue.
    Dim myQ As DotNetLib.Queue
    Set myQ = Queue.Create()
    myQ.Enqueue "The"
    myQ.Enqueue "brown"
    myQ.Enqueue "fox"
    myQ.Enqueue "jumps"
    
    ' Displays the count and values of the Queue.
    Debug.Print "Initially,"
    Debug.Print Strings.Format("   Count    : {0}", myQ.Count)
    Debug.Print "   Values:";
    PrintValues myQ
    
    ' Clears the Queue.
    myQ.Clear

    ' Displays the count and values of the Queue.
    Debug.Print "After Clear,"
    Debug.Print Strings.Format("   Count    : {0}", myQ.Count)
    Debug.Print "   Values:";
    PrintValues myQ
    
End Sub

Public Sub PrintValues(ByVal myQ As DotNetLib.Queue)
    Dim myObj As Variant
    For Each myObj In myQ
        Debug.Print Strings.Format("    {0}", myObj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'Initially,
'   Count    : 5
'   Values:    The    quick    brown    fox    jumps
'After Clear,
'   Count    : 0
'Values:
'
'*/
