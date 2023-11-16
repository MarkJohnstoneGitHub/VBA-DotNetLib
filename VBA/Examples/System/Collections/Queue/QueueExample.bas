Attribute VB_Name = "QueueExample"
'@Folder("Examples.System.Collections.Queue")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to create and add values to a Queue and how
' to print out its values.
''
Public Sub QueueExample()
    Dim myQ As DotNetLib.Queue
    Set myQ = Queue.Create()
    myQ.Enqueue "Hello"
    myQ.Enqueue "World"
    myQ.Enqueue "!"
    
    ' Displays the properties and values of the Queue.
    Debug.Print "myQ"
    Debug.Print VBAString.Format(Regex.Unescape("\tCount:    {0}"), myQ.count) 'Note Unescape .Net escape characters \t i.e. tab
    PrintValues myQ
End Sub

Private Sub PrintValues(ByVal myCollection As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myCollection
        Debug.Print VBAString.Format("    {0}", obj);
    Next
    Debug.Print
End Sub

' /*
' This code produces the following output.
'
' myQ
'     Count:    3
'     Values:    Hello    World    !
'*/
