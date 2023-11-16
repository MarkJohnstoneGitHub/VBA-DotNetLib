Attribute VB_Name = "StackExample"
'@Folder("Examples.System.Collections.Stack")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to create and add values to a Stack and how
' to display its values.
''
Public Sub StackExample()
    ' Creates and initializes a new Stack.
    Dim myStack As DotNetLib.Stack
    Set myStack = Stack.Create()
    myStack.Push "Hello"
    myStack.Push "World"
    myStack.Push "!"
    
    ' Displays the properties and values of the Stack.
    Debug.Print "myStack"
    Debug.Print VBAString.Format(Regex.Unescape("\tCount:    {0}"), myStack.count) 'Note Unescape .Net escape characters \t i.e. tab
    Debug.Print VBAString.Format(Regex.Unescape("\tValues:"));
    PrintValues myStack
End Sub

Private Sub PrintValues(ByVal myCollection As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myCollection
        Debug.Print VBAString.Format("    {0}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'myStack
'    Count:    3
'    Values:    !    World    Hello
'*/
