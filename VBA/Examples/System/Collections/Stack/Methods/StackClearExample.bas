Attribute VB_Name = "StackClearExample"
'@Folder("Examples.System.Collections.Stack.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack.clear?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to clear the values of the Stack.
''
Public Sub StackClear()
    ' Creates and initializes a new Stack.
    Dim myStack As DotNetLib.Stack
    Set myStack = Stack.Create()
    myStack.Push "The"
    myStack.Push "quick"
    myStack.Push "brown"
    myStack.Push "fox"
    myStack.Push "jumps"
    
    ' Displays the count and values of the Stack.
    Debug.Print "Initially,"
    Debug.Print VBAString.Format("   Count    : {0}", myStack.count)
    Debug.Print "   Values:";
    PrintValues myStack
    
    ' Clears the Stack.
    myStack.Clear

    ' Displays the count and values of the Stack.
    Debug.Print "After Clear,"
    Debug.Print VBAString.Format("   Count    : {0}", myStack.count)
    Debug.Print "   Values:";
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
'Initially,
'   Count    : 5
'   Values:    jumps    fox    brown    quick    The
'After Clear,
'   Count    : 0
'Values:
'*/
