Attribute VB_Name = "StackPeekExample"
'@Folder "Examples.System.Collections.Stack.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack.peek?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to add elements to the Stack, remove elements
' from the Stack, or view the element at the top of the Stack.
''
Public Sub StackPeek()
    ' Creates and initializes a new Stack.
    Dim myStack As DotNetLib.Stack
    Set myStack = Stack.Create()
    myStack.Push "The"
    myStack.Push "quick"
    myStack.Push "brown"
    myStack.Push "fox"

    '  Displays the Stack.
    Debug.Print "Stack values:";
    Call PrintValues(myStack, Regex.Unescape("\t"))

    ' Removes an element from the Stack.
    Debug.Print VBString.Format(Regex.Unescape("(Pop)\t\t{0}"), myStack.Pop())

    '  Displays the Stack.
    Debug.Print "Stack values:";
    Call PrintValues(myStack, Regex.Unescape("\t"))
    
    ' Removes another element from the Stack.
    Debug.Print VBString.Format(Regex.Unescape("(Pop)\t\t{0}"), myStack.Pop())

    '  Displays the Stack.
    Debug.Print "Stack values:";
    Call PrintValues(myStack, Regex.Unescape("\t"))
    
    ' Views the first element in the Stack but does not remove it.
    Debug.Print VBString.Format(Regex.Unescape("(Peek)\t\t{0}"), myStack.Peek())

    '  Displays the Stack.
    Debug.Print "Stack values:";
    Call PrintValues(myStack, Regex.Unescape("\t"))
    
End Sub

Private Sub PrintValues(ByVal myCollection As mscorlib.IEnumerable, ByVal mySeparator As String)
    Dim obj As Variant
    For Each obj In myCollection
        Debug.Print VBString.Format("{0}{1}", mySeparator, obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'Stack values:    fox    brown    quick    The
'(Pop)        fox
'Stack values:    brown    quick    The
'(Pop)        brown
'Stack Values:    quick The
'(Peek)        quick
'Stack Values:    quick The
'*/
