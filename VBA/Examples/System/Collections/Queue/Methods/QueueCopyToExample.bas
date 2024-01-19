Attribute VB_Name = "QueueCopyToExample"
'@Folder "Examples.System.Collections.Queue.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 15, 2023
'@LastModified October 15, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue.copyto?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to copy a Queue into a one-dimensional array.
''
Public Sub QueueCopyTo()
    ' Creates and initializes the source Queue.
    Dim mySourceQ As DotNetLib.Queue
    Set mySourceQ = Queue.Create()
    mySourceQ.Enqueue "three"
    mySourceQ.Enqueue "napping"
    mySourceQ.Enqueue "cats"
    mySourceQ.Enqueue "in"
    mySourceQ.Enqueue "the"
    mySourceQ.Enqueue "barn"

    ' Creates and initializes the one-dimensional target Array.
    Dim myTargetArray As DotNetLib.Array
    Set myTargetArray = Arrays.CreateInstance(VBString.GetType(), 15)
    Call myTargetArray.SetValue("The", 0)
    Call myTargetArray.SetValue("quick", 1)
    Call myTargetArray.SetValue("brown", 2)
    Call myTargetArray.SetValue("fox", 3)
    Call myTargetArray.SetValue("jumps", 4)
    Call myTargetArray.SetValue("over", 5)
    Call myTargetArray.SetValue("the", 6)
    Call myTargetArray.SetValue("lazy", 7)
    Call myTargetArray.SetValue("dog", 8)
    
    ' Displays the values of the target Array.
    Debug.Print "The target Array contains the following (before and after copying):"
    Call PrintValues(myTargetArray, " ")
    
    ' Copies the entire source Queue to the target Array, starting at index 6.
    Call mySourceQ.CopyTo(myTargetArray, 6)

    ' Displays the values of the target Array.
    Call PrintValues(myTargetArray, " ")
    
    ' Copies the entire source Queue to a new standard array.
    Dim myStandardArray As DotNetLib.Array
    Set myStandardArray = mySourceQ.ToArray()
    Call PrintValues(myStandardArray, " ")
End Sub

Private Sub PrintValues(ByVal myArr As DotNetLib.Array, ByVal mySeparator As String)
    Dim myObj As Variant
    For Each myObj In myArr
        Debug.Print VBString.Format("{0}{1}", mySeparator, myObj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The target Array contains the following (before and after copying):
' The quick brown fox jumps over the lazy dog
' The quick brown fox jumps over three napping cats in the barn
'The new standard array contains the following:
' three napping cats in the barn
'
'*/
