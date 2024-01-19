Attribute VB_Name = "SortedListCopyToExample"
'@Folder "Examples.System.Collections.SortedList.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 18, 2023
'@LastModified October 27, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist.copyto?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to copy the values in a SortedList
' object into a one-dimensional Array object.
''
Public Sub SortedListCopyTo()
    ' Creates and initializes the source SortedList.
    Dim mySourceList As DotNetLib.SortedList
    Set mySourceList = SortedList.Create()
    Call mySourceList.Add(2, "cats")
    Call mySourceList.Add(3, "in")
    Call mySourceList.Add(1, "napping")
    Call mySourceList.Add(4, "the")
    Call mySourceList.Add(0, "three")
    Call mySourceList.Add(5, "barn")
    
    ' Creates and initializes the one-dimensional target Array.
    Dim tempArray As DotNetLib.Array
    Set tempArray = Arrays.CreateInitialize1D(VBString.GetType(), "The", "quick", "brown", "fox", "jumps", "over", "the", "lazy", "dog")
    
    'Create an array of mscorlib.DictionaryEntry of size 15
    Dim myTargetArray As DotNetLib.Array
    Set myTargetArray = Arrays.CreateInstance(DictionaryEntry.GetType(), 15)
    
    Dim i As Long
    i = 0
    Dim s As Variant
    For Each s In tempArray
        Dim dictEntry As mscorlib.DictionaryEntry
        Call DictionaryEntry.Create(dictEntry, i, s)
        Call myTargetArray.SetValue(dictEntry, i)
        i = i + 1
    Next
    
    ' Displays the values of the target Array.
    Debug.Print "The target Array contains the following (before and after copying):"
    Call PrintValues(myTargetArray, " ")

    ' Copies the entire source SortedList to the target SortedList, starting at index 6.
    Call mySourceList.CopyTo(myTargetArray, 6)
    ' Displays the values of the target Array.
    Call PrintValues(myTargetArray, " ")
End Sub

Private Sub PrintValues(ByVal myArr As DotNetLib.Array, ByVal mySeparator As String)
    Dim i As Long
    For i = 0 To myArr.Length - 1
        Dim de As mscorlib.DictionaryEntry
        de = myArr(i)
        Debug.Print VBString.Format("{0}{1}", mySeparator, DictionaryEntry.value(de));
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The target Array contains the following (before and after copying):
' The quick brown fox jumps over the lazy dog
' The quick brown fox jumps over three napping cats in the barn
'
'*/
