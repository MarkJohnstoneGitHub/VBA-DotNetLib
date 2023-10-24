Attribute VB_Name = "ArrayListCopyToExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 24, 2023
'@LastModified October 24, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.copyto?view=netframework-4.8.1#system-collections-arraylist-copyto(system-array)

'@Remarks
' For assinging value types to an Array eg. myTargetArray[0] = "The" use SetValue(value,index) member.
' eg. myTargetArray.SetValue("The", 0)
Option Explicit

''
' The following code example shows how to copy an ArrayList into a
' one-dimensional System.Array.
''
Public Sub ArrayListCopyTo()
    ' Creates and initializes the source ArrayList.
    Dim mySourceList As DotNetLib.ArrayList
    Set mySourceList = ArrayList.Create()
    Call mySourceList.Add("three")
    Call mySourceList.Add("napping")
    Call mySourceList.Add("cats")
    Call mySourceList.Add("in")
    Call mySourceList.Add("the")
    Call mySourceList.Add("barn")
    
    ' Creates and initializes the one-dimensional target Array.
    Dim myTargetArray As DotNetLib.Array
    Set myTargetArray = Arrays.CreateInstance(Strings.GetType(), 15)
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

    ' Copies the second element from the source ArrayList to the target Array starting at index 7.
    Call mySourceList.CopyTo3(1, myTargetArray, 7, 1)

    ' Displays the values of the target Array.
    Call PrintValues(myTargetArray, " ")

    ' Copies the entire source ArrayList to the target Array starting at index 6.
    Call mySourceList.CopyTo2(myTargetArray, 6)

    ' Displays the values of the target Array.
    Call PrintValues(myTargetArray, " ")

    ' Copies the entire source ArrayList to the target Array starting at index 0.
    Call mySourceList.CopyTo(myTargetArray)

    ' Displays the values of the target Array.
    Call PrintValues(myTargetArray, " ")
End Sub

Private Sub PrintValues(ByVal myArr As DotNetLib.Array, ByVal mySeparator As String)
    Dim i As Long
    For i = 0 To myArr.Length - 1
        Debug.Print Strings.Format("{0}{1}", mySeparator, myArr(i));
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The target Array contains the following (before and after copying):
' The quick brown fox jumps over the lazy dog
' The quick brown fox jumps over the napping dog
' The quick brown fox jumps over three napping cats in the barn
' three napping cats in the barn three napping cats in the barn
'*/
