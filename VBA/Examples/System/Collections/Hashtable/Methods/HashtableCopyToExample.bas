Attribute VB_Name = "HashtableCopyToExample"
'@Folder("Examples.System.Collections.Hashtable.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 22, 2023
'@LastModified October 22, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable.copyto?view=netframework-4.8.1#examples

'@Remarks
' Using CopyTo(Array,arrayIndex) requires the Array.WrappedArray member i.e. the mscorlib.Array object

Option Explicit

''
' The following example shows how to copy the list of keys or the list of
' values in a Hashtable into a one-dimensional Array.
''
Public Sub HashtableCopyTo()
    ' Creates and initializes the source Hashtable.
    Dim mySourceHT As DotNetLib.Hashtable
    Set mySourceHT = Hashtable.Create()
    Call mySourceHT.Add("A", "valueA")
    Call mySourceHT.Add("B", "valueB")
    
    ' Creates and initializes the one-dimensional target Array.
    Dim myTargetArray As DotNetLib.Array
    Set myTargetArray = Arrays.CreateInstance(VBAString.GetType(), 15)
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
    Debug.Print "The target Array contains the following before:"
    Call PrintValues(myTargetArray, " ")
    
    ' Copies the keys in the source Hashtable to the target Hashtable, starting at index 6.
    Debug.Print "After copying the keys, starting at index 6:"
    Call mySourceHT.keys.CopyTo(myTargetArray.WrappedArray, 6)
    Call PrintValues(myTargetArray, " ")
    
    ' Copies the values in the source Hashtable to the target Hashtable, starting at index 6.
    Call mySourceHT.values.CopyTo(myTargetArray.WrappedArray, 6)
    ' Displays the values of the target Array.
    Call PrintValues(myTargetArray, " ")
End Sub

Public Sub PrintValues(ByVal myArr As DotNetLib.Array, ByVal mySeparator As String)
    Dim i As Long
    For i = 0 To myArr.length - 1
        Debug.Print VBAString.Format("{0}{1}", mySeparator, myArr(i));
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The target Array contains the following before:
' The quick brown fox jumps over the lazy dog
'After copying the keys, starting at index 6:
' The quick brown fox jumps over A B dog
'After copying the values, starting at index 6:
' The quick brown fox jumps over valueA valueB dog
'
'*/

