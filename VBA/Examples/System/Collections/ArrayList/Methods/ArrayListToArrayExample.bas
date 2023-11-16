Attribute VB_Name = "ArrayListToArrayExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.toarray?view=netframework-4.8.1#system-collections-arraylist-toarray

Option Explicit

''
' The following copy example shows how to copy the elements of an ArrayList
' to a string array.
''
Public Sub ArrayListToArray()
    ' Creates and initializes a new ArrayList.
    Dim myAL As DotNetLib.ArrayList
    Set myAL = ArrayList.Create()
    myAL.Add "The"
    myAL.Add "quick"
    myAL.Add "brown"
    myAL.Add "fox"
    myAL.Add "jumps"
    myAL.Add "over"
    myAL.Add "the"
    myAL.Add "lazy"
    myAL.Add "dog"
    
    ' Displays the values of the ArrayList.
    Debug.Print VBAString.Format("The ArrayList contains the following values:")
    PrintIndexAndValues myAL
    
    ' Copies the elements of the ArrayList to a string array.
    Dim myArr As DotNetLib.Array
    Set myArr = myAL.ToArray2(VBAString.GetType())
    Debug.Print "The string array contains the following values:"
    Call PrintIndexAndValues2(myArr)
End Sub

Private Sub PrintIndexAndValues(ByVal myList As DotNetLib.ArrayList)
    Dim pvtFormat As String
    pvtFormat = Regex.Unescape("\t[{0}]:\t{1}")
    Dim obj As Variant
    Dim i As Long
    For Each obj In myList
        Debug.Print VBAString.Format(pvtFormat, i, obj)
        i = i + 1
    Next
    Debug.Print
End Sub

Private Sub PrintIndexAndValues2(ByVal myArr As DotNetLib.Array)
    Dim i As Long
    Dim pvtFormat As String
    pvtFormat = Regex.Unescape("\t[{0}]:\t{1}")
    For i = 0 To myArr.length - 1
        Debug.Print VBAString.Format(pvtFormat, i, myArr(i))
    Next i
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The ArrayList contains the following values:
'0:              the
'1:              quick
'2:              brown
'3:              fox
'4:              jumps
'5:              over
'6:              the
'7:              lazy
'8:              dog
'
'The string array contains the following values:
'0:              the
'1:              quick
'2:              brown
'3:              fox
'4:              jumps
'5:              over
'6:              the
'7:              lazy
'8:              dog
'
'*/
