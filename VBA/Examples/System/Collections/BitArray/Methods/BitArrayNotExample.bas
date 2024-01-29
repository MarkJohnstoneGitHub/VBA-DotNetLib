Attribute VB_Name = "BitArrayNotExample"
'@IgnoreModule IndexedDefaultMemberAccess
'@Folder "Examples.System.Collections.BitArray.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified January 28, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.not?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to apply NOT to a BitArray.
''
Public Sub BitArrayNotExample()
    ' Creates and initializes two BitArrays of the same size.
    Dim myBA1 As DotNetLib.BitArray
    Set myBA1 = BitArray.Create(4)
    Dim myBA2 As DotNetLib.BitArray
    Set myBA2 = BitArray.Create(4)
    myBA1(1) = False
    myBA1(0) = myBA1(1)
    myBA1(3) = True
    myBA1(2) = myBA1(3)
    myBA1(3) = True
    myBA1(2) = myBA1(3)
    myBA2(2) = False
    myBA2(0) = myBA2(2)
    myBA2(3) = True
    myBA2(1) = myBA2(3)
    
    ' Performs a bitwise NOT operation between BitArray instances of the same size.
    Debug.Print "Initial values"
    Debug.Print "myBA1:";
    Call PrintValues(myBA1, 8)
    Debug.Print "myBA2:";
    Call PrintValues(myBA2, 8)
    Debug.Print
    
    Call myBA1.Not
    Call myBA2.Not
    
    Debug.Print "After NOT"
    Debug.Print "myBA1:";
    Call PrintValues(myBA1, 8)
    Debug.Print "myBA2:";
    Call PrintValues(myBA2, 8)
    Debug.Print
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable, ByVal myWidth As Long)
    Dim i As Long
    i = myWidth
    Dim obj As Variant
    For Each obj In myList
        If (i <= 0) Then
            i = myWidth
            Debug.Print
        End If
        i = i - 1
        Debug.Print VBString.Format("{0,8}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'Initial Values
'myBA1:   False   False    True    True
'myBA2:   False    True   False    True
'
'After NOT
'myBA1:    True    True   False   False
'myBA2:    True   False    True   False
'
'*/
