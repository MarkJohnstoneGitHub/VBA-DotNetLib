Attribute VB_Name = "BitArrayXorExample"
'@Folder("Examples.System.Collections.BitArray.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified November 2, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.xor?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to perform the XOR operation between
' two BitArray objects.
''
Public Sub BitArrayXorExample()
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
        
    ' Performs a bitwise XOR operation between BitArray instances of the same size.
    Debug.Print "Initial values"
    Debug.Print "myBA1:";
    Call PrintValues(myBA1, 8)
    Debug.Print "myBA2:";
    Call PrintValues(myBA2, 8)
    Debug.Print
    
    Debug.Print "Result"
    Debug.Print "XOR:";
    Call PrintValues(myBA1.Xor(myBA2), 8)
    Debug.Print
    
    Debug.Print "After XOR"
    Debug.Print "myBA1:";
    Call PrintValues(myBA1, 8)
    Debug.Print "myBA2:";
    Call PrintValues(myBA2, 8)
    Debug.Print

    ' Performing XOR between BitArray instances of different sizes returns an exception.
    On Error Resume Next
    Dim myBA3 As DotNetLib.BitArray
    Set myBA3 = BitArray.Create(8)
    myBA3(3) = False
    myBA3(2) = myBA3(3)
    myBA3(1) = myBA3(2)
    myBA3(0) = myBA3(1)
    
    myBA3(7) = True
    myBA3(6) = myBA3(7)
    myBA3(5) = myBA3(6)
    myBA3(4) = myBA3(5)
    Call myBA1.Xor(myBA3)
    If Err.Number Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
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
        Debug.Print VBAString.Format("{0,8}", obj);
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
'Result
'XOR:   False    True    True   False
'
'After XOR
'myBA1:   False    True    True   False
'myBA2:   False    True   False    True
'
'Exception: System.ArgumentException: Array lengths must be the same.
'   at System.Collections.BitArray.Xor(BitArray value)
'   at SamplesBitArray.Main()
'
'*/
