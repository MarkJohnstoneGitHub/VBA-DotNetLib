Attribute VB_Name = "BitArrayClassExample"
'@Folder "Examples.System.Collections.BitArray"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified November 2, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray?view=netframework-4.8.1#examples

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'@Dependencies ArrayEx.cls

Option Explicit

''
' The following code example shows how to create and initialize a BitArray and
' how to print out its values.
''
Public Sub BitArrayClassExample()
    ' Creates and initializes several BitArrays.
    Dim myBA1 As DotNetLib.BitArray
    Set myBA1 = BitArray.Create(5)
    
    Dim myBA2 As DotNetLib.BitArray
    Set myBA2 = BitArray.Create2(5, False)

    Dim myBytes() As Byte
    Call VBArray.CreateInitialize1D(myBytes, 1, 2, 3, 4, 5)
    Dim myBA3 As DotNetLib.BitArray
    Set myBA3 = BitArray.Create3(myBytes)

    Dim myBools() As Boolean
    Call VBArray.CreateInitialize1D(myBools, True, False, True, True, False)
    Dim myBA4 As DotNetLib.BitArray
    Set myBA4 = BitArray.Create4(myBools)

    Dim myInts() As Long
    Call VBArray.CreateInitialize1D(myInts, 6, 7, 8, 9, 10)
    Dim myBA5 As DotNetLib.BitArray
    Set myBA5 = BitArray.Create5(myInts)
    
    ' Displays the properties and values of the BitArrays.
    Debug.Print "myBA1"
    Debug.Print VBString.Format("   Count:    {0}", myBA1.Count)
    Debug.Print VBString.Format("   Length:   {0}", myBA1.Length)
    Debug.Print "   Values:"
    Call PrintValues(myBA1, 8)
    
    Debug.Print "myBA2"
    Debug.Print VBString.Format("   Count:    {0}", myBA2.Count)
    Debug.Print VBString.Format("   Length:   {0}", myBA2.Length)
    Debug.Print "   Values:"
    Call PrintValues(myBA2, 8)

    Debug.Print "myBA3"
    Debug.Print VBString.Format("   Count:    {0}", myBA3.Count)
    Debug.Print VBString.Format("   Length:   {0}", myBA3.Length)
    Debug.Print "   Values:"
    Call PrintValues(myBA3, 8)

    Debug.Print "myBA4"
    Debug.Print VBString.Format("   Count:    {0}", myBA4.Count)
    Debug.Print VBString.Format("   Length:   {0}", myBA4.Length)
    Debug.Print "   Values:"
    Call PrintValues(myBA4, 8)

    Debug.Print "myBA5"
    Debug.Print VBString.Format("   Count:    {0}", myBA5.Count)
    Debug.Print VBString.Format("   Length:   {0}", myBA5.Length)
    Debug.Print "   Values:"
    Call PrintValues(myBA5, 8)
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
'myBA1
'   Count:    5
'   Length:   5
'Values:
'   False   False   False   False   False
'myBA2
'   Count:    5
'   Length:   5
'Values:
'   False   False   False   False   False
'myBA3
'   Count:    40
'   Length:   40
'Values:
'    True   False   False   False   False   False   False   False
'   False    True   False   False   False   False   False   False
'    True    True   False   False   False   False   False   False
'   False   False    True   False   False   False   False   False
'    True   False    True   False   False   False   False   False
'myBA4
'   Count:    5
'   Length:   5
'Values:
'    True   False    True    True   False
'myBA5
'   Count:    160
'   Length:   160
'Values:
'   False    True    True   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'    True    True    True   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False    True   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'    True   False   False    True   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False    True   False    True   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'   False   False   False   False   False   False   False   False
'*/
