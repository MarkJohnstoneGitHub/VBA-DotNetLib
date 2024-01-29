Attribute VB_Name = "BitArrayCopyToExample"
'@IgnoreModule IndexedDefaultMemberAccess
'@Folder "Examples.System.Collections.BitArray.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 2, 2023
'@LastModified January 28, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.copyto?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example shows how to copy a BitArray into a
' one-dimensional Array.
''
Public Sub BitArrayCopyToExample()
    ' Creates and initializes the source BitArray.
    Dim myBA As DotNetLib.BitArray
    Set myBA = BitArray.Create(4)
    myBA(3) = True
    myBA(2) = myBA(3)
    myBA(1) = myBA(2)
    myBA(0) = myBA(1)

    ' Creates and initializes the one-dimensional target Array of type Boolean.
    Dim myBoolArray As DotNetLib.Array
    Set myBoolArray = Arrays.CreateInstance(Booleans.GetType(), 8)
    Call myBoolArray.SetValue(False, 0)
    Call myBoolArray.SetValue(False, 1)

    ' Displays the values of the target Array.
    Debug.Print "The target Boolean Array contains the following (before and after copying):"
    Call PrintValues(myBoolArray)
    
    ' Copies the entire source BitArray to the target BitArray, starting at index 3.
    Call myBA.CopyTo(myBoolArray, 3)
    
    ' Displays the values of the target Array.
    Call PrintValues(myBoolArray)
    
    ' Creates and initializes the one-dimensional target Array of type integer.
    Dim myIntArray As DotNetLib.Array
    Set myIntArray = Arrays.CreateInstance(Int32.GetType(), 8)
    Call myIntArray.SetValue(42, 0)
    Call myIntArray.SetValue(43, 1)

    ' Displays the values of the target Array.
    Debug.Print "The target integer Array contains the following (before and after copying):"
    Call PrintValues(myIntArray)
    
    ' Copies the entire source BitArray to the target BitArray, starting at index 3.
    Call myBA.CopyTo(myIntArray, 3)

    ' Displays the values of the target Array.
    Call PrintValues(myIntArray)
    
    ' Creates and initializes the one-dimensional target Array of type byte.
    Dim myByteArray As DotNetLib.Array
    Set myByteArray = Arrays.CreateInstance(bytes.GetType(), 8)
    Call myByteArray.SetValue(CByte(10), 0)
    Call myByteArray.SetValue(CByte(11), 1)

    ' Displays the values of the target Array.
    Debug.Print "The target byte Array contains the following (before and after copying):"
    Call PrintValues(myByteArray)

    ' Copies the entire source BitArray to the target BitArray, starting at index 3.
    Call myBA.CopyTo(myByteArray, 3)

    ' Displays the values of the target Array.
    Call PrintValues(myByteArray)

    ' Returns an exception if the array is not of type Boolean, integer or byte.
    On Error Resume Next
    Dim myStringArray As DotNetLib.Array
    Set myStringArray = Arrays.CreateInstance(VBString.GetType(), 8)
    Call myStringArray.SetValue("Hello", 0)
    Call myStringArray.SetValue("World", 1)
    Call myBA.CopyTo(myStringArray, 3)
    If Err.Number Then
        Debug.Print "Exception: " + Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBString.Format("{0,8}", obj);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The target Boolean Array contains the following (before and after copying):
'   False   False   False   False   False   False   False   False
'   False   False   False    True    True    True    True   False
'The target integer Array contains the following (before and after copying):
'      42      43       0       0       0       0       0       0
'      42      43       0      15       0       0       0       0
'The target byte Array contains the following (before and after copying):
'      10      11       0       0       0       0       0       0
'      10      11       0      15       0       0       0       0
'Exception: System.ArgumentException: Only supported array types for CopyTo on BitArrays are Boolean[], Int32[] and Byte[].
'   at System.Collections.BitArray.CopyTo(Array array, int index)
'   at SamplesBitArray.Main()
'
'*/
