Attribute VB_Name = "SortedListCopyToExample"
'@Folder("Examples.System.Collections.SortedList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 18, 2023
'@LastModified October 22 2023
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
    Set tempArray = Arrays.CreateInstanceInitialize(Strings.GetType(), "The", "quick", "brown", "fox", "jumps", "over", "the", "lazy", "dog")
    
    'Create an array of mscorlib.DictionaryEntry of size 15
    Dim myTargetArray As DotNetLib.Array
    Set myTargetArray = Arrays.CreateInstance(Types.GetType("System.Collections.DictionaryEntry"), 15)
    
    Dim i As Long
    i = 0
    Dim s As Variant
    For Each s In tempArray
        'Create an mscorlib.DictionaryEntry
        Dim wrappedDictEntry As DotNetLib.DictionaryEntry
        Set wrappedDictEntry = DictionaryEntry.Create3(i, s)
        Dim de As mscorlib.DictionaryEntry
        Call wrappedDictEntry.GetDictionaryEntry(de)
        'Call DictionaryEntry.Create3(i, s).GetDictionaryEntry(de) 'Note Could replace above to create an mscorlib.DictionaryEntry
        Call myTargetArray.SetValue(de, i)
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
        Dim wrappedDictEntry As DotNetLib.DictionaryEntry
        Set wrappedDictEntry = DictionaryEntry.Create2(de)
        Debug.Print Strings.Format("{0}{1}", mySeparator, wrappedDictEntry.Value);
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

'using System;
' using System.Collections;
' public class SamplesSortedList  {
'
'    public static void Main()  {
'
'       // Creates and initializes the source SortedList.
'       SortedList mySourceList = new SortedList();
'       mySourceList.Add( 2, "cats" );
'       mySourceList.Add( 3, "in" );
'       mySourceList.Add( 1, "napping" );
'       mySourceList.Add( 4, "the" );
'       mySourceList.Add( 0, "three" );
'       mySourceList.Add( 5, "barn" );
'
'       // Creates and initializes the one-dimensional target Array.
'       String[] tempArray = new String[] { "The", "quick", "brown", "fox", "jumps", "over", "the", "lazy", "dog" };
'       DictionaryEntry[] myTargetArray = new DictionaryEntry[15];
'       int i = 0;
'       foreach ( string s in tempArray )  {
'          myTargetArray[i].Key = i;
'          myTargetArray[i].Value = s;
'          i++;
'       }
'
'       // Displays the values of the target Array.
'       Console.WriteLine( "The target Array contains the following (before and after copying):" );
'       PrintValues( myTargetArray, ' ' );
'
'       // Copies the entire source SortedList to the target SortedList, starting at index 6.
'       mySourceList.CopyTo( myTargetArray, 6 );
'
'       // Displays the values of the target Array.
'       PrintValues( myTargetArray, ' ' );
'    }
'
'    public static void PrintValues( DictionaryEntry[] myArr, char mySeparator )  {
'       for ( int i = 0; i < myArr.Length; i++ )
'          Console.Write( "{0}{1}", mySeparator, myArr[i].Value );
'       Console.WriteLine();
'    }
' }
'
'
