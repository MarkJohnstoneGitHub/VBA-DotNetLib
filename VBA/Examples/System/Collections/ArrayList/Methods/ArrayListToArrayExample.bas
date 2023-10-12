Attribute VB_Name = "ArrayListToArrayExample"
'@Folder("Examples.System.Collections.ArrayList.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 6, 2023
'@LastModified October 6, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.toarray?view=netframework-4.8.1#system-collections-arraylist-toarray

Option Explicit

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
    Debug.Print Strings.Format("The ArrayList contains the following values:")
    PrintIndexAndValues myAL
    
    ' Copies the elements of the ArrayList to a string array.
    Dim myArr() As Variant
    myArr = myAL.ToArray()
    '      String[] myArr = (String[]) myAL.ToArray( typeof( string ) );
'
'      // Displays the contents of the string array.
'      Console.WriteLine( "The string array contains the following values:" );
'      PrintIndexAndValues( myArr );
End Sub

'   public static void PrintIndexAndValues( ArrayList myList )  {
'      int i = 0;
'      foreach ( Object o in myList )
'         Console.WriteLine( "\t[{0}]:\t{1}", i++, o );
'      Console.WriteLine();
'   }

Private Sub PrintIndexAndValues(ByVal myList As DotNetLib.ArrayList)
    Dim pvtFormat As String
    pvtFormat = Regex.Unescape("\t[{0}]:\t{1}")
    Dim obj As Variant
    Dim i As Long
    For Each obj In myList
        Debug.Print Strings.Format(pvtFormat, i, obj)
        i = i + 1
    Next
    Debug.Print
End Sub

Private Sub PrintIndexAndValues2(ByRef myArr() As String)
    Dim i As Long
    Dim pvtFormat As String
    pvtFormat = Regex.Unescape("\t[{0}]:\t{1}")
    For i = 0 To UBound(myArr)
        Debug.Print Strings.Format(pvtFormat, i, myArr(i))
    Next i
    Debug.Print
End Sub




'using System;
'using System.Collections;
'
'public class SamplesArrayList  {
'
'   public static void Main()  {
'
'      // Creates and initializes a new ArrayList.
'      ArrayList myAL = new ArrayList();
'      myAL.Add( "The" );
'      myAL.Add( "quick" );
'      myAL.Add( "brown" );
'      myAL.Add( "fox" );
'      myAL.Add( "jumps" );
'      myAL.Add( "over" );
'      myAL.Add( "the" );
'      myAL.Add( "lazy" );
'      myAL.Add( "dog" );
'
'      // Displays the values of the ArrayList.
'      Console.WriteLine( "The ArrayList contains the following values:" );
'      PrintIndexAndValues( myAL );
'
'      // Copies the elements of the ArrayList to a string array.
'      String[] myArr = (String[]) myAL.ToArray( typeof( string ) );
'
'      // Displays the contents of the string array.
'      Console.WriteLine( "The string array contains the following values:" );
'      PrintIndexAndValues( myArr );
'   }
'
'   public static void PrintIndexAndValues( ArrayList myList )  {
'      int i = 0;
'      foreach ( Object o in myList )
'         Console.WriteLine( "\t[{0}]:\t{1}", i++, o );
'      Console.WriteLine();
'   }
'
'   public static void PrintIndexAndValues( String[] myArr )  {
'      for ( int i = 0; i < myArr.Length; i++ )
'         Console.WriteLine( "\t[{0}]:\t{1}", i, myArr[i] );
'      Console.WriteLine();
'   }
'}
'
'
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
