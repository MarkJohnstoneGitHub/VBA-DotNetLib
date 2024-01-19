Attribute VB_Name = "ArrayListRemoveAtExample"
'@Folder "Examples.System.Collections.ArrayList.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 25, 2023
'@LastModified October 25, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist.removeat?view=netframework-4.8.1#examples

Option Explicit

Public Sub ArrayListRemoveAt()
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
    
    ' Displays the ArrayList.
    Debug.Print "The ArrayList initially contains the following:"
    Call PrintValues(myAL)
    
End Sub

Private Sub PrintValues(ByVal myList As mscorlib.IEnumerable)
    Dim obj As Variant
    For Each obj In myList
        Debug.Print VBString.Format("   {0}", obj);
    Next
    Debug.Print
End Sub


'using System;
'using System.Collections;
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
'      // Displays the ArrayList.
'      Console.WriteLine( "The ArrayList initially contains the following:" );
'      PrintValues( myAL );
'
'      // Removes the element containing "lazy".
'      myAL.Remove( "lazy" );
'
'      // Displays the current state of the ArrayList.
'      Console.WriteLine( "After removing \"lazy\":" );
'      PrintValues( myAL );
'
'      // Removes the element at index 5.
'      myAL.RemoveAt( 5 );
'
'      // Displays the current state of the ArrayList.
'      Console.WriteLine( "After removing the element at index 5:" );
'      PrintValues( myAL );
'
'      // Removes three elements starting at index 4.
'      myAL.RemoveRange( 4, 3 );
'
'      // Displays the current state of the ArrayList.
'      Console.WriteLine( "After removing three elements starting at index 4:" );
'      PrintValues( myAL );
'   }
'
'   public static void PrintValues( IEnumerable myList )  {
'      foreach ( Object obj in myList )
'         Console.Write( "   {0}", obj );
'      Console.WriteLine();
'   }
'}
'/*
'This code produces the following output.
'
'The ArrayList initially contains the following:
'   The   quick   brown   fox   jumps   over   the   lazy   dog
'After removing "lazy":
'   The   quick   brown   fox   jumps   over   the   dog
'After removing the element at index 5:
'   The   quick   brown   fox   jumps   the   dog
'After removing three elements starting at index 4:
'   The   quick   brown   fox
'*/
