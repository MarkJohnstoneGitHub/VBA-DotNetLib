Attribute VB_Name = "ArrayExample"
'@Folder("Examples.System.Array")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 8, 2023
'@LastModified October 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1

Option Explicit

Public Sub ArrayExample()
    ' Creates and initializes a one-dimensional Array of type int32.
    Dim longType As DotNetLib.Type
    
    'Set longType = Objects.Create(CLng(0)).GetType()
    'Set longType = Types.GetType("System.Int32")
    
    Dim my1DArray As DotNetLib.Array
    Set my1DArray = Arrays.CreateInstance(Int32.GetType(), 5)   'Int32.GetType()
    
    Dim i As Long
    For i = my1DArray.GetLowerBound(0) To my1DArray.GetUpperBound(0)
        my1DArray.SetValue (i + 1), i
    Next i
    
    ' Displays the values of the Array.
    Debug.Print "The one-dimensional Array contains the following values:"
    PrintValues my1DArray
    
    'Test resizing array
    Arrays.Resize my1DArray, 10
    PrintValues my1DArray
    
    'Test sorting array
    Arrays.Sort my1DArray
    PrintValues my1DArray
    
    'Test reversing array
    Arrays.Reverse my1DArray
    PrintValues my1DArray
    
End Sub

Private Sub PrintValues(ByVal myArr As DotNetLib.Array)
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}")
    Dim cols As Long
    cols = myArr.GetLength(myArr.Rank - 1)
    Dim i As Long
    i = 0
    Dim varItem As Variant
    For Each varItem In myArr
        If (i < cols) Then
             i = i + 1
        Else
            Debug.Print
            i = 1
        End If
        Debug.Print Strings.Format(formatString, varItem);
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The one-dimensional Array contains the following values:
'    1    2    3    4    5
'*/

