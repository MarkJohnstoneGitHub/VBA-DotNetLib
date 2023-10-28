Attribute VB_Name = "ArrayRankExample"
'@Folder("Examples.System.Array.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 27, 2023
'@LastModified October 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.rank?view=netframework-4.8.1#examples

Option Explicit

''
' The following example initializes a one-dimensional array, a two-dimensional
' array, and retrieves the Rank property of each.
''
Public Sub ArrayRank()
    Dim array1 As DotNetLib.Array
    Set array1 = Arrays.CreateInstance(Int32.GetType(), 10)
    
    Dim array2 As DotNetLib.Array
    Set array2 = Arrays.CreateInstance2(Int32.GetType(), 10, 3)
    
    Debug.Print Strings.Format("{0}: {1} dimension(s)", _
                        array1.ToString(), array1.Rank)
    Debug.Print Strings.Format("{0}: {1} dimension(s)", _
                        array2.ToString(), array2.Rank)
End Sub

' The example displays the following output:
'       System.Int32[]: 1 dimension(s)
'       System.Int32[,]: 2 dimension(s)


