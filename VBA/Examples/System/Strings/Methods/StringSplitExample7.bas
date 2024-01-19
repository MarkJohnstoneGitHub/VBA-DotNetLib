Attribute VB_Name = "StringSplitExample7"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 6, 2024
'@LastModified January 6, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.split?view=netframework-4.8.1#system-string-split(system-string()-system-stringsplitoptions)

Option Explicit

''
' The following example defines an array of separators that include punctuation
' and white-space characters. Passing this array along with a value of
' StringSplitOptions.RemoveEmptyEntries to the Split(String[], StringSplitOptions)
' method returns an array that consists of the individual words from the string.
''
Public Sub StringSplitExample7()
    Dim separators() As String
    Call VBArray.CreateInitialize1D(separators, ",", ".", "!", "?", ";", ":", " ")
    'separators = StringArray.Initialize1D(",", ".", "!", "?", ";", ":", " ")
    'Call VBArray.Initialize1D(separators, ",", ".", "!", "?", ";", ":", " ")
    'Call VBString.Format()
    

End Sub

'string[] separators = { ",", ".", "!", "?", ";", ":", " " };
'string value = "The handsome, energetic, young dog was playing with his smaller, more lethargic litter mate.";
'string[] words = value.Split(separators, StringSplitOptions.RemoveEmptyEntries);
'foreach (var word in words)
'    Console.WriteLine(word);
'
'// The example displays the following output:
'//       The
'//       handsome
'//       energetic
'//       young
'//       dog
'//       was
'//       playing
'//       with
'//       his
'//       smaller
'//       more
'//       lethargic
'//       litter
'//       mate
