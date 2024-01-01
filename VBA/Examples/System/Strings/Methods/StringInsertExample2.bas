Attribute VB_Name = "StringInsertExample2"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 1, 2024
'@LastModified January 1, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.insert?view=netframework-4.8.1#examples

Option Explicit

''
' The following console application prompts the users to enter one or more
' adjectives to describe two animals. It then calls the Insert method to
' insert the text entered by the user into a string.
''
Public Sub StringInsertExample2()
    Dim animal1 As DotNetLib.String
    Set animal1 = Strings.Create("fox")
    Dim animal2 As DotNetLib.String
    Set animal2 = Strings.Create("dog")
    Dim strTarget As DotNetLib.String
    Set strTarget = Strings.Format("The {0} jumps over the {1}.", _
                                         animal1, animal2)
                                         
    Debug.Print VBAString.Format("The original string is:{0}{1}{0}", _
                         Environment.NewLine, strTarget)
                         
    Dim pvtInput As String
    pvtInput = InputBox(Strings.Format("Enter an adjective (or group of adjectives) " + _
                "to describe the {0}: ==> ", animal1))
    Dim adj1 As DotNetLib.String
    Set adj1 = Strings.Create(pvtInput)
    
    pvtInput = InputBox(Strings.Format("Enter an adjective (or group of adjectives) " + _
                "to describe the {0}: ==> ", animal2))
    Dim adj2 As DotNetLib.String
    Set adj2 = Strings.Create(pvtInput)
    
    Set adj1 = Strings.Concat2(adj1.Trim(), " ")
    Set adj2 = Strings.Concat2(adj2.Trim(), " ")

    Set strTarget = strTarget.Insert(strTarget.IndexOf(animal1), adj1)
    Set strTarget = strTarget.Insert(strTarget.IndexOf(animal2), adj2)

    Debug.Print VBAString.Format("{0}The final string is:{0}{1}", _
                          Environment.NewLine, strTarget)
End Sub

' Output from the example might appear as follows:
'       The original string is:
'       The fox jumps over the dog.
'
'       Enter an adjective (or group of adjectives) to describe the fox: ==> bold
'       Enter an adjective (or group of adjectives) to describe the dog: ==> lazy
'
'       The final string is:
'       The bold fox jumps over the lazy dog.
