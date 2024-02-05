Attribute VB_Name = "GroupCollectionItemExample"
'@IgnoreModule IndexedDefaultMemberAccess
'@Folder("Examples.System.Text.RegularExpressions.GroupCollection.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 2, 2024
'@LastModified February 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.groupcollection.item?view=netframework-4.8.1#system-text-regularexpressions-groupcollection-item(system-int32)

Option Explicit

''
' The following example defines a regular expression that consists of two
' numbered groups. The first group captures one or more consecutive digits.
' The second group matches a single character. Because the regular expression
' engine looks for zero or one occurrence of the first group, it does not
' always find a match even if the regular expression match is successful.
' The example then illustrates the result when the Item[Int32] property is used
' to retrieve an unmatched group, a matched group, and a group that is not
' defined in the regular expression. The example defines a regular expression
' pattern (\d+)*(\w)\2, which is interpreted as shown in the following table.
'
'
' pattern Description
' (\d+)*  Match one or more occurrence of a decimal digit.
'         This is the first capturing group.
'         Match this pattern either zero or one time.
' (\w)    This is the second capturing group.
' \k      Match the string captured by the second capturing group.
''
Public Sub GroupCollectionItemExample()
    Dim pattern As String
    pattern = "(\d+)*(\w)\2"
    Dim pvtInput As String
    pvtInput = "AA"
    Dim pvtMatch As DotNetLib.Match
    Set pvtMatch = Regex.Match(pvtInput, pattern)
    
    ' Get the first named group.
    Dim group1  As DotNetLib.Group
    Set group1 = pvtMatch.Groups(1)
    Debug.Print VBString.Format("Group 1 value: {0}", IIf(group1.Success, group1.value, "Empty"))

    ' Get the second named group.
    Dim group2  As DotNetLib.Group
    Set group2 = pvtMatch.Groups(2)
    Debug.Print VBString.Format("Group 2 value: {0}", IIf(group2.Success, group2.value, "Empty"))
    
    ' Get a non-existent group.
    Dim group3  As DotNetLib.Group
    Set group3 = pvtMatch.Groups(3)
    Debug.Print VBString.Format("Group 3 value: {0}", IIf(group3.Success, group3.value, "Empty"))
End Sub

' The example displays the following output:
'       Group 1 value: Empty
'       Group 2 value: A
'       Group 3 value: Empty

