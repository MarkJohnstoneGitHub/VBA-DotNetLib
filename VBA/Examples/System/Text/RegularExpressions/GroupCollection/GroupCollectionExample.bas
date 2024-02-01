Attribute VB_Name = "GroupCollectionExample"
'@Folder("Examples.System.Text.RegularExpressions.GroupCollection")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 1, 2024
'@LastModified February 1, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.groupcollection?view=netframework-4.8.1#examples

Option Explicit

''
' The following example uses a regular expression with capturing groups to
' extract information about trademarks and registered trademarks used in text.
' The regular expression pattern is \b(\w+?)([\u00AE\u2122]), which is
' interpreted as shown in the following table.
' Pattern           Description
' \b                Look for a word boundary.
' (\w+?)            Look for one or more word characters. Together, these form the trademarked name.
'                   (Note that this regular expression assumes that a trademark consists of a single word.)
'                   This is the first capturing group.
' ([\u00AE\u2122])  Look for either the ® or the ™ character. This is the second capturing group.
'
' For each match, the GroupCollection contains three Group objects. The first
' object contains the string that matches the entire regular expression. The
' second object, which represents the first captured group, contains the
' product name. The third object, which represents the second captured group,
' contains the trademark or registered trademark symbol.
''
Public Sub GroupCollectionExample()
    Dim pattern As String
    pattern = "\b(\w+?)([\u00AE\u2122])"
    Dim pvtInput As String
    pvtInput = "Microsoft® Office Professional Edition combines several office " + _
                "productivity products, including Word, Excel®, Access®, Outlook®, " + _
                "PowerPoint®, and several others. Some guidelines for creating " + _
                "corporate documents using these productivity tools are available " + _
                "from the documents created using Silverlight™ on the corporate " + _
                "intranet site."
                
    Dim pvtMatches As DotNetLib.MatchCollection
    Set pvtMatches = Regex.Matches(pvtInput, pattern)
    Dim pvtMatch As Variant
    For Each pvtMatch In pvtMatches
        Dim pvtGroups As DotNetLib.GroupCollection
        Set pvtGroups = pvtMatch.Groups
        Debug.Print VBString.Format("{0}: {1}", pvtGroups(2), pvtGroups(1))
    Next
End Sub

' The example displays the following output:
'       ®: Microsoft
'       ®: Excel
'       ®: Access
'       ®: Outlook
'       ®: PowerPoint
'       ™: Silverlight
