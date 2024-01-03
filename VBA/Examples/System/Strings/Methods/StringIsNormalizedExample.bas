Attribute VB_Name = "StringIsNormalizedExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 1, 2024
'@LastModified January 1, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.isnormalized?view=netframework-4.8.1#examples

Option Explicit

''
' The following example determines whether a string is successfully normalized
' to various normalization forms.
''
Public Sub StringIsNormalizedExample()
    ' Character c; combining characters acute and cedilla; character 3/4
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("\u0063\u0301\u0327\u00BE")
    Dim s2 As DotNetLib.String
    Set s2 = Nothing
    Dim divider As DotNetLib.String
    Set divider = Strings.Create2("-", 80)
    Set divider = Strings.Concat3(Environment.NewLine, divider, Environment.NewLine)
    
    'Call show("s1", s1)
    
    Debug.Print
    Debug.Print "U+0063 = LATIN SMALL LETTER C"
    Debug.Print "U+0301 = COMBINING ACUTE ACCENT"
    Debug.Print "U+0327 = COMBINING CEDILLA"
    Debug.Print "U+00BE = VULGAR FRACTION THREE QUARTERS"
    Debug.Print divider.ToString()
    
    Debug.Print VBAString.Format("A1) Is s1 normalized to the default form (Form C)?: {0}", _
                                    s1.IsNormalized())

    Debug.Print VBAString.Format("A2) Is s1 normalized to Form C?:  {0}", _
                                    s1.IsNormalized(NormalizationForm.NormalizationForm_FormC))
    Debug.Print VBAString.Format("A3) Is s1 normalized to Form D?:  {0}", _
                                    s1.IsNormalized(NormalizationForm.NormalizationForm_FormD))
    Debug.Print VBAString.Format("A4) Is s1 normalized to Form KC?: {0}", _
                                    s1.IsNormalized(NormalizationForm.NormalizationForm_FormKC))
    Debug.Print VBAString.Format("A5) Is s1 normalized to Form KD?: {0}", _
                                    s1.IsNormalized(NormalizationForm.NormalizationForm_FormKD))
                                    
    Debug.Print divider.ToString()

    Debug.Print ("Set string s2 to each normalized form of string s1.");
    Debug.Print
    Debug.Print "U+1E09 = LATIN SMALL LETTER C WITH CEDILLA AND ACUTE"
    Debug.Print "U+0033 = DIGIT THREE"
    Debug.Print "U+2044 = FRACTION SLASH"
    Debug.Print "U+0034 = DIGIT FOUR"
    Debug.Print divider.ToString()

    Set s2 = s1.Normalize()
    Debug.Print "B1) Is s2 normalized to the default form (Form C)?: ";
    Debug.Print (s2.IsNormalized())
    'Show("s2", s2);
    Debug.Print

    Set s2 = s1.Normalize(NormalizationForm.NormalizationForm_FormC)
    Debug.Print "B2) Is s2 normalized to Form C?: ";
    Debug.Print s2.IsNormalized(NormalizationForm.NormalizationForm_FormC)
    'Show("s2", s2);
    Debug.Print

    Set s2 = s1.Normalize(NormalizationForm.NormalizationForm_FormD)
    Debug.Print "B3) Is s2 normalized to Form D?: ";
    Debug.Print s2.IsNormalized(NormalizationForm.NormalizationForm_FormD)
    'Show("s2", s2);
    Debug.Print

    Set s2 = s1.Normalize(NormalizationForm.NormalizationForm_FormKC)
    Debug.Print "B4) Is s2 normalized to Form KC?: ";
    Debug.Print s2.IsNormalized(NormalizationForm.NormalizationForm_FormKC)
    'Show("s2", s2);
    Debug.Print

    Set s2 = s1.Normalize(NormalizationForm.NormalizationForm_FormKD)
    Debug.Print "B5) Is s2 normalized to Form KD?: ";
    Debug.Print s2.IsNormalized(NormalizationForm.NormalizationForm_FormKD)
    'Show("s2", s2);
    Debug.Print
End Sub

'@TODO String enumeration of characters
Private Sub show(ByVal title As String, ByVal s As DotNetLib.String)
'    {
'       Console.Write("Characters in string {0} = ", title);
'       foreach(short x in s) {
'           Console.Write("{0:X4} ", x);
'       }
'       Console.WriteLine();
'    }
End Sub

'/*
'This example produces the following results:
'
'Characters in string s1 = 0063 0301 0327 00BE
'
'U+0063 = LATIN SMALL LETTER C
'U+0301 = COMBINING ACUTE ACCENT
'U+0327 = COMBINING CEDILLA
'U+00BE = VULGAR FRACTION THREE QUARTERS
'
'--------------------------------------------------------------------------------
'
'A1) Is s1 normalized to the default form (Form C)?: False
'A2) Is s1 normalized to Form C?:  False
'A3) Is s1 normalized to Form D?:  False
'A4) Is s1 normalized to Form KC?: False
'A5) Is s1 normalized to Form KD?: False
'
'--------------------------------------------------------------------------------
'
'Set string s2 to each normalized form of string s1.
'
'U+1E09 = LATIN SMALL LETTER C WITH CEDILLA AND ACUTE
'U+0033 = DIGIT THREE
'U+2044 = FRACTION SLASH
'U+0034 = DIGIT FOUR
'
'--------------------------------------------------------------------------------
'
'B1) Is s2 normalized to the default form (Form C)?: True
'Characters in string s2 = 1E09 00BE
'
'B2) Is s2 normalized to Form C?: True
'Characters in string s2 = 1E09 00BE
'
'B3) Is s2 normalized to Form D?: True
'Characters in string s2 = 0063 0327 0301 00BE
'
'B4) Is s2 normalized to Form KC?: True
'Characters in string s2 = 1E09 0033 2044 0034
'
'B5) Is s2 normalized to Form KD?: True
'Characters in string s2 = 0063 0327 0301 0033 2044 0034
'
'*/
