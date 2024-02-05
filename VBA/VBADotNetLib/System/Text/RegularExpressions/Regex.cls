VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an immutable regular expression."
'@IgnoreModule ParameterCanBeByVal
'@Folder "VBADotNetLib.System.Text.RegularExpressions"
'@PredeclaredId
'@ModuleDescription("Represents an immutable regular expression.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 25, 2023
'@LastModified February 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1
' https://learn.microsoft.com/en-us/dotnet/standard/base-types/regular-expressions

'@Remarks
' Singleton class for Regex containing constructors and static members.

Option Explicit

Private Type TRegexSingleton
   Regex As DotNetLib.RegexSingleton
End Type

Private this As TRegexSingleton

Private Sub Class_Initialize()
   Set this.Regex = New DotNetLib.RegexSingleton
End Sub

Private Sub Class_Terminate()
   Set this.Regex = Nothing
End Sub

'@Description("Initializes a new instance of the Regex class for the specified regular expression, with options that modify the pattern.")
Public Function Create(ByRef pattern As String, Optional ByVal options As DotNetLib.RegexOptions = RegexOptions_None) As DotNetLib.Regex
Attribute Create.VB_Description = "Initializes a new instance of the Regex class for the specified regular expression, with options that modify the pattern."
   Set Create = this.Regex.Create(pattern, options)
End Function

'@Description("Initializes a new instance of the Regex class for the specified regular expression, with options that modify the pattern and a value that specifies how long a pattern matching method should attempt a match before it times out.")
Public Function Create2(ByRef pattern As String, ByVal options As DotNetLib.RegexOptions, ByVal pMatchTimeout As DotNetLib.TimeSpan) As DotNetLib.Regex
Attribute Create2.VB_Description = "Initializes a new instance of the Regex class for the specified regular expression, with options that modify the pattern and a value that specifies how long a pattern matching method should attempt a match before it times out."
   Set Create2 = this.Regex.Create2(pattern, options, pMatchTimeout)
End Function

'@Description("Gets or sets the maximum number of entries in the current static cache of compiled regular expressions.")
Public Property Get CacheSize() As Long
Attribute CacheSize.VB_Description = "Gets or sets the maximum number of entries in the current static cache of compiled regular expressions."
   CacheSize = this.Regex.CacheSize()
End Property

'@Description("Gets or sets the maximum number of entries in the current static cache of compiled regular expressions.")
Public Property Let CacheSize(ByVal value As Long)
Attribute CacheSize.VB_Description = "Gets or sets the maximum number of entries in the current static cache of compiled regular expressions."
   this.Regex.CacheSize = value
End Property

'@Description("Specifies that a pattern-matching operation should not time out.")
Public Property Get InfiniteMatchTimeout() As DotNetLib.TimeSpan
Attribute InfiniteMatchTimeout.VB_Description = "Specifies that a pattern-matching operation should not time out."
   Set InfiniteMatchTimeout = this.Regex.InfiniteMatchTimeout()
End Property

''
'@Description("Escapes a minimal set of characters by replacing them with their escape codes. This instructs the regular expression engine to interpret these characters literally rather than as metacharacters.")
'@Parameters
'   str String
'       The input string that contains the text to convert.
'
'@Returns String
'   A string of characters with metacharacters converted to their escaped form.
'
'@Exceptions
'   ArgumentNullException
'       str is null.
''
Public Function Escape(ByRef str As String) As String
Attribute Escape.VB_Description = "Escapes a minimal set of characters by replacing them with their escape codes. This instructs the regular expression engine to interpret these characters literally rather than as metacharacters."
   Escape = this.Regex.Escape(str)
End Function

'@Description("Indicates whether the specified regular expression finds a match in the specified input string, using the specified matching options.")
Public Function IsMatch(ByRef strInput As String, ByRef pattern As String, Optional ByVal options As DotNetLib.RegexOptions = RegexOptions_None) As Boolean
Attribute IsMatch.VB_Description = "Indicates whether the specified regular expression finds a match in the specified input string, using the specified matching options."
   IsMatch = this.Regex.IsMatch(strInput, pattern, options)
End Function

'@Description("Indicates whether the specified regular expression finds a match in the specified input string, using the specified matching options and time-out interval.")
Public Function IsMatch2(ByRef strInput As String, ByRef pattern As String, ByVal options As DotNetLib.RegexOptions, ByVal pMatchTimeout As DotNetLib.TimeSpan) As Boolean
Attribute IsMatch2.VB_Description = "Indicates whether the specified regular expression finds a match in the specified input string, using the specified matching options and time-out interval."
   IsMatch2 = this.Regex.IsMatch2(strInput, pattern, options, pMatchTimeout)
End Function

'@Description("Searches the input string for the first occurrence of the specified regular expression, using the specified matching options.")
Public Function Match(ByRef strInput As String, ByRef pattern As String, Optional ByVal options As DotNetLib.RegexOptions = RegexOptions_None) As DotNetLib.Match
Attribute Match.VB_Description = "Searches the input string for the first occurrence of the specified regular expression, using the specified matching options."
   Set Match = this.Regex.Match(strInput, pattern, options)
End Function

'@Description("Searches the input string for the first occurrence of the specified regular expression, using the specified matching options and time-out interval")
Public Function Match2(ByRef strInput As String, ByRef pattern As String, ByVal options As DotNetLib.RegexOptions, ByVal pMatchTimeout As DotNetLib.TimeSpan) As DotNetLib.Match
Attribute Match2.VB_Description = "Searches the input string for the first occurrence of the specified regular expression, using the specified matching options and time-out interval"
   Set Match2 = this.Regex.Match2(strInput, pattern, options, pMatchTimeout)
End Function

'@Description("Searches the specified input string for all occurrences of a specified regular expression, using the specified matching options.")
Public Function Matches(ByRef strInput As String, ByRef pattern As String, Optional ByVal options As DotNetLib.RegexOptions = RegexOptions_None) As DotNetLib.MatchCollection
Attribute Matches.VB_Description = "Searches the specified input string for all occurrences of a specified regular expression, using the specified matching options."
   Set Matches = this.Regex.Matches(strInput, pattern, options)
End Function

'@Description("Searches the specified input string for all occurrences of a specified regular expression, using the specified matching options and time-out interval.")
Public Function Matches2(ByRef strInput As String, ByRef pattern As String, ByVal options As DotNetLib.RegexOptions, ByVal pMatchTimeout As DotNetLib.TimeSpan) As DotNetLib.MatchCollection
Attribute Matches2.VB_Description = "Searches the specified input string for all occurrences of a specified regular expression, using the specified matching options and time-out interval."
   Set Matches2 = this.Regex.Matches2(strInput, pattern, options, pMatchTimeout)
End Function

'@Description("In a specified input string, replaces all strings that match a specified regular expression with a specified replacement string. Specified options modify the matching operation.")
Public Function Replace(ByRef strInput As String, ByRef pattern As String, ByVal replacement As String, Optional ByVal options As DotNetLib.RegexOptions = RegexOptions_None) As String
Attribute Replace.VB_Description = "In a specified input string, replaces all strings that match a specified regular expression with a specified replacement string. Specified options modify the matching operation."
   Replace = this.Regex.Replace(strInput, pattern, replacement, options)
End Function

'@Description("In a specified input string, replaces all strings that match a specified regular expression with a specified replacement string. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found.")
Public Function Replace2(ByRef strInput As String, ByRef pattern As String, ByVal replacement As String, ByVal options As DotNetLib.RegexOptions, ByVal pMatchTimeout As DotNetLib.TimeSpan) As String
Attribute Replace2.VB_Description = "In a specified input string, replaces all strings that match a specified regular expression with a specified replacement string. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found."
   Replace2 = this.Regex.Replace2(strInput, pattern, replacement, options, pMatchTimeout)
End Function

'@Description("In a specified input string, replaces all strings that match a specified regular expression with a string returned by a MatchEvaluator delegate. Specified options modify the matching operation.")
Public Function Replace3(ByRef strInput As String, ByRef pattern As String, ByVal evaluator As DotNetLib.MatchEvaluator, Optional ByVal options As DotNetLib.RegexOptions = RegexOptions_None) As String
Attribute Replace3.VB_Description = "In a specified input string, replaces all strings that match a specified regular expression with a string returned by a MatchEvaluator delegate. Specified options modify the matching operation."
   Replace3 = this.Regex.Replace3(strInput, pattern, evaluator, options)
End Function

'@Description("In a specified input string, replaces all substrings that match a specified regular expression with a string returned by a MatchEvaluator delegate. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found.")
Public Function Replace4(ByRef strInput As String, ByRef pattern As String, ByVal evaluator As DotNetLib.MatchEvaluator, ByVal options As DotNetLib.RegexOptions, ByVal pMatchTimeout As DotNetLib.TimeSpan) As String
Attribute Replace4.VB_Description = "In a specified input string, replaces all substrings that match a specified regular expression with a string returned by a MatchEvaluator delegate. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found."
   Replace4 = this.Regex.Replace4(strInput, pattern, evaluator, options, pMatchTimeout)
End Function

'@Description("Splits an input string into an array of substrings at the positions defined by a specified regular expression pattern. Specified options modify the matching operation.")
Public Function Split(ByRef strInput As String, ByRef pattern As String, Optional ByVal options As DotNetLib.RegexOptions = RegexOptions_None) As String()
Attribute Split.VB_Description = "Splits an input string into an array of substrings at the positions defined by a specified regular expression pattern. Specified options modify the matching operation."
   Split = this.Regex.Split(strInput, pattern, options)
End Function

'@Description("Splits an input string into an array of substrings at the positions defined by a specified regular expression pattern. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found.")
Public Function Split2(ByRef strInput As String, ByRef pattern As String, ByVal options As DotNetLib.RegexOptions, ByVal pMatchTimeout As DotNetLib.TimeSpan) As String()
Attribute Split2.VB_Description = "Splits an input string into an array of substrings at the positions defined by a specified regular expression pattern. Additional parameters specify options that modify the matching operation and a time-out interval if no match is found."
   Split2 = this.Regex.Split2(strInput, pattern, options, pMatchTimeout)
End Function

''
'@Description("Converts any escaped characters in the input string.")
'@Parameters
'   str String
'       The input string containing the text to convert.
'
'@Returns String
'   A string of characters with any escaped characters converted to their unescaped form.
'
'@Exceptions
'   ArgumentException
'       str includes an unrecognized escape sequence.
'
'   ArgumentNullException
'       str is null.
''
Public Function Unescape(ByRef str As String) As String
Attribute Unescape.VB_Description = "Converts any escaped characters in the input string."
   Unescape = this.Regex.Unescape(str)
End Function
