VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Uri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides an object representation of a uniform resource identifier (URI) and easy access to the parts of the URI."
'@IgnoreModule ParameterCanBeByVal
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("Provides an object representation of a uniform resource identifier (URI) and easy access to the parts of the URI.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 21, 2024
'@LastModified January 23, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri?view=netframework-4.8.1

Option Explicit

Private Type TUriSingleton
   UriSingleton As DotNetLib.UriSingleton
End Type

Private this As TUriSingleton

Private Sub Class_Initialize()
   Set this.UriSingleton = New DotNetLib.UriSingleton
End Sub

Private Sub Class_Terminate()
   Set this.UriSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the Uri class with the specified URI.")
Public Function Create(ByRef uriString As String) As DotNetLib.Uri
Attribute Create.VB_Description = "Initializes a new instance of the Uri class with the specified URI."
   Set Create = this.UriSingleton.Create(uriString)
End Function

'@Description("Initializes a new instance of the Uri class based on the specified base URI and relative URI string.")
Public Function Create2(ByVal baseUri As DotNetLib.Uri, ByRef relativeUri As String) As DotNetLib.Uri
Attribute Create2.VB_Description = "Initializes a new instance of the Uri class based on the specified base URI and relative URI string."
   Set Create2 = this.UriSingleton.Create2(baseUri, relativeUri)
End Function

'@Description("Initializes a new instance of the Uri class based on the combination of a specified base Uri instance and a relative Uri instance.")
Public Function Create3(ByVal baseUri As DotNetLib.Uri, ByVal relativeUri As DotNetLib.Uri) As DotNetLib.Uri
Attribute Create3.VB_Description = "Initializes a new instance of the Uri class based on the combination of a specified base Uri instance and a relative Uri instance."
   Set Create3 = this.UriSingleton.Create3(baseUri, relativeUri)
End Function

'@Description("Initializes a new instance of the Uri class with the specified URI. This constructor allows you to specify if the URI string is a relative URI, absolute URI, or is indeterminate.")
Public Function Create4(ByRef uriString As String, ByVal UriKind As DotNetLib.UriKind) As DotNetLib.Uri
Attribute Create4.VB_Description = "Initializes a new instance of the Uri class with the specified URI. This constructor allows you to specify if the URI string is a relative URI, absolute URI, or is indeterminate."
   Set Create4 = this.UriSingleton.Create4(uriString, UriKind)
End Function

'@Description("Specifies the characters that separate the communication protocol scheme from the address portion of the URI. This field is read-only.")
Public Property Get SchemeDelimiter() As String
Attribute SchemeDelimiter.VB_Description = "Specifies the characters that separate the communication protocol scheme from the address portion of the URI. This field is read-only."
   SchemeDelimiter = this.UriSingleton.SchemeDelimiter()
End Property

'@Description("Specifies that the URI is a pointer to a file. This field is read-only.")
Public Property Get UriSchemeFile() As String
Attribute UriSchemeFile.VB_Description = "Specifies that the URI is a pointer to a file. This field is read-only."
   UriSchemeFile = this.UriSingleton.UriSchemeFile()
End Property

'@Description("Specifies that the URI is accessed through the File Transfer Protocol (FTP). This field is read-only.")
Public Property Get UriSchemeFtp() As String
Attribute UriSchemeFtp.VB_Description = "Specifies that the URI is accessed through the File Transfer Protocol (FTP). This field is read-only."
   UriSchemeFtp = this.UriSingleton.UriSchemeFtp()
End Property

'@Description("Specifies that the URI is accessed through the Gopher protocol. This field is read-only.")
Public Property Get UriSchemeGopher() As String
Attribute UriSchemeGopher.VB_Description = "Specifies that the URI is accessed through the Gopher protocol. This field is read-only."
   UriSchemeGopher = this.UriSingleton.UriSchemeGopher()
End Property

'@Description("Specifies that the URI is accessed through the Hypertext Transfer Protocol (HTTP). This field is read-only.")
Public Property Get UriSchemeHttp() As String
Attribute UriSchemeHttp.VB_Description = "Specifies that the URI is accessed through the Hypertext Transfer Protocol (HTTP). This field is read-only."
   UriSchemeHttp = this.UriSingleton.UriSchemeHttp()
End Property

'@Description("Specifies that the URI is accessed through the Secure Hypertext Transfer Protocol (HTTPS). This field is read-only.")
Public Property Get UriSchemeHttps() As String
Attribute UriSchemeHttps.VB_Description = "Specifies that the URI is accessed through the Secure Hypertext Transfer Protocol (HTTPS). This field is read-only."
   UriSchemeHttps = this.UriSingleton.UriSchemeHttps()
End Property

'@Description("Specifies that the URI is an email address and is accessed through the Simple Mail Transport Protocol (SMTP). This field is read-only.")
Public Property Get UriSchemeMailto() As String
Attribute UriSchemeMailto.VB_Description = "Specifies that the URI is an email address and is accessed through the Simple Mail Transport Protocol (SMTP). This field is read-only."
   UriSchemeMailto = this.UriSingleton.UriSchemeMailto()
End Property

'@Description("Specifies that the URI is accessed through the NetPipe scheme used by Windows Communication Foundation (WCF). This field is read-only.")
Public Property Get UriSchemeNetPipe() As String
Attribute UriSchemeNetPipe.VB_Description = "Specifies that the URI is accessed through the NetPipe scheme used by Windows Communication Foundation (WCF). This field is read-only."
   UriSchemeNetPipe = this.UriSingleton.UriSchemeNetPipe()
End Property

'@Description("Specifies that the URI is accessed through the NetTcp scheme used by Windows Communication Foundation (WCF). This field is read-only.")
Public Property Get UriSchemeNetTcp() As String
Attribute UriSchemeNetTcp.VB_Description = "Specifies that the URI is accessed through the NetTcp scheme used by Windows Communication Foundation (WCF). This field is read-only."
   UriSchemeNetTcp = this.UriSingleton.UriSchemeNetTcp()
End Property

'@Description("Specifies that the URI is an Internet news group and is accessed through the Network News Transport Protocol (NNTP). This field is read-only.")
Public Property Get UriSchemeNews() As String
Attribute UriSchemeNews.VB_Description = "Specifies that the URI is an Internet news group and is accessed through the Network News Transport Protocol (NNTP). This field is read-only."
   UriSchemeNews = this.UriSingleton.UriSchemeNews()
End Property

'@Description("Specifies that the URI is an Internet news group and is accessed through the Network News Transport Protocol (NNTP). This field is read-only.")
Public Property Get UriSchemeNntp() As String
Attribute UriSchemeNntp.VB_Description = "Specifies that the URI is an Internet news group and is accessed through the Network News Transport Protocol (NNTP). This field is read-only."
   UriSchemeNntp = this.UriSingleton.UriSchemeNntp()
End Property

'@Description("Determines whether the specified host name is a valid DNS name.")
Public Function CheckHostName(ByRef pName As String) As DotNetLib.UriHostNameType
Attribute CheckHostName.VB_Description = "Determines whether the specified host name is a valid DNS name."
   CheckHostName = this.UriSingleton.CheckHostName(pName)
End Function

'@Description("Determines whether the specified scheme name is valid.")
Public Function CheckSchemeName(ByRef schemeName As String) As Boolean
Attribute CheckSchemeName.VB_Description = "Determines whether the specified scheme name is valid."
   CheckSchemeName = this.UriSingleton.CheckSchemeName(schemeName)
End Function

'@Description("Compares the specified parts of two URIs using the specified comparison rules.")
Public Function Compare(ByVal uri1 As DotNetLib.Uri, ByRef uri2 As DotNetLib.Uri, ByVal partsToCompare As DotNetLib.UriComponents, ByVal compareFormat As DotNetLib.UriFormat, ByVal comparisonType As StringComparison) As Long
Attribute Compare.VB_Description = "Compares the specified parts of two URIs using the specified comparison rules."
   Compare = this.UriSingleton.Compare(uri1, uri2, partsToCompare, compareFormat, comparisonType)
End Function

'@Description("Converts a string to its escaped representation.")
Public Function EscapeDataString(ByRef stringToEscape As String) As String
Attribute EscapeDataString.VB_Description = "Converts a string to its escaped representation."
   EscapeDataString = this.UriSingleton.EscapeDataString(stringToEscape)
End Function

'@Description("Converts a URI string to its escaped representation.")
Public Function EscapeUriString(ByRef stringToEscape As String) As String
Attribute EscapeUriString.VB_Description = "Converts a URI string to its escaped representation."
   EscapeUriString = this.UriSingleton.EscapeUriString(stringToEscape)
End Function

'@Description("Gets the decimal value of a hexadecimal digit.")
Public Function FromHex(ByRef digit As String) As Long
Attribute FromHex.VB_Description = "Gets the decimal value of a hexadecimal digit."
   FromHex = this.UriSingleton.FromHex(digit)
End Function

'@Description("Converts a specified character into its hexadecimal equivalent.")
Public Function HexEscape(ByRef character As String) As String
Attribute HexEscape.VB_Description = "Converts a specified character into its hexadecimal equivalent."
   HexEscape = this.UriSingleton.HexEscape(character)
End Function

'@Description("Converts a specified hexadecimal representation of a character to the character")
Public Function HexUnescape(ByRef pattern As String, ByRef index As Long) As String
Attribute HexUnescape.VB_Description = "Converts a specified hexadecimal representation of a character to the character"
   HexUnescape = this.UriSingleton.HexUnescape(pattern, index)
End Function

'@Description("Determines whether a specified character is a valid hexadecimal digit.")
Public Function IsHexDigit(ByRef character As String) As Boolean
Attribute IsHexDigit.VB_Description = "Determines whether a specified character is a valid hexadecimal digit."
   IsHexDigit = this.UriSingleton.IsHexDigit(character)
End Function

'@Description("Determines whether a character in a string is hexadecimal encoded.")
Public Function IsHexEncoding(ByRef pattern As String, ByVal index As Long) As Boolean
Attribute IsHexEncoding.VB_Description = "Determines whether a character in a string is hexadecimal encoded."
   IsHexEncoding = this.UriSingleton.IsHexEncoding(pattern, index)
End Function

'@Description("Indicates whether the string is well-formed by attempting to construct a URI with the string and ensures that the string does not require further escaping.")
Public Function IsWellFormedUriString(ByRef uriString As String, ByVal UriKind As DotNetLib.UriKind) As Boolean
Attribute IsWellFormedUriString.VB_Description = "Indicates whether the string is well-formed by attempting to construct a URI with the string and ensures that the string does not require further escaping."
   IsWellFormedUriString = this.UriSingleton.IsWellFormedUriString(uriString, UriKind)
End Function

'@Description("Creates a new Uri using the specified base and relative Uri instances.")
Public Function TryCreate(ByVal baseUri As DotNetLib.Uri, ByVal relativeUri As DotNetLib.Uri, ByRef result As DotNetLib.Uri) As Boolean
Attribute TryCreate.VB_Description = "Creates a new Uri using the specified base and relative Uri instances."
   TryCreate = this.UriSingleton.TryCreate(baseUri, relativeUri, result)
End Function

'@Description("Creates a new Uri using the specified String instance and a UriKind.")
Public Function TryCreate3(ByVal uriString As String, ByVal UriKind As DotNetLib.UriKind, ByRef result As DotNetLib.Uri) As Boolean
Attribute TryCreate3.VB_Description = "Creates a new Uri using the specified String instance and a UriKind."
   TryCreate3 = this.UriSingleton.TryCreate3(uriString, UriKind, result)
End Function

'@Description("Creates a new Uri using the specified base and relative String instances.")
Public Function TryCreate2(ByVal baseUri As DotNetLib.Uri, ByVal relativeUri As String, ByRef result As DotNetLib.Uri) As Boolean
Attribute TryCreate2.VB_Description = "Creates a new Uri using the specified base and relative String instances."
   TryCreate2 = this.UriSingleton.TryCreate2(baseUri, relativeUri, result)
End Function

'@Description("Converts a string to its unescaped representation.")
Public Function UnescapeDataString(ByRef stringToUnescape As String) As String
Attribute UnescapeDataString.VB_Description = "Converts a string to its unescaped representation."
   UnescapeDataString = this.UriSingleton.UnescapeDataString(stringToUnescape)
End Function

'@Description("Determines whether two Uri instances have the same value.")
Public Function Equality(ByVal uri1 As DotNetLib.Uri, ByVal uri2 As DotNetLib.Uri) As Boolean
Attribute Equality.VB_Description = "Determines whether two Uri instances have the same value."
   Equality = this.UriSingleton.Equality(uri1, uri2)
End Function

'@Description("Determines whether two Uri instances do not have the same value.")
Public Function Inequality(ByVal uri1 As DotNetLib.Uri, ByVal uri2 As DotNetLib.Uri) As Boolean
Attribute Inequality.VB_Description = "Determines whether two Uri instances do not have the same value."
   Inequality = this.UriSingleton.Inequality(uri1, uri2)
End Function
