VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StreamWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Implements a TextWriter for writing characters to a stream in a particular encoding."
'@Folder "VBADotNetLib.System.IO"
'@PredeclaredId
'@ModuleDescription("Implements a TextWriter for writing characters to a stream in a particular encoding.")
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 7, 2023
'@LastModified November 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8.1

'@Remarks
'Singleton class for StreamWriter containing factory methods and static members.

Option Explicit

Private Type TStreamWriterSingleton
   StreamWriterSingleton As DotNetLib.StreamWriterSingleton
End Type

Private this As TStreamWriterSingleton

Private Sub Class_Initialize()
   Set this.StreamWriterSingleton = New DotNetLib.StreamWriterSingleton
End Sub

Private Sub Class_Terminate()
   Set this.StreamWriterSingleton = Nothing
End Sub

'@Description("Initializes a new instance of the StreamWriter class for the specified file by using the default encoding and buffer size.")
Public Function Create(ByVal pPath As String) As DotNetLib.StreamWriter
Attribute Create.VB_Description = "Initializes a new instance of the StreamWriter class for the specified file by using the default encoding and buffer size."
   Set Create = this.StreamWriterSingleton.Create(pPath)
End Function

'@Description("Initializes a new instance of the StreamWriter class for the specified file by using the default encoding and buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file.")
Public Function Create2(ByVal pPath As String, ByVal pAppend As Boolean) As DotNetLib.StreamWriter
Attribute Create2.VB_Description = "Initializes a new instance of the StreamWriter class for the specified file by using the default encoding and buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file."
   Set Create2 = this.StreamWriterSingleton.Create_2(pPath, pAppend)
End Function

'@Description("Initializes a new instance of the StreamWriter class for the specified file by using the specified encoding and default buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file.")
Public Function Create3(ByVal pPath As String, ByVal pAppend As Boolean, ByVal pEncoding As DotNetLib.Encoding) As DotNetLib.StreamWriter
Attribute Create3.VB_Description = "Initializes a new instance of the StreamWriter class for the specified file by using the specified encoding and default buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file."
   Set Create3 = this.StreamWriterSingleton.Create_3(pPath, pAppend, pEncoding)
End Function

'@Description("Initializes a new instance of the StreamWriter class for the specified file on the specified path, using the specified encoding and buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file")
Public Function Create4(ByVal Path As String, ByVal pAppend As Boolean, ByVal pEncoding As DotNetLib.Encoding, ByVal bufferSize As Long) As DotNetLib.StreamWriter
Attribute Create4.VB_Description = "Initializes a new instance of the StreamWriter class for the specified file on the specified path, using the specified encoding and buffer size. If the file exists, it can be either overwritten or appended to. If the file does not exist, this constructor creates a new file"
   Set Create4 = this.StreamWriterSingleton.Create_4(Path, pAppend, pEncoding, bufferSize)
End Function

'@Description("Initializes a new instance of the StreamWriter class for the specified stream by using UTF-8 encoding and the default buffer size.")
Public Function Create5(ByVal pStream As mscorlib.Stream) As DotNetLib.StreamWriter
Attribute Create5.VB_Description = "Initializes a new instance of the StreamWriter class for the specified stream by using UTF-8 encoding and the default buffer size."
   Set Create5 = this.StreamWriterSingleton.Create_5(pStream)
End Function

'@Description("Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and the default buffer size.")
Public Function Create6(ByVal pStream As mscorlib.Stream, ByVal pAppend As DotNetLib.Encoding) As DotNetLib.StreamWriter
Attribute Create6.VB_Description = "Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and the default buffer size."
   Set Create6 = this.StreamWriterSingleton.Create_6(pStream, pAppend)
End Function

'@Description("Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and buffer size.")
Public Function Create7(ByVal pStream As mscorlib.Stream, ByVal pEncoding As DotNetLib.Encoding, ByVal bufferSize As Long) As DotNetLib.StreamWriter
Attribute Create7.VB_Description = "Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and buffer size."
   Set Create7 = this.StreamWriterSingleton.Create_7(pStream, pEncoding, bufferSize)
End Function

'@Description("Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and buffer size, and optionally leaves the stream open.")
Public Function Create8(ByVal pStream As mscorlib.Stream, ByVal pEncoding As DotNetLib.Encoding, ByVal bufferSize As Long, ByVal leaveOpen As Boolean) As DotNetLib.StreamWriter
Attribute Create8.VB_Description = "Initializes a new instance of the StreamWriter class for the specified stream by using the specified encoding and buffer size, and optionally leaves the stream open."
   Set Create8 = this.StreamWriterSingleton.Create_8(pStream, pEncoding, bufferSize, leaveOpen)
End Function

'@Description("Provides a StreamWriter with no backing store that can be written to, but not read from.")
Public Property Get NullStreamWriter() As DotNetLib.StreamWriter
Attribute NullStreamWriter.VB_Description = "Provides a StreamWriter with no backing store that can be written to, but not read from."
   Set NullStreamWriter = this.StreamWriterSingleton.NullStreamWriter()
End Property
