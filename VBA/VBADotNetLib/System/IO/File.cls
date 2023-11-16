VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides static methods for the creation, copying, deletion, moving, and opening of a single file, and aids in the creation of FileStream objects."
'@ModuleDescription("Provides static methods for the creation, copying, deletion, moving, and opening of a single file, and aids in the creation of FileStream objects.")
'@Folder "VBADotNetLib.System.IO"
'@PredeclaredId
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 6, 2023
'@LastModified November 14, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.file?view=netframework-4.8.1

'@Remarks
'Singleton class for File containing factory methods and static members.

Option Explicit

Private Type TFileSingleton
   FileSingleton As DotNetLib.FileSingleton
End Type

Private this As TFileSingleton

Private Sub Class_Initialize()
   Set this.FileSingleton = New DotNetLib.FileSingleton
End Sub

Private Sub Class_Terminate()
   Set this.FileSingleton = Nothing
End Sub

'@Description("Appends lines to a file, and then closes the file. If the specified file does not exist, this method creates a file, writes the specified lines to the file, and then closes the file.")
Public Sub AppendAllLines(ByVal pPath As String, ByVal contents As mscorlib.IEnumerable)
Attribute AppendAllLines.VB_Description = "Appends lines to a file, and then closes the file. If the specified file does not exist, this method creates a file, writes the specified lines to the file, and then closes the file."
    Call this.FileSingleton.AppendAllLines(pPath, contents)
End Sub

'@Description("Appends lines to a file by using a specified encoding, and then closes the file. If the specified file does not exist, this method creates a file, writes the specified lines to the file, and then closes the file.")
Public Sub AppendAllLines2(ByVal pPath As String, ByVal contents As IEnumerable, ByVal pEncoding As mscorlib.Encoding)
Attribute AppendAllLines2.VB_Description = "Appends lines to a file by using a specified encoding, and then closes the file. If the specified file does not exist, this method creates a file, writes the specified lines to the file, and then closes the file."
    Call this.FileSingleton.AppendAllLines_2(pPath, contents, pEncoding)
End Sub

'@Description("Opens a file, appends the specified string to the file, and then closes the file. If the file does not exist, this method creates a file, writes the specified string to the file, then closes the file.")
Public Sub AppendAllText(ByVal pPath As String, ByVal contents As String)
Attribute AppendAllText.VB_Description = "Opens a file, appends the specified string to the file, and then closes the file. If the file does not exist, this method creates a file, writes the specified string to the file, then closes the file."
    Call this.FileSingleton.AppendAllText(pPath, contents)
End Sub

'@Description("Appends the specified string to the file using the specified encoding, creating the file if it does not already exist.")
Public Sub AppendAllText2(ByVal pPath As String, ByVal contents As String, ByVal pEncoding As mscorlib.Encoding)
Attribute AppendAllText2.VB_Description = "Appends the specified string to the file using the specified encoding, creating the file if it does not already exist."
    Call this.FileSingleton.AppendAllText_2(pPath, contents, pEncoding)
End Sub

'@Description("Creates a StreamWriter that appends UTF-8 encoded text to an existing file, or to a new file if the specified file does not exist.")
Public Function AppendText(ByVal pPath As String) As DotNetLib.StreamWriter
Attribute AppendText.VB_Description = "Creates a StreamWriter that appends UTF-8 encoded text to an existing file, or to a new file if the specified file does not exist."
   Set AppendText = this.FileSingleton.AppendText(pPath)
End Function

'@Description("Copies an existing file to a new file. Overwriting a file of the same name is not allowed.")
Public Sub Copy(ByVal sourceFileName As String, ByVal destFileName As String)
Attribute Copy.VB_Description = "Copies an existing file to a new file. Overwriting a file of the same name is not allowed."
 Call this.FileSingleton.Copy(sourceFileName, destFileName)
End Sub

'@Description("Copies an existing file to a new file. Overwriting a file of the same name is allowed.")
Public Sub Copy2(ByVal sourceFileName As String, ByVal destFileName As String, ByVal overwrite As Boolean)
Attribute Copy2.VB_Description = "Copies an existing file to a new file. Overwriting a file of the same name is allowed."
 Call this.FileSingleton.Copy_2(sourceFileName, destFileName, overwrite)
End Sub

'@Description("Creates or overwrites a file in the specified path.")
Public Function Create(ByVal pPath As String) As mscorlib.FileStream
Attribute Create.VB_Description = "Creates or overwrites a file in the specified path."
   Set Create = this.FileSingleton.Create(pPath)
End Function

'@Description("Creates or overwrites a file in the specified path, specifying a buffer size.")
Public Function Create2(ByVal pPath As String, ByVal bufferSize As Long) As mscorlib.FileStream
Attribute Create2.VB_Description = "Creates or overwrites a file in the specified path, specifying a buffer size."
   Set Create2 = this.FileSingleton.Create_2(pPath, bufferSize)
End Function

'@Description("Creates or overwrites a file in the specified path, specifying a buffer size and options that describe how to create or overwrite the file.")
Public Function Create3(ByVal pPath As String, ByVal bufferSize As Long, ByVal options As mscorlib.FileOptions) As mscorlib.FileStream
Attribute Create3.VB_Description = "Creates or overwrites a file in the specified path, specifying a buffer size and options that describe how to create or overwrite the file."
   Set Create3 = this.FileSingleton.Create_3(pPath, bufferSize, options)
End Function

'@Description("Creates or overwrites a file in the specified path, specifying a buffer size, options that describe how to create or overwrite the file, and a value that determines the access control and audit security for the file.")
Public Function Create4(ByVal pPath As String, ByVal bufferSize As Long, ByVal options As FileOptions, ByVal fileSecurity As Object) As mscorlib.FileStream
Attribute Create4.VB_Description = "Creates or overwrites a file in the specified path, specifying a buffer size, options that describe how to create or overwrite the file, and a value that determines the access control and audit security for the file."
   Set Create4 = this.FileSingleton.Create_4(pPath, bufferSize, options, fileSecurity)
End Function

'@Description("Creates or opens a file for writing UTF-8 encoded text. If the file already exists, its contents are overwritten.")
Public Function CreateText(ByVal pPath As String) As DotNetLib.StreamWriter
Attribute CreateText.VB_Description = "Creates or opens a file for writing UTF-8 encoded text. If the file already exists, its contents are overwritten."
   Set CreateText = this.FileSingleton.CreateText(pPath)
End Function

'@Description("Decrypts a file that was encrypted by the current account using the Encrypt(String) method.")
Public Sub Decrypt(ByVal pPath As String)
Attribute Decrypt.VB_Description = "Decrypts a file that was encrypted by the current account using the Encrypt(String) method."
    Call this.FileSingleton.Decrypt(pPath)
End Sub

'@Description("Deletes the specified file.")
Public Sub Delete(ByVal pPath As String)
Attribute Delete.VB_Description = "Deletes the specified file."
    Call this.FileSingleton.Delete(pPath)
End Sub

'@Description("Encrypts a file so that only the account used to encrypt the file can decrypt it.")
Public Sub Encrypt(ByVal pPath As String)
Attribute Encrypt.VB_Description = "Encrypts a file so that only the account used to encrypt the file can decrypt it."
    Call this.FileSingleton.Encrypt(pPath)
End Sub

'@Description("Determines whether the specified file exists.")
Public Function Exists(ByVal pPath As String) As Boolean
Attribute Exists.VB_Description = "Determines whether the specified file exists."
   Exists = this.FileSingleton.Exists(pPath)
End Function

'@Description("Gets a FileSecurity object that encapsulates the access control list (ACL) entries for a specified file.")
Public Function GetAccessControl(ByVal pPath As String) As Object
Attribute GetAccessControl.VB_Description = "Gets a FileSecurity object that encapsulates the access control list (ACL) entries for a specified file."
   Set GetAccessControl = this.FileSingleton.GetAccessControl(pPath)
End Function

'@Description("Gets a FileSecurity object that encapsulates the specified type of access control list (ACL) entries for a particular file.")
Public Function GetAccessControl2(ByVal pPath As String, ByVal includeSections As DotNetLib.AccessControlSections) As Object
Attribute GetAccessControl2.VB_Description = "Gets a FileSecurity object that encapsulates the specified type of access control list (ACL) entries for a particular file."
   Set GetAccessControl2 = this.FileSingleton.GetAccessControl_2(pPath, includeSections)
End Function

'@Description("Gets the FileAttributes of the file on the path.")
Public Function GetAttributes(ByVal pPath As String) As mscorlib.FileAttributes
Attribute GetAttributes.VB_Description = "Gets the FileAttributes of the file on the path."
   GetAttributes = this.FileSingleton.GetAttributes(pPath)
End Function

'@Description("Returns the creation date and time of the specified file or directory.")
Public Function GetCreationTime(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetCreationTime.VB_Description = "Returns the creation date and time of the specified file or directory."
   Set GetCreationTime = this.FileSingleton.GetCreationTime(pPath)
End Function

'@Description("Returns the creation date and time, in Coordinated Universal Time (UTC), of the specified file or directory.")
Public Function GetCreationTimeUtc(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetCreationTimeUtc.VB_Description = "Returns the creation date and time, in Coordinated Universal Time (UTC), of the specified file or directory."
   Set GetCreationTimeUtc = this.FileSingleton.GetCreationTimeUtc(pPath)
End Function

'@Description("Returns the date and time the specified file or directory was last accessed.")
Public Function GetLastAccessTime(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastAccessTime.VB_Description = "Returns the date and time the specified file or directory was last accessed."
   Set GetLastAccessTime = this.FileSingleton.GetLastAccessTime(pPath)
End Function

'@Description("Returns the date and time, in Coordinated Universal Time (UTC), that the specified file or directory was last accessed.")
Public Function GetLastAccessTimeUtc(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastAccessTimeUtc.VB_Description = "Returns the date and time, in Coordinated Universal Time (UTC), that the specified file or directory was last accessed."
   Set GetLastAccessTimeUtc = this.FileSingleton.GetLastAccessTimeUtc(pPath)
End Function

'@Description("Returns the date and time the specified file or directory was last written to.")
Public Function GetLastWriteTime(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastWriteTime.VB_Description = "Returns the date and time the specified file or directory was last written to."
   Set GetLastWriteTime = this.FileSingleton.GetLastWriteTime(pPath)
End Function

'@Description("Returns the date and time, in Coordinated Universal Time (UTC), that the specified file or directory was last written to.")
Public Function GetLastWriteTimeUtc(ByVal pPath As String) As DotNetLib.DateTime
Attribute GetLastWriteTimeUtc.VB_Description = "Returns the date and time, in Coordinated Universal Time (UTC), that the specified file or directory was last written to."
   Set GetLastWriteTimeUtc = this.FileSingleton.GetLastWriteTimeUtc(pPath)
End Function

'@Description("Moves a specified file to a new location, providing the option to specify a new file name.")
Public Sub Move(ByVal sourceFileName As String, ByVal destFileName As String)
Attribute Move.VB_Description = "Moves a specified file to a new location, providing the option to specify a new file name."
    Call this.FileSingleton.Move(sourceFileName, destFileName)
End Sub

'@Description("Opens a FileStream on the specified path with read/write access with no sharing.")
Public Function Open1(ByVal pPath As String, ByVal mode As mscorlib.FileMode) As mscorlib.FileStream
Attribute Open1.VB_Description = "Opens a FileStream on the specified path with read/write access with no sharing."
   Set Open1 = this.FileSingleton.Open(pPath, mode)
End Function

'@Description("Opens a FileStream on the specified path, with the specified mode and access with no sharing.")
Public Function Open2(ByVal pPath As String, ByVal mode As mscorlib.FileMode, ByVal access As mscorlib.FileAccess) As mscorlib.FileStream
Attribute Open2.VB_Description = "Opens a FileStream on the specified path, with the specified mode and access with no sharing."
   Set Open2 = this.FileSingleton.Open_2(pPath, mode, access)
End Function

'@Description("Opens a FileStream on the specified path, having the specified mode with read, write, or read/write access and the specified sharing option.")
Public Function Open3(ByVal pPath As String, ByVal mode As mscorlib.FileMode, ByVal access As mscorlib.FileAccess, ByVal share As mscorlib.FileShare) As mscorlib.FileStream
Attribute Open3.VB_Description = "Opens a FileStream on the specified path, having the specified mode with read, write, or read/write access and the specified sharing option."
   Set Open3 = this.FileSingleton.Open_3(pPath, mode, access, share)
End Function

'@Description("Opens an existing file for reading.")
Public Function OpenRead(ByVal pPath As String) As mscorlib.FileStream
Attribute OpenRead.VB_Description = "Opens an existing file for reading."
   Set OpenRead = this.FileSingleton.OpenRead(pPath)
End Function

'@Description("Opens an existing UTF-8 encoded text file for reading.")
Public Function OpenText(ByVal pPath As String) As mscorlib.StreamReader
Attribute OpenText.VB_Description = "Opens an existing UTF-8 encoded text file for reading."
   Set OpenText = this.FileSingleton.OpenText(pPath)
End Function

'@Description("Opens an existing file or creates a new file for writing.")
Public Function OpenWrite(ByVal pPath As String) As mscorlib.FileStream
Attribute OpenWrite.VB_Description = "Opens an existing file or creates a new file for writing."
   Set OpenWrite = this.FileSingleton.OpenWrite(pPath)
End Function

'@Description("Opens a binary file, reads the contents of the file into a byte array, and then closes the file.")
Public Function ReadAllBytes(ByVal pPath As String) As Byte()
Attribute ReadAllBytes.VB_Description = "Opens a binary file, reads the contents of the file into a byte array, and then closes the file."
   ReadAllBytes = this.FileSingleton.ReadAllBytes(pPath)
End Function

'@Description("Opens a text file, reads all lines of the file, and then closes the file.")
Public Function ReadAllLines(ByVal pPath As String) As String()
Attribute ReadAllLines.VB_Description = "Opens a text file, reads all lines of the file, and then closes the file."
   ReadAllLines = this.FileSingleton.ReadAllLines(pPath)
End Function

'@Description("Opens a file, reads all lines of the file with the specified encoding, and then closes the file.")
Public Function ReadAllLines2(ByVal pPath As String, ByVal pEncoding As mscorlib.Encoding) As String()
Attribute ReadAllLines2.VB_Description = "Opens a file, reads all lines of the file with the specified encoding, and then closes the file."
   ReadAllLines2 = this.FileSingleton.ReadAllLines_2(pPath, pEncoding)
End Function

'@Description("Opens a text file, reads all the text in the file, and then closes the file.")
Public Function ReadAllText(ByVal pPath As String) As String
Attribute ReadAllText.VB_Description = "Opens a text file, reads all the text in the file, and then closes the file."
   ReadAllText = this.FileSingleton.ReadAllText(pPath)
End Function

'@Description("Reads the lines of a file.")
Public Function ReadLines(ByVal pPath As String) As mscorlib.IEnumerable
Attribute ReadLines.VB_Description = "Reads the lines of a file."
   Set ReadLines = this.FileSingleton.ReadLines(pPath)
End Function

'@Description("Read the lines of a file that has a specified encoding.")
Public Function ReadLines2(ByVal pPath As String, ByVal pEncoding As mscorlib.Encoding) As mscorlib.IEnumerable
Attribute ReadLines2.VB_Description = "Read the lines of a file that has a specified encoding."
   Set ReadLines2 = this.FileSingleton.ReadLines_2(pPath, pEncoding)
End Function

'@Description("Replaces the contents of a specified file with the contents of another file, deleting the original file, and creating a backup of the replaced file.")
Public Sub Replace(ByVal sourceFileName As String, ByVal destinationFileName As String, ByVal destinationBackupFileName As String)
Attribute Replace.VB_Description = "Replaces the contents of a specified file with the contents of another file, deleting the original file, and creating a backup of the replaced file."
    Call this.FileSingleton.Replace(sourceFileName, destinationFileName, destinationBackupFileName)
End Sub

'@Description("Replaces the contents of a specified file with the contents of another file, deleting the original file, and creating a backup of the replaced file and optionally ignores merge errors.")
Public Sub Replace2(ByVal sourceFileName As String, ByVal destinationFileName As String, ByVal destinationBackupFileName As String, ByVal ignoreMetadataErrors As Boolean)
Attribute Replace2.VB_Description = "Replaces the contents of a specified file with the contents of another file, deleting the original file, and creating a backup of the replaced file and optionally ignores merge errors."
    Call this.FileSingleton.Replace_2(sourceFileName, destinationFileName, destinationBackupFileName, ignoreMetadataErrors)
End Sub

'@Description("Applies access control list (ACL) entries described by a FileSecurity object to the specified file.")
Public Sub SetAccessControl(ByVal pPath As String, ByVal fileSecurity As Object)
Attribute SetAccessControl.VB_Description = "Applies access control list (ACL) entries described by a FileSecurity object to the specified file."
    Call this.FileSingleton.SetAccessControl(pPath, fileSecurity)
End Sub

'@Description("Sets the specified FileAttributes of the file on the specified path.")
Public Sub SetAttributes(ByVal pPath As String, ByVal pFileAttributes As mscorlib.FileAttributes)
Attribute SetAttributes.VB_Description = "Sets the specified FileAttributes of the file on the specified path."
    Call this.FileSingleton.SetAttributes(pPath, pFileAttributes)
End Sub

'@Description("Sets the date and time the file was created.")
Public Sub SetCreationTime(ByVal pPath As String, ByVal creationTime As DotNetLib.DateTime)
Attribute SetCreationTime.VB_Description = "Sets the date and time the file was created."
    Call this.FileSingleton.SetCreationTime(pPath, creationTime)
End Sub

'@Description("Sets the date and time, in Coordinated Universal Time (UTC), that the file was created.")
Public Sub SetCreationTimeUtc(ByVal pPath As String, ByVal creationTimeUtc As DotNetLib.DateTime)
Attribute SetCreationTimeUtc.VB_Description = "Sets the date and time, in Coordinated Universal Time (UTC), that the file was created."
    Call this.FileSingleton.SetCreationTimeUtc(pPath, creationTimeUtc)
End Sub

'@Description("Sets the date and time the specified file was last accessed.")
Public Sub SetLastAccessTime(ByVal pPath As String, ByVal lastAccessTime As DotNetLib.DateTime)
Attribute SetLastAccessTime.VB_Description = "Sets the date and time the specified file was last accessed."
    Call this.FileSingleton.SetLastAccessTime(pPath, lastAccessTime)
End Sub

'@Description("Sets the date and time, in Coordinated Universal Time (UTC), that the specified file was last accessed.")
Public Sub SetLastAccessTimeUtc(ByVal pPath As String, ByVal lastAccessTimeUtc As DotNetLib.DateTime)
Attribute SetLastAccessTimeUtc.VB_Description = "Sets the date and time, in Coordinated Universal Time (UTC), that the specified file was last accessed."
    Call this.FileSingleton.SetLastAccessTimeUtc(pPath, lastAccessTimeUtc)
End Sub

'@Description("Sets the date and time that the specified file was last written to.")
Public Sub SetLastWriteTime(ByVal pPath As String, ByVal lastWriteTime As DotNetLib.DateTime)
Attribute SetLastWriteTime.VB_Description = "Sets the date and time that the specified file was last written to."
    Call this.FileSingleton.SetLastWriteTime(pPath, lastWriteTime)
End Sub

'@Description("Sets the date and time, in Coordinated Universal Time (UTC), that the specified file was last written to.")
Public Sub SetLastWriteTimeUtc(ByVal pPath As String, ByVal lastWriteTimeUtc As DotNetLib.DateTime)
Attribute SetLastWriteTimeUtc.VB_Description = "Sets the date and time, in Coordinated Universal Time (UTC), that the specified file was last written to."
    Call this.FileSingleton.SetLastWriteTimeUtc(pPath, lastWriteTimeUtc)
End Sub

'@Description("Creates a new file, writes the specified byte array to the file, and then closes the file. If the target file already exists, it is overwritten.")
Public Sub WriteAllBytes(ByVal pPath As String, ByRef Bytes() As Byte)
Attribute WriteAllBytes.VB_Description = "Creates a new file, writes the specified byte array to the file, and then closes the file. If the target file already exists, it is overwritten."
    Call this.FileSingleton.WriteAllBytes(pPath, Bytes)
End Sub

'@Description("Creates a new file, write the specified string array to the file, and then closes the file.")
Public Sub WriteAllLines(ByVal pPath As String, ByRef contents() As String)
Attribute WriteAllLines.VB_Description = "Creates a new file, write the specified string array to the file, and then closes the file."
   Call this.FileSingleton.WriteAllLines(pPath, contents)
End Sub

'@Description("Creates a new file, writes the specified string array to the file by using the specified encoding, and then closes the file.")
Public Sub WriteAllLines2(ByVal pPath As String, ByRef contents() As String, ByVal pEncoding As mscorlib.Encoding)
Attribute WriteAllLines2.VB_Description = "Creates a new file, writes the specified string array to the file by using the specified encoding, and then closes the file."
   Call this.FileSingleton.WriteAllLines_2(pPath, contents, pEncoding)
End Sub

'@Description("Creates a new file, writes a collection of strings to the file, and then closes the file.")
Public Sub WriteAllLines3(ByVal pPath As String, ByVal contents As mscorlib.IEnumerable)
Attribute WriteAllLines3.VB_Description = "Creates a new file, writes a collection of strings to the file, and then closes the file."
    Call this.FileSingleton.WriteAllLines_3(pPath, contents)
End Sub

'@Description("Creates a new file by using the specified encoding, writes a collection of strings to the file, and then closes the file.")
Public Sub WriteAllLines4(ByVal pPath As String, ByVal contents As mscorlib.IEnumerable, ByVal pEncoding As mscorlib.Encoding)
Attribute WriteAllLines4.VB_Description = "Creates a new file by using the specified encoding, writes a collection of strings to the file, and then closes the file."
    Call this.FileSingleton.WriteAllLines_4(pPath, contents, pEncoding)
End Sub

'@Description("Creates a new file, writes the specified string to the file, and then closes the file. If the target file already exists, it is overwritten.")
Public Sub WriteAllText(ByVal pPath As String, ByVal contents As String)
Attribute WriteAllText.VB_Description = "Creates a new file, writes the specified string to the file, and then closes the file. If the target file already exists, it is overwritten."
    Call this.FileSingleton.WriteAllText(pPath, contents)
End Sub

'@Description("Creates a new file, writes the specified string to the file using the specified encoding, and then closes the file. If the target file already exists, it is overwritten.")
Public Sub WriteAllText2(ByVal pPath As String, ByVal contents As String, ByVal pEncoding As mscorlib.Encoding)
Attribute WriteAllText2.VB_Description = "Creates a new file, writes the specified string to the file using the specified encoding, and then closes the file. If the target file already exists, it is overwritten."
    Call this.FileSingleton.WriteAllText_2(pPath, contents, pEncoding)
End Sub
