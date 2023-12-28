VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides information about, and means to manipulate, the current environment and platform. This class cannot be inherited."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Provides information about, and means to manipulate, the current environment and platform. This class cannot be inherited.")
'@PredeclaredId
'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 8, 2023
'@LastModified November 9, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.environment?view=netframework-4.8.1

'@Remarks
'Singleton class for Environment containing factory methods and static members.

Option Explicit

Private Type TEnvironmentSingleton
   EnvironmentSingleton As DotNetLib.EnvironmentSingleton
End Type

Private this As TEnvironmentSingleton

Private Sub Class_Initialize()
   Set this.EnvironmentSingleton = New DotNetLib.EnvironmentSingleton
End Sub

Private Sub Class_Terminate()
   Set this.EnvironmentSingleton = Nothing
End Sub

'@Description("A string containing command-line arguments.")
Public Property Get CommandLine() As String
Attribute CommandLine.VB_Description = "A string containing command-line arguments."
   CommandLine = this.EnvironmentSingleton.CommandLine()
End Property

'@Description("Gets or sets the fully qualified path of the current working directory.")
Public Property Get CurrentDirectory() As String
Attribute CurrentDirectory.VB_Description = "Gets or sets the fully qualified path of the current working directory."
   CurrentDirectory = this.EnvironmentSingleton.CurrentDirectory()
End Property

'@Description("Gets or sets the fully qualified path of the current working directory.")
Public Property Let CurrentDirectory(ByVal pCurrentDirectory As String)
Attribute CurrentDirectory.VB_Description = "Gets or sets the fully qualified path of the current working directory."
    this.EnvironmentSingleton.CurrentDirectory = pCurrentDirectory
End Property

'@Description("Gets a unique identifier for the current managed thread.")
Public Property Get CurrentManagedThreadId() As Long
Attribute CurrentManagedThreadId.VB_Description = "Gets a unique identifier for the current managed thread."
   CurrentManagedThreadId = this.EnvironmentSingleton.CurrentManagedThreadId()
End Property

'@Description("Gets or sets the exit code of the process.")
Public Property Get ExitCode() As Long
Attribute ExitCode.VB_Description = "Gets or sets the exit code of the process."
   ExitCode = this.EnvironmentSingleton.ExitCode()
End Property

'@Description("Gets or sets the exit code of the process.")
Public Property Let ExitCode(ByVal pExitCode As Long)
Attribute ExitCode.VB_Description = "Gets or sets the exit code of the process."
    this.EnvironmentSingleton.ExitCode = pExitCode
End Property

'@Description("Gets a value that indicates whether the current application domain is being unloaded or the common language runtime (CLR) is shutting down.")
Public Property Get HasShutdownStarted() As Boolean
Attribute HasShutdownStarted.VB_Description = "Gets a value that indicates whether the current application domain is being unloaded or the common language runtime (CLR) is shutting down."
   HasShutdownStarted = this.EnvironmentSingleton.HasShutdownStarted()
End Property

'@Description("Gets a value that indicates whether the current operating system is a 64-bit operating system.")
Public Property Get Is64BitOperatingSystem() As Boolean
Attribute Is64BitOperatingSystem.VB_Description = "Gets a value that indicates whether the current operating system is a 64-bit operating system."
   Is64BitOperatingSystem = this.EnvironmentSingleton.Is64BitOperatingSystem()
End Property

'@Description("Gets a value that indicates whether the current process is a 64-bit process.")
Public Property Get Is64BitProcess() As Boolean
Attribute Is64BitProcess.VB_Description = "Gets a value that indicates whether the current process is a 64-bit process."
   Is64BitProcess = this.EnvironmentSingleton.Is64BitProcess()
End Property

'@Description("Gets the NetBIOS name of this local computer.")
Public Property Get MachineName() As String
Attribute MachineName.VB_Description = "Gets the NetBIOS name of this local computer."
   MachineName = this.EnvironmentSingleton.MachineName()
End Property

'@Description("Gets the newline string defined for this environment.")
Public Property Get NewLine() As String
Attribute NewLine.VB_Description = "Gets the newline string defined for this environment."
   NewLine = this.EnvironmentSingleton.NewLine()
End Property

'@Description("Gets the current platform identifier and version number.")
Public Property Get OSVersion() As mscorlib.OperatingSystem
Attribute OSVersion.VB_Description = "Gets the current platform identifier and version number."
   Set OSVersion = this.EnvironmentSingleton.OSVersion()
End Property

'@Description("Gets the number of processors available to the current process.")
Public Property Get ProcessorCount() As Long
Attribute ProcessorCount.VB_Description = "Gets the number of processors available to the current process."
   ProcessorCount = this.EnvironmentSingleton.ProcessorCount()
End Property

'@Description("Gets current stack trace information.")
Public Property Get StackTrace() As String
Attribute StackTrace.VB_Description = "Gets current stack trace information."
   StackTrace = this.EnvironmentSingleton.StackTrace()
End Property

'@Description("Gets the fully qualified path of the system directory.")
Public Property Get SystemDirectory() As String
Attribute SystemDirectory.VB_Description = "Gets the fully qualified path of the system directory."
   SystemDirectory = this.EnvironmentSingleton.SystemDirectory()
End Property

'@Description("Gets the number of bytes in the operating system's memory page.")
Public Property Get SystemPageSize() As Long
Attribute SystemPageSize.VB_Description = "Gets the number of bytes in the operating system's memory page."
   SystemPageSize = this.EnvironmentSingleton.SystemPageSize()
End Property

'@Description("Gets the number of milliseconds elapsed since the system started.")
Public Property Get TickCount() As Long
Attribute TickCount.VB_Description = "Gets the number of milliseconds elapsed since the system started."
   TickCount = this.EnvironmentSingleton.TickCount()
End Property

'@Description("Gets the network domain name associated with the current user.")
Public Property Get UserDomainName() As String
Attribute UserDomainName.VB_Description = "Gets the network domain name associated with the current user."
   UserDomainName = this.EnvironmentSingleton.UserDomainName()
End Property

'@Description("Gets a value indicating whether the current process is running in user interactive mode.")
Public Property Get UserInteractive() As Boolean
Attribute UserInteractive.VB_Description = "Gets a value indicating whether the current process is running in user interactive mode."
   UserInteractive = this.EnvironmentSingleton.UserInteractive()
End Property

'@Description("Gets the user name of the person who is associated with the current thread.")
Public Property Get UserName() As String
Attribute UserName.VB_Description = "Gets the user name of the person who is associated with the current thread."
   UserName = this.EnvironmentSingleton.UserName()
End Property

'@Description("Gets a version consisting of the major, minor, build, and revision numbers of the common language runtime.")
Public Property Get Version() As mscorlib.Version
Attribute Version.VB_Description = "Gets a version consisting of the major, minor, build, and revision numbers of the common language runtime."
   Set Version = this.EnvironmentSingleton.Version()
End Property

'@Description("Gets the amount of physical memory mapped to the process context.")
Public Property Get WorkingSet() As LongLong
Attribute WorkingSet.VB_Description = "Gets the amount of physical memory mapped to the process context."
   WorkingSet = this.EnvironmentSingleton.WorkingSet()
End Property

'@Description("Terminates this process and returns an exit code to the operating system.")
'@Remarks Renamed due to VBA reserved word
Public Sub ExitProcess(ByVal pExitCode As Long)
Attribute ExitProcess.VB_Description = "Terminates this process and returns an exit code to the operating system."
    Call this.EnvironmentSingleton.Exit(pExitCode)
End Sub

'@Description("Replaces the name of each environment variable embedded in the specified string with the string equivalent of the value of the variable, then returns the resulting string.")
Public Function ExpandEnvironmentVariables(ByVal Name As String) As String
Attribute ExpandEnvironmentVariables.VB_Description = "Replaces the name of each environment variable embedded in the specified string with the string equivalent of the value of the variable, then returns the resulting string."
   ExpandEnvironmentVariables = this.EnvironmentSingleton.ExpandEnvironmentVariables(Name)
End Function

'@Description("Immediately terminates a process after writing a message to the Windows Application event log, and then includes the message in error reporting to Microsoft.")
Public Sub FailFast(ByVal message As String)
Attribute FailFast.VB_Description = "Immediately terminates a process after writing a message to the Windows Application event log, and then includes the message in error reporting to Microsoft."
    Call this.EnvironmentSingleton.FailFast(message)
End Sub

'@Description("Immediately terminates a process after writing a message to the Windows Application event log, and then includes the message and exception information in error reporting to Microsoft.")
Public Sub FailFast2(ByVal message As String, ByVal pException As mscorlib.Exception)
Attribute FailFast2.VB_Description = "Immediately terminates a process after writing a message to the Windows Application event log, and then includes the message and exception information in error reporting to Microsoft."
    Call this.EnvironmentSingleton.FailFast_2(message, pException)
End Sub

'@Description("Returns a string array containing the command-line arguments for the current process.")
Public Function GetCommandLineArgs() As String()
Attribute GetCommandLineArgs.VB_Description = "Returns a string array containing the command-line arguments for the current process."
   GetCommandLineArgs = this.EnvironmentSingleton.GetCommandLineArgs()
End Function

'@Description("Retrieves the value of an environment variable from the current process.")
Public Function GetEnvironmentVariable(ByVal variable As String) As String
Attribute GetEnvironmentVariable.VB_Description = "Retrieves the value of an environment variable from the current process."
   GetEnvironmentVariable = this.EnvironmentSingleton.GetEnvironmentVariable(variable)
End Function

'@Description("Retrieves the value of an environment variable from the current process or from the Windows operating system registry key for the current user or local machine.")
Public Function GetEnvironmentVariable2(ByVal variable As String, ByVal target As mscorlib.EnvironmentVariableTarget) As String
Attribute GetEnvironmentVariable2.VB_Description = "Retrieves the value of an environment variable from the current process or from the Windows operating system registry key for the current user or local machine."
   GetEnvironmentVariable2 = this.EnvironmentSingleton.GetEnvironmentVariable_2(variable, target)
End Function

'@Description("Retrieves all environment variable names and their values from the current process.")
Public Function GetEnvironmentVariables() As mscorlib.IDictionary
Attribute GetEnvironmentVariables.VB_Description = "Retrieves all environment variable names and their values from the current process."
   Set GetEnvironmentVariables = this.EnvironmentSingleton.GetEnvironmentVariables()
End Function

'@Description("Retrieves all environment variable names and their values from the current process, or from the Windows operating system registry key for the current user or local machine.")
Public Function GetEnvironmentVariables2(ByVal target As mscorlib.EnvironmentVariableTarget) As mscorlib.IDictionary
Attribute GetEnvironmentVariables2.VB_Description = "Retrieves all environment variable names and their values from the current process, or from the Windows operating system registry key for the current user or local machine."
   Set GetEnvironmentVariables2 = this.EnvironmentSingleton.GetEnvironmentVariables_2(target)
End Function

'@Description("Gets the path to the specified system special folder.")
Public Function GetFolderPath(ByVal folder As DotNetLib.SpecialFolder) As String
Attribute GetFolderPath.VB_Description = "Gets the path to the specified system special folder."
   GetFolderPath = this.EnvironmentSingleton.GetFolderPath(folder)
End Function

'@Description("Gets the path to the specified system special folder using a specified option for accessing special folders.")
Public Function GetFolderPath2(ByVal pfolder As DotNetLib.SpecialFolder, ByVal pOption As DotNetLib.SpecialFolderOption) As String
Attribute GetFolderPath2.VB_Description = "Gets the path to the specified system special folder using a specified option for accessing special folders."
   GetFolderPath2 = this.EnvironmentSingleton.GetFolderPath_2(pfolder, pOption)
End Function

'@Description("Returns an array of string containing the names of the logical drives on the current computer.")
Public Function GetLogicalDrives() As String()
Attribute GetLogicalDrives.VB_Description = "Returns an array of string containing the names of the logical drives on the current computer."
   GetLogicalDrives = this.EnvironmentSingleton.GetLogicalDrives()
End Function

'@Description("Creates, modifies, or deletes an environment variable stored in the current process.")
Public Sub SetEnvironmentVariable(ByVal variable As String, ByVal value As String)
Attribute SetEnvironmentVariable.VB_Description = "Creates, modifies, or deletes an environment variable stored in the current process."
    Call this.EnvironmentSingleton.SetEnvironmentVariable(variable, value)
End Sub

'@Description("Creates, modifies, or deletes an environment variable stored in the current process or in the Windows operating system registry key reserved for the current user or local machine.")
Public Sub SetEnvironmentVariable2(ByVal variable As String, ByVal value As String, ByVal target As mscorlib.EnvironmentVariableTarget)
Attribute SetEnvironmentVariable2.VB_Description = "Creates, modifies, or deletes an environment variable stored in the current process or in the Windows operating system registry key reserved for the current user or local machine."
    Call this.EnvironmentSingleton.SetEnvironmentVariable_2(variable, value, target)
End Sub
