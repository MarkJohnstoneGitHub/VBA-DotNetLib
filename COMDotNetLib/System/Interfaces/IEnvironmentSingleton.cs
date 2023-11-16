//  https://learn.microsoft.com/en-us/dotnet/api/system.environment?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using DotNetLib.System.Enviroment;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("379D6AB2-5FAF-4E7B-A979-2FB291F72312")]
    [Description("Provides information about, and means to manipulate, the current environment and platform. This class cannot be inherited.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IEnvironmentSingleton
    {
        string CommandLine 
        {
            [Description("A string containing command-line arguments.")]
            get;
        }

        string CurrentDirectory 
        {
            [Description("Gets or sets the fully qualified path of the current working directory.")]
            get;
            [Description("Gets or sets the fully qualified path of the current working directory.")]
            set;
        }

        int CurrentManagedThreadId 
        {
            [Description("Gets a unique identifier for the current managed thread.")]
            get;
        }

        int ExitCode 
        {
            [Description("Gets or sets the exit code of the process.")]
            get;
            [Description("Gets or sets the exit code of the process.")]
            set;
        }

        bool HasShutdownStarted 
        {
            [Description("Gets a value that indicates whether the current application domain is being unloaded or the common language runtime (CLR) is shutting down.")]
            get;
        }

        bool Is64BitOperatingSystem 
        {
            [Description("Gets a value that indicates whether the current operating system is a 64-bit operating system.")]
            get;
        }

        bool Is64BitProcess 
        {
            [Description("Gets a value that indicates whether the current process is a 64-bit process.")]
            get;
        }

        string MachineName 
        {
            [Description("Gets the NetBIOS name of this local computer.")]
            get;
        }

        string NewLine 
        {
            [Description("Gets the newline string defined for this environment.")]
            get;
        }

        GSystem.OperatingSystem OSVersion //Todo GSystem.OperatingSystem
        {
            [Description("Gets the current platform identifier and version number.")]
            get;
        }

        int ProcessorCount 
        {
            [Description("Gets the number of processors available to the current process.")]
            get;
        }

        string StackTrace 
        {
            [Description("Gets current stack trace information.")]
            get;
        }

        string SystemDirectory 
        {
            [Description("Gets the fully qualified path of the system directory.")]
            get;
        }

        int SystemPageSize 
        {
            [Description("Gets the number of bytes in the operating system's memory page.")]
            get;
        }


        int TickCount 
        {
            [Description("Gets the number of milliseconds elapsed since the system started.")]
            get;
        }

        string UserDomainName 
        {
            [Description("Gets the network domain name associated with the current user.")]
            get;
        }

        bool UserInteractive 
        {
            [Description("Gets a value indicating whether the current process is running in user interactive mode.")]
            get;
        }

        string UserName 
        {
            [Description("Gets the user name of the person who is associated with the current thread.")]
            get;
        }

        GSystem.Version Version //Todo GSystem.Version 
        {
            [Description("Gets a version consisting of the major, minor, build, and revision numbers of the common language runtime.")]
            get;
        }

        long WorkingSet 
        {
            [Description("Gets the amount of physical memory mapped to the process context.")]
            get;
        }

        [Description("Terminates this process and returns an exit code to the operating system.")]
        void Exit(int exitCode);

        [Description("Replaces the name of each environment variable embedded in the specified string with the string equivalent of the value of the variable, then returns the resulting string.")]
        string ExpandEnvironmentVariables(string name);

        [Description("Immediately terminates a process after writing a message to the Windows Application event log, and then includes the message in error reporting to Microsoft.")]
        void FailFast(string message);

        [Description("Immediately terminates a process after writing a message to the Windows Application event log, and then includes the message and exception information in error reporting to Microsoft.")]
        void FailFast(string message, Exception exception); //Todo  Exception

        [Description("Returns a string array containing the command-line arguments for the current process.")]
        string[] GetCommandLineArgs();

        [Description("Retrieves the value of an environment variable from the current process.")]
        string GetEnvironmentVariable(string variable);

        [Description("Retrieves the value of an environment variable from the current process or from the Windows operating system registry key for the current user or local machine.")]
        string GetEnvironmentVariable(string variable, GSystem.EnvironmentVariableTarget target);

        [Description("Retrieves all environment variable names and their values from the current process.")]
        GCollections.IDictionary GetEnvironmentVariables();

        [Description("Retrieves all environment variable names and their values from the current process, or from the Windows operating system registry key for the current user or local machine.")]
        GCollections.IDictionary GetEnvironmentVariables(GSystem.EnvironmentVariableTarget target);

        [Description("Gets the path to the specified system special folder.")]
        string GetFolderPath(SpecialFolder folder);

        [Description("Gets the path to the specified system special folder using a specified option for accessing special folders.")]
        string GetFolderPath(SpecialFolder folder, SpecialFolderOption option);

        [Description("Returns an array of string containing the names of the logical drives on the current computer.")]
        string[] GetLogicalDrives();

        [Description("Creates, modifies, or deletes an environment variable stored in the current process.")]
        void SetEnvironmentVariable(string variable, string value);

        [Description("Creates, modifies, or deletes an environment variable stored in the current process or in the Windows operating system registry key reserved for the current user or local machine.")]
        void SetEnvironmentVariable(string variable, string value, GSystem.EnvironmentVariableTarget target);

    }
}
