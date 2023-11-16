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
    [Description("Provides information about, and means to manipulate, the current environment and platform. This class cannot be inherited.")]
    [Guid("96497DFC-5D40-4F5C-8546-3394620D3610")]
    [ProgId("DotNetLib.System.EnvironmentSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IEnvironmentSingleton))]
    public class EnvironmentSingleton : IEnvironmentSingleton
    {
        public EnvironmentSingleton() { }

        public string CommandLine => GSystem.Environment.CommandLine;

        public string CurrentDirectory 
        {
            get => GSystem.Environment.CurrentDirectory;
            set => GSystem.Environment.CurrentDirectory = value;
        }

        public int CurrentManagedThreadId => GSystem.Environment.CurrentManagedThreadId;

        public int ExitCode 
        {
            get => GSystem.Environment.ExitCode;
            set => GSystem.Environment.ExitCode = value;
        }

        public bool HasShutdownStarted => GSystem.Environment.HasShutdownStarted;

        public bool Is64BitOperatingSystem => GSystem.Environment.Is64BitOperatingSystem;

        public bool Is64BitProcess => GSystem.Environment.Is64BitProcess;

        public string MachineName => GSystem.Environment.MachineName;

        public string NewLine => GSystem.Environment.NewLine;

        public OperatingSystem OSVersion => GSystem.Environment.OSVersion;

        public int ProcessorCount => GSystem.Environment.ProcessorCount;

        public string StackTrace => GSystem.Environment.StackTrace;

        public string SystemDirectory => GSystem.Environment.SystemDirectory;

        public int SystemPageSize => GSystem.Environment.SystemPageSize;

        public int TickCount => GSystem.Environment.TickCount;

        public string UserDomainName => GSystem.Environment.UserDomainName;

        public bool UserInteractive => GSystem.Environment.UserInteractive;

        public string UserName => GSystem.Environment.UserName;

        public GSystem.Version Version => GSystem.Environment.Version; // Todo

        public long WorkingSet => GSystem.Environment.WorkingSet;

        // Methods
        public void Exit(int exitCode)
        { 
            GSystem.Environment.Exit(exitCode);
        }

        public string ExpandEnvironmentVariables(string name)
        {
            return GSystem.Environment.ExpandEnvironmentVariables(name);
        }

        public void FailFast(string message)
        {
            GSystem.Environment.FailFast(message);
        }

        public void FailFast(string message, Exception exception)
        {
            GSystem.Environment.FailFast(message, exception);
        }

        public string[] GetCommandLineArgs()
        {
            return GSystem.Environment.GetCommandLineArgs();
        }

        public string GetEnvironmentVariable(string variable)
        {
            return GSystem.Environment.GetEnvironmentVariable(variable);
        }

        public string GetEnvironmentVariable(string variable, GSystem.EnvironmentVariableTarget target)
        {
            return GSystem.Environment.GetEnvironmentVariable(variable, target);
        }

        public GCollections.IDictionary GetEnvironmentVariables()
        {
            return GSystem.Environment.GetEnvironmentVariables();
        }

        public GCollections.IDictionary GetEnvironmentVariables(GSystem.EnvironmentVariableTarget target)
        {
            return GSystem.Environment.GetEnvironmentVariables(target);
        }

        public string GetFolderPath(SpecialFolder folder)
        {
            return GSystem.Environment.GetFolderPath((GSystem.Environment.SpecialFolder)folder);
        }

        public string GetFolderPath(SpecialFolder folder, SpecialFolderOption option)
        {
            return GSystem.Environment.GetFolderPath((GSystem.Environment.SpecialFolder)folder, (GSystem.Environment.SpecialFolderOption)option);
        }

        public string[] GetLogicalDrives()
        {
            return GSystem.Environment.GetLogicalDrives();
        }

        public void SetEnvironmentVariable(string variable, string value)
        {
            GSystem.Environment.SetEnvironmentVariable(variable, value);
        }

        public void SetEnvironmentVariable(string variable, string value, GSystem.EnvironmentVariableTarget target)
        {
            GSystem.Environment.SetEnvironmentVariable(variable,value,target);
        }
    }
}
