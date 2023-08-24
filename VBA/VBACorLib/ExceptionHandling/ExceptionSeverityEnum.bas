Attribute VB_Name = "ExceptionSeverityEnum"
Attribute VB_Description = "Represents the severity of error that occur during application execution."
'@ModuleDescription("Represents the severity of error that occur during application execution.")
'@Folder "VBACorLib.ExceptionHandling"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 23, 2023
'@LastModified August 24, 2023

Option Explicit

Public Enum ExceptionSeverity
    Unspecified = 0
    Critical = 1
    Warning = 2
End Enum
