Attribute VB_Name = "RubberduckExport"
Attribute VB_Description = "Rubberduck utility to export all components in the active project according to RD @Folder annotation"
'@ModuleDescription("Rubberduck utility to export all components in the active project according to RD @Folder annotation")
'@Folder "<Rubberduck Utilities>"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/RubberduckUtility
'@Version v2.0 August 18, 2023
'@LastModified August 24, 2023

'@Dependenicies
'   RubberduckUtility.cls
'   ExceptionSeverityEnum.bas
'   Exception.cls
'   IException.cls
'   Exceptions.cls

Option Explicit

Public Sub RubberduckExportProject()
    RubberduckUtility.ExportAll "C:\VBA\Output"
    RubberduckUtility.ErrorReport Critical
    Debug.Print
    RubberduckUtility.ErrorReport Warning
    Debug.Print
    RubberduckUtility.SummaryReport
End Sub

' Example output:
'    Warning invalid Rubberduck folder characters, <Rubberduck Utilities> RubberduckUtility.cls exported to C:\VBA\Output
'    Warning invalid Rubberduck folder characters, <Rubberduck Utilities> RubberDuckExport.bas exported to C:\VBA\Output
'
'    Total files exported to C:\VBA\Output : 216
'    Total warnings: 2
'    Total failed exports : 0

