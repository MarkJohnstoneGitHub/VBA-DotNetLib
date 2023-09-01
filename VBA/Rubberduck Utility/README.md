# Rubberduck Export All Utility
[Rubberduck](https://rubberduckvba.com/) utility to export all components in the active project according to Rubberduck [@Folder annotation](https://github.com/rubberduck-vba/Rubberduck/wiki/Using-@Folder-Annotations). 

**Required references: VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3)**

**Dependenicies**
  - [RubberduckUtility.cls](https://github.com/MarkJohnstoneGitHub/RubberduckUtility/blob/main/RubberduckUtility/RubberduckUtility.cls)
  - [ExceptionSeverityEnum.bas](https://github.com/MarkJohnstoneGitHub/RubberduckUtility/blob/main/RubberduckUtility/VBACorLib/ExceptionHandling/ExceptionSeverityEnum.bas)
  - [Exception.cls](https://github.com/MarkJohnstoneGitHub/RubberduckUtility/blob/main/RubberduckUtility/VBACorLib/ExceptionHandling/Exception.cls)
  - [IException.cls](https://github.com/MarkJohnstoneGitHub/RubberduckUtility/blob/main/RubberduckUtility/VBACorLib/ExceptionHandling/IException.cls)
  - [Exceptions.cls](https://github.com/MarkJohnstoneGitHub/RubberduckUtility/blob/main/RubberduckUtility/VBACorLib/ExceptionHandling/Exceptions.cls)

**Usage**

May also enter in the immediate window and run the eg. RubberduckUtility.ExportAll "C:\VBA\Output"

 ```
Public Sub RubberduckExportProject()
    RubberduckUtility.ExportAll "C:\VBA\Output"
    RubberduckUtility.ErrorReport Critical
    Debug.Print
    RubberduckUtility.ErrorReport Warning
    Debug.Print
    RubberduckUtility.SummaryReport
End Sub
 ```

**Output Example**
 ```
Warning invalid Rubberduck folder characters, <Rubberduck Utilities> RubberduckUtility.cls exported to C:\VBA\Output
Warning invalid Rubberduck folder characters, <Rubberduck Utilities> RubberDuckExport.bas exported to C:\VBA\Output

Total files exported to C:\VBA\Output : 216
Total warnings: 2
Total failed exports : 0
 ```

RubberduckUtility.ExportAll exports all components to the working directory provided for the active project. Sub folders are created according to the according to Rubberduck @Folder annotation and ***existing files are overwritten***.

Added error handling for: 

Invalid working directory.  
  - Error is raised.  The working directory must already exist.
  - Critical Errors logged for failed exports. This may occur if files attempting to overwrite are read-only or don't have permission.
  - Warnings logged for Rubberduck @Folder annotations that contain invalid folder characters. They are exported to the output directory provided.


