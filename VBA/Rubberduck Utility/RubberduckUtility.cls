VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RubberduckUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Rubberduck utility to export all components according to @Folder annotation"
'@ModuleDescription("Rubberduck utility to export all components according to @Folder annotation")
'@Folder "<Rubberduck Utilities>"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/RubberduckUtility
'@Version v2.0 August 18, 2023
'@LastModified September 3, 2023

'@ReferenceAddin VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3)

'@Dependenicies
'   ExceptionSeverityEnum.bas
'   Exception.cls
'   IException.cls
'   Exceptions.cls

'@References
' http://www.cpearson.com/excel/vbe.aspx
' https://codereview.stackexchange.com/questions/248109/add-default-rubberduck-vba-folder-annotation-to-vbproject-vbcomponents
' https://codereview.stackexchange.com/questions/256782/export-import-rubberduck-vba-virtual-folder-structure-with-library-references-to
' https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files/17362688#17362688
' https://stackoverflow.com/a/27385063/10759363
' https://en.wikipedia.org/wiki/Path_(computing)#:~:text=The%20delimiting%20character%20is%20most,may%20use%20a%20different%20delimiter.
' https://codereview.stackexchange.com/questions/233587/source-control-and-custom-vba-code-exporter
' https://www.engram9.info/excel-2002-vba-xml-asp/vbproject-object-vbcomponent-object-and-reference-object.html
' https://www.rondebruin.nl/win/s9/win002.htm
' https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes

'Useage:
'   Exports all components in the active project to a base file location, overwritting any existing files.
'   RubberduckUtility.ExportAll "C:\VBA\Output"

'@Remarks
' Warnings for:
'   Invalid Rubberduck folder characters.
'       If invalid folder characters attempt to save to the working directory provided
'   File name length greater than 31 characters.
'       Module names in MS-Excel are limited to a maximum of 31 characters.
'       MS-Access appears to be the expection to this limit.
' Exports may fail if attempting to overwrite read only file in working directory or folder permissions.

Option Explicit
Private Const VBAErrorFileNotFound As Long = 76   ' DirectoryNotFoundException
Private Const DirectorySeparatorChar As String = "\"

Private Type TProject
    Project As VBIDE.VBProject
    WorkingDirectory As String
    ErrorLog As Exceptions
    TotalExported As Long
End Type

Private this As TProject

'--------------------------------------------------------------------
'Constructors
'--------------------------------------------------------------------
Private Sub Class_Initialize()
    With this
        Set .Project = Application.VBE.ActiveVBProject
        Set .ErrorLog = New Exceptions
    End With
End Sub

Private Sub Initialize()
    Set this.ErrorLog = New Exceptions
    this.TotalExported = 0
End Sub

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------

'@Description("Obtains the output directory provided for exporting project components.")
Public Property Get WorkingDirectory() As String
Attribute WorkingDirectory.VB_Description = "Obtains the output directory provided for exporting project components."
    WorkingDirectory = this.WorkingDirectory
End Property

'@Description("Obtains the error log of all errors generated from ExportAll.")
Public Property Get ErrorLog() As Exceptions
Attribute ErrorLog.VB_Description = "Obtains the error log of all errors generated from ExportAll."
    Set ErrorLog = this.ErrorLog
End Property

'@Description("Returns the number of errors in the error list.")
Public Property Get TotalExported() As Long
Attribute TotalExported.VB_Description = "Returns the number of errors in the error list."
    TotalExported = this.TotalExported
End Property

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Description("Export the active project components to file location according to the Rubberduck @Folder annotation.")
'@param workingDirectory file path to export to. eg. C:\Output
'@Exceptions: VBAErrorFileNotFound Err.Number 76
'@Remarks
'   If Rubberduck @Folder annotation contains invalid characters for a folder name
'   the file will attempted to exported to the working directory provided.
'   Components may not be exported due to attempting to overwrite read only files or permissions.
Public Sub ExportAll(ByVal workingPath As String)
Attribute ExportAll.VB_Description = "Export the active project components to file location according to the Rubberduck @Folder annotation."
    Initialize ' Clear Error log and exported total
    '@TODO remove trailing directory separator or append if doesn't exist?
    this.WorkingDirectory = Trim$(workingPath)
    
    If Not DirectoryExist(this.WorkingDirectory) Then
        Dim exceptionFileNotFound As IException
        Set exceptionFileNotFound = Exception.Create(VBAErrorFileNotFound, _
                            "RubberduckUtility.ExportAllComponents", _
                            "Output directory not found: " & this.WorkingDirectory, _
                            ExceptionSeverity.Critical)
        Exception.Throw exceptionFileNotFound
    End If
    
    Dim component As VBComponent
    For Each component In this.Project.VBComponents
        Dim fileName As String
        fileName = component.Name
        Dim extension As String
        extension = ComponentExtension(component)
        Dim folderAnnotation As String
        folderAnnotation = ParseFolderAnnotation(component)
        Dim relativePath As String
        relativePath = GetFilePathFromFolderAnnotation(folderAnnotation)
        
        On Error Resume Next
        Dim filepath As String
        CreateRelativeDirectory relativePath
        If Err.Number = 0 Then  ' Error if invalid Rubberduck folder characters or invalid working folder permission
            filepath = this.WorkingDirectory & DirectorySeparatorChar & relativePath & _
                        IIf(relativePath <> VBA.vbNullString, DirectorySeparatorChar, VBA.vbNullString) & fileName & extension
        Else
            this.ErrorLog.Add Exception.Create(Err.Number, _
                    "RubberduckUtility.ExportAllComponents", _
                    "Warning invalid Rubberduck folder characters, " & folderAnnotation & _
                    " " & fileName & extension & " exported to " & _
                    this.WorkingDirectory, ExceptionSeverity.Warning)
            filepath = this.WorkingDirectory & DirectorySeparatorChar & fileName & extension
        End If
        On Error GoTo 0 'reset error handling
        
        On Error Resume Next
        component.Export filepath
        If Err.Number = 0 Then
            this.TotalExported = this.TotalExported + 1
            If Len(fileName) > 31 Then
                this.ErrorLog.Add Exception.Create( _
                                Err.Number, _
                                "RubberduckUtility.ExportAllComponents", _
                                "Warning file name length greater than 31 characters for: " & fileName, _
                                ExceptionSeverity.Warning)
                                
            End If
        ' Error may occur for attempting to overwrite read only file etc
        ElseIf (this.WorkingDirectory & DirectorySeparatorChar & fileName & extension) = filepath Then
            this.ErrorLog.Add Exception.Create(Err.Number, _
                                "RubberduckUtility.ExportAllComponents", _
                                "Export failed for " & filepath & _
                                " Attempted to overwrite read only file or invalid folder permissions.", _
                                ExceptionSeverity.Critical)
        Else
            ' Attempt to export to working directory
            this.ErrorLog.Add Exception.Create( _
                Err.Number, _
                "RubberduckUtility.ExportAllComponents", _
                "Warning export failed for: " & filepath & " exported to " & this.WorkingDirectory, _
                ExceptionSeverity.Warning)
            filepath = this.WorkingDirectory & DirectorySeparatorChar & fileName & extension
            Err.Clear
            component.Export filepath
            If Err.Number = 0 Then
                this.TotalExported = this.TotalExported + 1
                If Len(fileName) > 31 Then
                    this.ErrorLog.Add Exception.Create( _
                        Err.Number, _
                        "RubberduckUtility.ExportAllComponents", _
                        "Warning file name length greater than 31 characters for: " & fileName, _
                        ExceptionSeverity.Warning)
                                
                End If
            Else
                'Export will fail if attempting to overwrite read only file in working directory or invalid folder permission
                this.ErrorLog.Add Exception.Create( _
                                    Err.Number, _
                                    "RubberduckUtility.ExportAllComponents", _
                                    "Export failed for " & filepath & _
                                    " Attempted to overwrite read only file or invalid folder permissions.", _
                                    ExceptionSeverity.Critical)
            End If
        End If
        On Error GoTo 0 'reset error handling
    Next
End Sub

'@Description("List of all warnings and critical errrors when exported.")
'@Reference https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
Public Sub ErrorReport(Optional ByVal errType As ExceptionSeverity = 0)
Attribute ErrorReport.VB_Description = "List of all warnings and critical errrors when exported."
    Dim pvtErrorList As Exceptions
    If errType = 0 Then
        Set pvtErrorList = this.ErrorLog
    Else
        Set pvtErrorList = this.ErrorLog.GetErrorListForType(errType)
    End If

    Dim varError As Variant
    Set varError = Nothing ' VBA Win 64 crashes without this
    For Each varError In pvtErrorList
        Dim pvtError As Exception
        Set pvtError = varError
        Debug.Print pvtError.Description
    Next
End Sub

'@Description("Summary report for exported components.")
Public Sub SummaryReport()
Attribute SummaryReport.VB_Description = "Summary report for exported components."
    Debug.Print "Total files exported to " & this.WorkingDirectory & " : " & this.TotalExported
    Debug.Print "Total warnings: " & this.ErrorLog.TotalWarnings
    Debug.Print "Total failed exports : " & this.ErrorLog.TotalCritical
End Sub


'@Description "Obtains the Rubberduck @Folder annotation from a code module declaration."
Private Function ParseFolderAnnotation(ByVal component As VBComponent) As String
Attribute ParseFolderAnnotation.VB_Description = "Obtains the Rubberduck @Folder annotation from a code module declaration."
    Const RubberduckFolderAnnotation As String = "'@Folder"
    Const annotationStartMatch As String = """"
    Const annotationEndMatch As String = """"

    Dim componentDeclareLines As String
    componentDeclareLines = component.CodeModule.Lines(1, component.CodeModule.CountOfDeclarationLines)
    Dim annotateStart As Long
    annotateStart = InStr(1, componentDeclareLines, RubberduckFolderAnnotation, vbTextCompare)
    If annotateStart > 0 Then
        annotateStart = InStr(annotateStart, componentDeclareLines, annotationStartMatch)
        Dim annotateEnd As Long
        annotateEnd = InStr(annotateStart + 1, componentDeclareLines, annotationEndMatch)
        ParseFolderAnnotation = Mid$(componentDeclareLines, annotateStart + 1, annotateEnd - annotateStart - 1)
    End If
End Function

'@Description "Obtains the path from the Rubberduck @Folder annotation"
Private Function GetFilePathFromFolderAnnotation(ByVal folderAnnotation As String) As String
Attribute GetFilePathFromFolderAnnotation.VB_Description = "Obtains the path from the Rubberduck @Folder annotation"
    Const RubberduckPathSeparator As String = "."
    
    Dim outputPath As String
    If folderAnnotation <> VBA.vbNullString Then
        outputPath = Replace$(folderAnnotation, RubberduckPathSeparator, DirectorySeparatorChar)
    End If
    GetFilePathFromFolderAnnotation = outputPath
End Function

'@Description "Obtains the componet file extension according to its type."
Private Function ComponentExtension(ByVal component As VBComponent) As String
Attribute ComponentExtension.VB_Description = "Obtains the componet file extension according to its type."
    Select Case component.Type
        Case vbext_ComponentType.vbext_ct_StdModule
            ComponentExtension = ".bas"
        Case vbext_ComponentType.vbext_ct_ClassModule
            ComponentExtension = ".cls"
        Case vbext_ComponentType.vbext_ct_MSForm
            ComponentExtension = ".frm"
        Case vbext_ComponentType.vbext_ct_Document
            ComponentExtension = ".doccls"
    End Select
End Function

'@Description "Creates all directories and subdirectories for the specified relative path unless they already exist."
Private Sub CreateRelativeDirectory(ByVal relativePath As String)
Attribute CreateRelativeDirectory.VB_Description = "Creates all directories and subdirectories for the specified relative path unless they already exist."
    Dim folders As Variant
    folders = Split(relativePath, DirectorySeparatorChar)

    Dim folderIndex As Long
    For folderIndex = LBound(folders) To UBound(folders)
        Dim Path As String
        Path = Path & DirectorySeparatorChar & folders(folderIndex)
        Dim fullPath As String
        fullPath = this.WorkingDirectory & Path
        If dir$(fullPath, vbDirectory) = vbNullString Then
            MkDir fullPath
        End If
    Next folderIndex
End Sub

'@Description("Determines whether the given path refers to an existing directory on disk.")
Private Function DirectoryExist(ByVal Path As String) As Boolean
Attribute DirectoryExist.VB_Description = "Determines whether the given path refers to an existing directory on disk."
    If dir$(Path, VBA.vbDirectory) <> vbNullString Then
        DirectoryExist = True
    End If
End Function
