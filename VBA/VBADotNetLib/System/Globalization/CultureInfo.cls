VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CultureInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "CultureInfo factory methods and static members."
'Rubberduck annotations
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("CultureInfo factory methods and static members.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 9, 2023
'@LastModified September 2, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1

'@Remarks
' Singleton class for CultureInfo containing constructors and static members.

Option Explicit

Private Type TCultureInfo
    CultureInfoSingleton As DotNetLib.CultureInfoSingleton
End Type

Private this As TCultureInfo

Private Sub Class_Initialize()
    Set this.CultureInfoSingleton = New DotNetLib.CultureInfoSingleton
End Sub

Private Sub Class_Terminate()
    Set this.CultureInfoSingleton = Nothing
End Sub

'--------------------------------------------------------------------
'Constructors
'--------------------------------------------------------------------

'@Static
'@Description("Initializes a new instance of the CultureInfo class based on the culture specified by the culture identifier and on a value that specifies whether to use the user-selected culture settings from Windows.")
Public Function CreateFromCultureId(ByVal culture As Long, Optional ByVal UseUserOverride As Boolean = False) As DotNetLib.CultureInfo
Attribute CreateFromCultureId.VB_Description = "Initializes a new instance of the CultureInfo class based on the culture specified by the culture identifier and on a value that specifies whether to use the user-selected culture settings from Windows."
    Set CreateFromCultureId = this.CultureInfoSingleton.Create3(culture, UseUserOverride)
End Function

'@Static
'@Description("Initializes a new instance of the CultureInfo class based on the culture specified by name and on a value that specifies whether to use the user-selected culture settings from Windows.")
Public Function CreateFromName(ByVal Name As String, Optional ByVal UseUserOverride As Boolean = False) As DotNetLib.CultureInfo
Attribute CreateFromName.VB_Description = "Initializes a new instance of the CultureInfo class based on the culture specified by name and on a value that specifies whether to use the user-selected culture settings from Windows."
    Set CreateFromName = this.CultureInfoSingleton.Create4(Name, UseUserOverride)
End Function

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------

'@Static
'@Description("Gets or sets the CultureInfo object that represents the culture used by the current thread and task-based asynchronous operations.")
Public Property Get CurrentCulture() As DotNetLib.CultureInfo
Attribute CurrentCulture.VB_Description = "Gets or sets the CultureInfo object that represents the culture used by the current thread and task-based asynchronous operations."
    Set CurrentCulture = this.CultureInfoSingleton.CurrentCulture()
End Property

'@Static
''@Description("Gets or sets the CultureInfo object that represents the culture used by the current thread and task-based asynchronous operations.")
Public Property Set CurrentCulture(ByVal value As DotNetLib.CultureInfo)
    Set this.CultureInfoSingleton.CurrentCulture = value
End Property

'@Static
'@Description "Gets or sets the CultureInfo object that represents the current user interface culture used by the Resource Manager to look up culture-specific resources at run time."
Public Property Get CurrentUICulture() As DotNetLib.CultureInfo
Attribute CurrentUICulture.VB_Description = "Gets or sets the CultureInfo object that represents the current user interface culture used by the Resource Manager to look up culture-specific resources at run time."
    Set CurrentUICulture = this.CultureInfoSingleton.CurrentUICulture()
End Property

'@Static
''@Description("Gets or sets the CultureInfo object that represents the current user interface culture used by the Resource Manager to look up culture-specific resources at run time.")
Public Property Set CurrentUICulture(ByVal value As DotNetLib.CultureInfo)
    Set this.CultureInfoSingleton.CurrentUICulture = value
End Property

'@Static
'@Description("Gets or sets the default culture for threads in the current application domain.")
Public Property Get DefaultThreadCurrentCulture() As DotNetLib.CultureInfo
Attribute DefaultThreadCurrentCulture.VB_Description = "Gets or sets the default culture for threads in the current application domain."
    Set DefaultThreadCurrentCulture = this.CultureInfoSingleton.DefaultThreadCurrentCulture()
End Property

'@Static
''@Description("Gets or sets the default culture for threads in the current application domain.")
Public Property Set DefaultThreadCurrentCulture(ByVal value As DotNetLib.CultureInfo)
    Set this.CultureInfoSingleton.DefaultThreadCurrentCulture = value
End Property

'@Static
'@Description("Gets or sets the default UI culture for threads in the current application domain.")
Public Property Get DefaultThreadCurrentUICulture() As DotNetLib.CultureInfo
Attribute DefaultThreadCurrentUICulture.VB_Description = "Gets or sets the default UI culture for threads in the current application domain."
    Set DefaultThreadCurrentUICulture = this.CultureInfoSingleton.DefaultThreadCurrentUICulture()
End Property

'@Static
''@Description("Gets or sets the default UI culture for threads in the current application domain.")
Public Property Set DefaultThreadCurrentUICulture(ByVal value As DotNetLib.CultureInfo)
    Set this.CultureInfoSingleton.DefaultThreadCurrentUICulture = value
End Property

'@Static
'@Description("Gets the CultureInfo that represents the culture installed with the operating system.")
Public Property Get InstalledUICulture() As DotNetLib.CultureInfo
Attribute InstalledUICulture.VB_Description = "Gets the CultureInfo that represents the culture installed with the operating system."
    Set InstalledUICulture = this.CultureInfoSingleton.InstalledUICulture()
End Property

'@Static
'@Description("Gets the CultureInfo object that is culture-independent(invariant).")
Public Property Get InvariantCulture() As DotNetLib.CultureInfo
Attribute InvariantCulture.VB_Description = "Gets the CultureInfo object that is culture-independent(invariant)."
    Set InvariantCulture = this.CultureInfoSingleton.InvariantCulture()
End Property

'@Static
'@Description("Creates a CultureInfo that represents the specific culture that is associated with the specified name.")
Public Function CreateSpecificCulture(ByVal Name As String) As DotNetLib.CultureInfo
Attribute CreateSpecificCulture.VB_Description = "Creates a CultureInfo that represents the specific culture that is associated with the specified name."
    Set CreateSpecificCulture = this.CultureInfoSingleton.CreateSpecificCulture(Name)
End Function

'@Static
'@Description("Retrieves a cached, read-only instance of a culture by using the specified culture identifier.")
Public Function GetCultureInfo(ByVal culture As Long) As DotNetLib.CultureInfo
Attribute GetCultureInfo.VB_Description = "Retrieves a cached, read-only instance of a culture by using the specified culture identifier."
    Set GetCultureInfo = this.CultureInfoSingleton.GetCultureInfo(culture)
End Function

'@Static
'@Description("Retrieves a cached, read-only instance of a culture using the specified culture name.")
Public Function GetCultureInfo2(ByVal Name As String) As DotNetLib.CultureInfo
Attribute GetCultureInfo2.VB_Description = "Retrieves a cached, read-only instance of a culture using the specified culture name."
    Set GetCultureInfo2 = this.CultureInfoSingleton.GetCultureInfo2(Name)
End Function

'@Static
'@Description("Retrieves a cached, read-only instance of a culture. Parameters specify a culture that is initialized with the TextInfo and CompareInfo objects specified by another culture.")
Public Function GetCultureInfo3(ByVal Name As String, ByVal altName As String) As DotNetLib.CultureInfo
Attribute GetCultureInfo3.VB_Description = "Retrieves a cached, read-only instance of a culture. Parameters specify a culture that is initialized with the TextInfo and CompareInfo objects specified by another culture."
    Set GetCultureInfo3 = this.CultureInfoSingleton.GetCultureInfo3(Name, altName)
End Function

'@Deprecated
'@Static
'@Description("Deprecated. Retrieves a read-only CultureInfo object having linguistic characteristics that are identified by the specified RFC 4646 language tag.")
Public Function GetCultureInfoByIetfLanguageTag(ByVal Name As String) As DotNetLib.CultureInfo
Attribute GetCultureInfoByIetfLanguageTag.VB_Description = "Deprecated. Retrieves a read-only CultureInfo object having linguistic characteristics that are identified by the specified RFC 4646 language tag."
    Set GetCultureInfoByIetfLanguageTag = this.CultureInfoSingleton.GetCultureInfoByIetfLanguageTag(Name)
End Function

'@Description("Gets the list of supported cultures filtered by the specified CultureTypes parameter.")
Public Function GetCultures(ByVal Types As CultureTypes) As DotNetLib.CultureInfo()
Attribute GetCultures.VB_Description = "Gets the list of supported cultures filtered by the specified CultureTypes parameter."
    GetCultures = this.CultureInfoSingleton.GetCultures(Types)
End Function

'@Description("Returns a read-only wrapper around the specified CultureInfo object.")
Public Function ReadOnly(ByVal ci As DotNetLib.CultureInfo) As DotNetLib.CultureInfo
Attribute ReadOnly.VB_Description = "Returns a read-only wrapper around the specified CultureInfo object."
    Set ReadOnly = this.CultureInfoSingleton.ReadOnly(ci)
End Function
