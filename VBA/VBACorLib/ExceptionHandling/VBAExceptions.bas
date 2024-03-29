Attribute VB_Name = "VBAExceptions"
'@Folder "VBACorLib.ExceptionHandling"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified August 24, 2023

Option Explicit

'https://github.com/cocus/openmsvbvm/blob/a54ed4686c55d510cc55b5b46cd3f1d44673322f/openmsvbvm/vba_exception.h
Public Enum VBAException
    NO_ERROR = 0
    RETURN_WITHOUT_GOSUB = 3                        ' Return without GoSub
    INVALID_PROCEDURE_CALL = 5                      ' Invalid procedure call
    OVERFLOW = 6                                    ' Overflow
    OUT_OF_MEMORY = 7                               ' Out of memory
    SUBSCRIPT_OUT_OF_RANGE = 9                      ' Subscript out of range
    ARRAY_FIXED_OR_TEMPORARILY_LOCKED = 10          ' This array is fixed or temporarily locked
    DIVISION_BY_ZERO = 11                           ' Division by zero
    TYPE_MISMATCH = 13                              ' Type mismatch
    OUT_OF_STRING_SPACE = 14                        ' Out of string space
    EXPRESSION_TOO_COMPLEX = 16                     ' Expression too complex
    CANT_PERFORM_REQUESTED_OPERATION = 17           ' Can't perform requested operation
    USER_INTERRUPT_OCCURRED = 18                    ' User interrupt occurred
    RESUME_WITHOUT_ERROR = 20                       ' Resume without error
    OUT_OF_STACK_SPACE = 28                         ' Out of stack space
    SUB_FUNCTION_OR_PROPERTY_NOT_DEFINED = 35       ' Sub 'Function 'or Property not defined
    TOO_MANY_DLL_APPLICATION_CLIENTS = 47           ' Too many DLL application clients
    ERROR_IN_LOADING_DLL = 48                       ' Error in loading DLL
    BAD_DLL_CALLING_CONVENTION = 49                 ' Bad DLL calling convention
    INTERNAL_ERROR = 51                             ' Internal error
    BAD_FILENAME_OR_NUMBER = 52                     ' Bad file name or number
    FILE_NOT_FOUND = 53                             ' File not found
    BAD_FILE_MODE = 54                              ' Bad file mode
    FILE_ALREADY_OPEN = 55                          ' File already open
    DEVICE_IO_ERROR = 57                            ' Device I/O error
    FILE_ALREADY_EXISTS = 58                        ' File already exists
    BAD_RECORD_LENGTH = 59                          ' Bad record length
    DISK_FULL = 61                                  ' Disk full
    INPUT_PAST_END_OF_FILE = 62                     ' Input past end of file
    BAD_RECORD_NUMBER = 63                          ' Bad record number
    TOO_MANY_FILES = 67                             ' Too many files
    DEVICE_UNAVAILABLE = 68                         ' Device unavailable
    PERMISSION_DENIED = 70                          ' Permission denied
    DISK_NOT_READY = 71                             ' Disk not ready
    CANT_RENAME_WITH_DIFFERENT_DRIVE = 74           ' Can't rename with different drive
    PATH_FILE_ACCESS_ERROR = 75                     ' Path/File access error
    PATH_NOT_FOUND = 76                             ' Path not found
    OBJECT_VARIABLE_OR_WITH_BLOCK_VARIABLE_NOT_SET = 91 ' Object variable or With block variable not set
    FOR_LOOP_NOT_INITIALIZED = 92                   ' For loop not initialized
    INVALID_PATTERN_STRING = 93                     ' Invalid pattern string
    INVALID_USE_OF_NULL = 94                        ' Invalid use of Null
    CANT_CALL_FRIEND_PROCEDURE = 97                 ' Can't call Friend procedure on an object that is not an instance of the defining class
    PROPERTY_OR_METHOD_CALL_INCLUDING_REFERENCE_TO_PRIVATE_OBJECT = 98 ' A property or method call cannot include a reference to a private object 'either as an argument or as a return value
    SYSTEM_DLL_COULD_NOT_BE_LOADED = 298            ' System DLL could not be loaded
    CANT_USE_CHARACTER_DEVICES_IN_SPECIFIED_FILE_NAMES = 320 ' Can't use character device names in specified file names
    INVALID_FILE_FORMAT = 321                       ' Invalid file format
    CANT_CREATE_NECESSARY_TEMPORARY_FILE = 322      ' Can�t create necessary temporary file
    INVALID_FORMAT_IN_RESOURCE_FILE = 325           ' Invalid format in resource file
    DATA_VALUE_NAMED_NOT_FOUND = 327                ' Data value named not found
    ILLEGAL_PARAMETER_CANT_WRITE_ARRAYS = 328       ' Illegal parameter; can't write arrays
    COULD_NOT_ACCESS_SYSTEM_REGISTRY = 335          ' Could not access system registry
    COMPONENT_NOT_CORRECTLY_REGISTRED = 336         ' Component not correctly registered
    COMPONENT_NOT_FOUND = 337                       ' Component not found
    COMPONENT_DID_NOT_RUN_CORRECTLY = 338           ' Component did not run correctly
    OBJECT_ALREADY_LOADED = 360                     ' Object already loaded
    CANT_LOAD_OR_UNLOAD_THIS_OBJECT = 361           ' Can't load or unload this object
    CONTROL_SPECIFIED_NOT_FOUND = 363               ' Control specified not found
    OBJECT_WAS_UNLOADED = 364                       ' Object was unloaded
    UNABLE_TO_UNLOAD_WITHIN_THIS_CONTEXT = 365      ' Unable to unload within this context
    SPECIFIED_FILE_OUT_OF_DATE_PROGRAM_REQUIRES_LATER_VERSION = 368 ' The specified file is out of date. This program requires a later version
    SPECIFIED_OBJECT_CANT_BE_USED_AS_OWNER_FORM_FOR_SHOW = 371 ' The specified object can't be used as an owner form for Show
    INVALID_PROPERTY_VALUE = 380                    ' Invalid property value
    INVALID_PROPERTY_ARRAY_INDEX = 381              ' Invalid property-array index
    PROPERTY_SET_CANT_BE_EXECUTED_AT_RUN_TIME = 382 ' Property Set can't be executed at run time
    PROPERTY_SET_CANT_BE_USED_WITH_A_READ_ONLY_PROPERTY = 383 ' Property Set can't be used with a read-only property
    NEED_PROPERTY_ARRAY_INDEX = 385                 ' Need property-array index
    PROPERTY_SET_NOT_PERMITTED = 387                ' Property Set not permitted
    PROPERTY_GET_CANT_BE_EXECUTED_AT_RUN_TIME = 393 ' Property Get can't be executed at run time
    PROPERTY_GET_CANT_BE_EXECUTED_ON_WRITE_ONLY_PROPERTY = 394 ' Property Get can't be executed on write-only property
    FORM_ALREADY_DISPLAYED_CANT_SHOW_MODALLY = 400  ' Form already displayed; can't show modally
    CODE_MUST_CLOSE_TOPMOST_MODAL_FORM_FIRST = 402  ' Code must close topmost modal form first
    PERMISSION_TO_USE_OBJECT_DENIED = 419           ' Permission to use object denied
    PROPERTY_NOT_FOUND = 422                        ' Property not found
    PROPERTY_OR_METHOD_NOT_FOUND = 423              ' Property or method not found
    OBJECT_REQUIRED = 424                           ' Object required
    INVALID_OBJECT_USE = 425                        ' Invalid object use
    COMPONENT_CANT_CREATE_OBJECT_OR_RETURN_REFERENCE_TO_THIS_OBJECT = 429 ' Component can't create object or return reference to this object
    CLASS_DOESNT_SUPPORT_AUTOMATION = 430           ' Class doesn't support Automation
    FILE_NAME_OR_CLASS_NAME_NOT_FOUND_DURING_AUTOMATION_OPERATION = 432 ' File name or class name not found during Automation operation
    OBJECT_DOESNT_SUPPORT_THIS_PROPERTY_OR_METHOD = 438 ' Object doesn't support this property or method
    AUTOMATION_ERROR = 440                          ' Automation error
    CONNECTION_TO_TYPE_LIBRARY_OR_OBJECT_LIBRARY_FOR_REMOTE_PROCESS_HAS_BEEN_LOST = 442 ' Connection to type library or object library for remote process has been lost
    AUTOMATION_DOESNT_HAVE_A_DEFAULT_VALUE = 443    ' Automation object doesn't have a default value
    OBJECT_DOESNT_SUPPORT_THIS_ACTION = 445         ' Object doesn't support this action
    OBJECT_DOESNT_SUPPORT_NAMED_ARGUMENTS = 446     ' Object doesn't support named arguments
    OBJECT_DOESNT_SUPPORT_CURRENT_LOCALE_SETTING = 447 ' Object doesn't support current locale setting
    NAMED_ARGUMENT_NOT_FOUND = 448                  ' Named argument not found
    ARGUMENT_NOT_OPTIONAL_OR_INVALID_PROPERTY_ASSIGNMENT = 449 ' Argument not optional or invalid property assignment
    WRONG_NUMBER_OF_ARGUMENTS_OR_INVALID_PROPERTY_ASSIGNMENT = 450 ' Wrong number of arguments or invalid property assignment
    OBJECT_NOT_A_COLLECTION = 451                   ' Object not a collection
    INVALID_ORDINAL = 452                           ' Invalid ordinal
    SPECIFIED_NOT_FOUND = 453                       ' Specified not found
    CODE_RESOURCE_NOT_FOUND = 454                   ' Code resource not found
    CODE_RESOURCE_LOCK_ERROR = 455                  ' Code resource lock error
    THIS_KEY_IS_ALREADY_ASSOCIATED_WITH_AN_ELEMENT_OF_THIS_COLLECTION = 457 ' This key is already associated with an element of this collection
    VARIABLE_USES_A_TYPE_NOT_SUPPORTED_IN_VISUAL_BASIC = 458    ' Variable uses a type not supported in Visual Basic
    THIS_COMPONENT_DOESNT_SUPPORT_THE_SET_OF_EVENTS = 459       ' This component doesn't support the set of events
    INVALID_CLIPBOARD_FORMAT = 460                  ' Invalid Clipboard format
    METHOD_OR_DATA_MEMBER_NOT_FOUND = 461           ' Method or data member not found
    THE_REMOTE_SERVER_MACHINE_DOES_NOT_EXIST_OR_IS_UNAVAILABLE = 462 ' The remote server machine does not exist or is unavailable
    CLASS_NOT_REGISTREED_ON_LOCAL_MACHINE = 463     ' Class not registered on local machine
    CANT_CREATE_AUTOREDRAW_IMAGE = 480              ' Can't create AutoRedraw image
    INVALID_PICTURE = 481                           ' Invalid picture
    PRINTER_ERROR = 482                             ' Printer error
    PRINTER_DRIVER_DOES_NOT_SUPPORT_SPECIFIED_PROPERTY = 483    ' Printer driver does not support specified property
    PROBLEM_GETTING_PRINTER_INFORMATION_FROM_SYSTEM = 484       ' Problem getting printer information from the system. Make sure the printer is set up correctly
    INVALID_PICTURE_TYPE = 485                      ' Invalid picture type
    CANT_PRINT_FORM_IMAGE_TO_THIS_TYPE_OF_PRINTER = 486     ' Can't print form image to this type of printer
    CANT_EMPTY_CLIPBOARD = 520                      ' Can't empty Clipboard
    CANT_OPEN_CLIPBOARD = 521                       ' Can't open Clipboard
    CANT_SAVE_FILE_TO_TEMP_DIRECTORY = 735          ' Can't save file to TEMP directory
    SEARCH_TEXT_NOT_fOUND = 744                     ' Search text not found
    REPLACEMENTS_TOO_LONG = 746                     ' Replacements too long
    OUT_OF_MEMORY2 = 31001                          ' Out of memory
    NO_OBJECT = 31004                               ' No object
    CLASS_IS_NOT_SET = 31018                        ' Class is not set
    UNABLE_TO_ACTIVATE_OBJECT = 31027               ' Unable to activate object
    UNABLE_TO_CREATE_EMBEDDED_OBJECT = 31032        ' Unable to create embedded object
    ERROR_SAVING_TO_FILE = 31036                    ' Error saving to file
    ERROR_LOADING_FROM_FILE = 31037                 ' Error loading from file
End Enum

