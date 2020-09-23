VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLstVwExportImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' SUPPORTS V6 Common Controls only!!!

' -----------------------------------------------------------------------------------------------------------------------------------------------------------------
' These two events are very important for Importing. Respond to them
' So, always declare this class WithEvents

Public Event ImportSummary(ByVal ColumnHeadersIncluded As Boolean, ByVal ImageListsNeeded As Long, ByVal ListItems As Boolean, ByRef Continue As Boolean)

' This event is called to allow you to change any class properties and/or allow/prevent some things from being imported

' ColumnHeadersIncluded: Indicates whether or not the import file includes column headers.
'   :: Here you should either allow or prevent column headers from being imported if the parameter is True
'   :: To Allow
'       ::: Ensure the class' IncludeColumnHeaders property is True
'       ::: As a side note, set the class' IncludeHeaderTags property as desired
'   :: To Prevent
'       ::: Ensure the class' IncludeColumnHeaders property is False
'   :: note... If no headers are provided or you prevent them from being imported, you should manually add them
'               if the listview's View property is Report; otherwise, you may not see the items even though they exist

' ImageListsNeeded: If non-zero, ImageList(s) exist in the file and it is assumed that columnheaders and/or listitems use the imagelists
'   :: Here you should either allow or prevent imagelists from being imported
'   :: To Allow
'       ::: Ensure the class' IncludeImageLists property is True
'       ::: Reply to the SetImageList event and pass the ImportToImageList parameter when called
'           -- At most, 3 image lists may be required. These can already exist or be created/destroyed dynamically
'           -- Ensure any passed ImageList is not referenced by any other controls, else no images will be imported
'           -- Images are never appended, the ImageList control is cleared and the exported images are uploaded
'           -- The SetImageList event can be called up to 3 times
'   :: To Prevent
'       ::: Ensure the class' IncludeImageLists property is False
'       ::: You may also reply to the SetImageList event and pass Nothing as the ImportToImageList parameter
'       ::: If your ListView will be using an existing/populated imagelist control, ensure you set its references

' ListItems: FYI only, indicates how many list items exist in the file
'   :: Here you should decide whether to clear the listview or not.
'   :: If not cleared, the listitems will be appended to the listview
'       ::: As a side note, set the class' IncludeListItemTags & IncludeTextFormatting properties as desired

' Continue: Passed as False. Set to True to allow import to continue else importing is aborted

Public Event SetImageList(ByVal ListViewSection As lvImportImageList, ByRef ImportToImageList As MSComctlLib.ImageList)
' Called only when these 2 cases are true
'   1. The file to import contains ImageList(s)
'   2. This class' IncludeImageLists property is True
'   :: When called, it will only be called for each unique imagelist control required.
'       -- If the SmallIcon and ColumnHeaderIcons ImageList references for the ListView are the same ImageList,
'           then this is called only one time instead of twice
'       -- This can be called from 1 to 3 times depending on how many unique ImageLists exist in the file to be imported
'   :: If this event is not called (you prevented it from being called or there was no image list exported), you
'       should manually set the listview's imagelist references if needed

' :: When Exporting, these properties should be set as desired BEFORE calling the ExportToFile function
' IncludeControlFormatting: Exports the ListView control's physical appearance both client & non-client area
' IncludeTextFormatting: Exports list item text forecolor & boldness
' IncludeColumnHeaders: Exports column headers and all their properties
' IncludeHeaderTags: Exports the column header tags, if any
' IncludeImageLists: Exports each ImageList control used by the ListView
' IncludeListItemTags: Exports each ListView ListItem & SubItem tag, if any

' Notes
' 1. The MsgBox calls in ImportFromFile & ExportToFile are more for debugging purposes. You may want to disable those lines & rely on the return function
' -----------------------------------------------------------------------------------------------------------------------------------------------------------------



' Kernel32/User32 APIs for Unicode Filename Support
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function IsWindowUnicode Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, Optional ByVal lpOverlapped As Long = 0&) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_NORMAL = &H80&

' used for workaround of VB not exposing IStream interface
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal ppv As Long, ByVal oVft As Long, ByVal cc As Long, ByVal rtTYP As VbVarType, ByVal paCNT As Long, ByVal paTypes As Long, ByVal paValues As Long, ByRef fuReturn As Variant) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, pPStm As Any) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal pPStm As Long, hGlobal As Long) As Long
Private Declare Function OleSaveToStream Lib "ole32.dll" (ByVal pPStm As Long, ByVal pStm As stdole.IUnknown) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' These structures are used to help organize code & make code more user-friendly
Private Type StructureFileHeader
    MagicNumber As Long                 ' identifier for this type of file
    Version As Single                    ' version (integer=Ansi, decimal=Unicode)
    ImageListOffset As Long
    HeadersOffset As Long               ' position in file where column headers begin
    ListItemOffset As Long              ' position in file where listitems begin
    Columns As Long                     ' number of column headers
    Items As Long                       ' how many list items
    ImageLists As Long                  ' number of exported listimages
    Flags As Long                       ' whether headers/imagelists/tags are written to file
    ListViewStructSize As Long          ' size of the StructureListview used
End Type                                ' 36 bytes used for import/export
Private Type StructureListView
    AllowColumnReorder As Byte
    AppearanceFlat As Byte
    BorderStyleNone As Byte
    Checkboxes As Byte
    FlatScrollBar As Byte
    FullRowSelect As Byte
    GridLines As Byte
    HideColumnHeaders As Byte
    HotTracking As Byte
    HoverSelection As Byte
    LabelEditManual As Byte
    LabelWrap As Byte
    MultiSelect As Byte
    Sorted As Byte
    SortOrderDesc As Byte
    TextBkgOpaque As Byte               ' 16 bytes
    Arrange As Integer
    PictureAlignment As Integer
    SortKey As Integer
    View As Integer                     ' 8 bytes
    BackColor As Long
    ForeColor As Long                   ' 8 bytes
    Tag As String                       ' 4 bytes (not used for export, used for import)
End Type                                ' 32 bytes used for export
Private Type StructureHeaders
    Position As Long
    Left As Single
    Width As Single                     ' 12 bytes
    Alignment As Integer
    TagType As Integer
    IconType As Integer                 ' 6 bytes (18 bytes used for export)
    Text As String                      ' 4 bytes (following used during import only)
    Key As String                       ' 4 bytes
    Tag As Variant                      ' 16 bytes
    Icon As Variant                     ' 16 bytes
End Type
Private Type StructureImageList
    Count As Long
    BackColor As Long
    MaskColor As Long                   ' 12 bytes
    ImageWidth As Integer
    ImageHeight As Integer
    UseMaskColor As Boolean             ' 6 bytes
    Tag As String                       ' 4 bytes
End Type                                ' 18 bytes used for export
Private Type StructureListItem
    SubCount As Long
    ForeColor As Long
    IconTypes As Long                   ' 12 bytes
    TagType As Integer
    BoldChecked As Integer              ' 4 bytes
    Text As String
    Tooltip As String
    Key As String
    IconSm As Variant
    IconLg As Variant
    Tag As Variant                      ' 16 bytes used for export, all used for import
End Type
Private Type StructureListSubItem
    ForeColor As Long                   ' 4 bytes
    IconType As Integer
    TagType As Integer                  ' 4 bytes
    Bold As Byte                        ' 1 byte
    Text As String
    Tooltip As String
    Key As String
    ReportIcon As Variant
    Tag As Variant                      ' 9 bytes used for export, all used for import
End Type
Private Type StructureFont
    Attributes As Long              ' bold,italic,underline,strikethru
    CharSet As Integer
    Weight As Integer
    Size As Currency                ' 16 bytes used for export, all used for import
    Name As String
End Type

Public Enum lvImportImageList
    ilColumnHeaders = 1
    ilLargeIcons = 2
    ilSmallIcons = 3
End Enum
Private Enum lvFlags
    fgHasImageList = 1
    fgHasHeaderTags = 2
    fgHasItemTags = 4
    fgHasHeaders = 8
    fgHasCtrlFormat = 16
    fgHasTextFormat = 32
End Enum

Private c_Flags As Long
Private lvObject As MSComctlLib.ListView

Public Property Let IncludeImageLists(ByVal bInclude As Boolean)
    ' initialized to False during class load
    ' Exporting: If True, will cache the imagelist controls used by the listview to the same file, including all images else will not
    ' Importing: If True, will overwrite existing imagelist controls (see 3 ImageList properties below) control & repopulate with the one saved in the file
    '   :: If no imagelist was saved in the file, nothing will be done the imagelists
    '   :: If property is False, no images will be extracted from the file
    If bInclude <> Me.IncludeImageLists Then c_Flags = c_Flags Xor fgHasImageList
End Property
Public Property Get IncludeImageLists() As Boolean
    IncludeImageLists = CBool(c_Flags And fgHasImageList)
End Property
    
Public Property Let IncludeColumnHeaders(ByVal bInclude As Boolean)
    ' initialized to True during class load
    ' Exporting: If True, will include column headers and their individual properties else will not
    ' Importing: If True, will clear the listiview and repopulate the headers from the file
    '   :: If no column headers were saved in the file, then nothing will be done to the listview's original header columns
    '   :: If property is False, no column headers will be extracted from the file
    If bInclude <> Me.IncludeColumnHeaders Then c_Flags = c_Flags Xor fgHasHeaders
End Property
Public Property Get IncludeColumnHeaders() As Boolean
    IncludeColumnHeaders = CBool(c_Flags And fgHasHeaders)
End Property
    
Public Property Let IncludeHeaderTags(ByVal bInclude As Boolean)
    ' initialized to False during class load
    ' Exporting: If True, column header tags will be exported to the file else they will not
    ' Importing: Applies only if IncludeColumnHeaders is True
    '   :: If True, exported column header tags will be added to new column headers else null tags will be set
    '   :: If the file was exported without column header tags, no action is taken
    If bInclude <> Me.IncludeHeaderTags Then c_Flags = c_Flags Xor fgHasHeaderTags
End Property
Public Property Get IncludeHeaderTags() As Boolean
    IncludeHeaderTags = CBool(c_Flags And fgHasHeaderTags)
End Property
    
Public Property Let IncludeListItemTags(ByVal bInclude As Boolean)
    ' initialized to False during class load
    ' Exporting: If True, list item tags & subItem tags will be exported to the file else they will not
    ' Importing: If True, exported tags will be added to listItems & subItems else null tags will be set
    '   :: if the file was exported without tags, no action is taken
    If bInclude <> Me.IncludeListItemTags Then c_Flags = c_Flags Xor fgHasItemTags
End Property
Public Property Get IncludeListItemTags() As Boolean
    IncludeListItemTags = CBool(c_Flags And fgHasItemTags)
End Property

Public Property Let IncludeControlFormatting(ByVal bInclude As Boolean)
    ' initialized to True during class load
    ' Exporting: If True, the physical appearance of the listview control will be exported (borders, colors, font, picture, etc)
    ' Importing: If True, the imported listview control will be changed to match the exported appearance
    '   :: if the file was exported without the appearance, no action is taken
    If bInclude <> Me.IncludeControlFormatting Then c_Flags = c_Flags Xor fgHasCtrlFormat
End Property
Public Property Get IncludeControlFormatting() As Boolean
    IncludeControlFormatting = CBool(c_Flags And fgHasCtrlFormat)
End Property

Public Property Let IncludeTextFormatting(ByVal bInclude As Boolean)
    ' initialized to True during class load
    ' Exporting: If True, the physical appearance of the listview items will be exported (colors, boldness)
    ' Importing: If True, the imported listitems will be changed to match the exported appearance
    '   :: if the file was exported without the appearance, no action is taken
    If bInclude <> Me.IncludeTextFormatting Then c_Flags = c_Flags Xor fgHasTextFormat
End Property
Public Property Get IncludeTextFormatting() As Boolean
    IncludeTextFormatting = CBool(c_Flags And fgHasTextFormat)
End Property

Public Property Get ListViewReference() As MSComctlLib.ListView
    ' ReadOnly property. Can be called during one of the public events if desired
    ' to see which listview is being imported into
    Set ListViewReference = lvObject
End Property

Public Function ExportToFile(ListView As MSComctlLib.ListView, FileName As String, Optional SaveStringsANSI As Boolean = True) As Boolean

    ' Parameter SaveStringsANSI
    ' :: can reduce filesize, but is slower overall
    ' :: if using unicode strings for any part of the listview, set this to False

    If ListView Is Nothing Then Exit Function
    If FileName = vbNullString Then Exit Function
    
    Dim fileNum As Long, useUnicode As Boolean, bSuccess As Boolean
    Dim offsetHDR As Long, offsetLI As Long, offsetIL As Long
    Dim vWritten As Long, imgListCount As Long
    If IsWindowUnicode(GetDesktopWindow) Then useUnicode = True
    
    On Error GoTo EH
    fileNum = GetFileHandle(FileName, False, useUnicode)
    
    Set lvObject = ListView
    If WriteObjAttributes(fileNum, SaveStringsANSI, imgListCount) Then
        offsetIL = SetFilePointer(fileNum, 0&, 0&, 1&)
        If Me.IncludeImageLists = True And imgListCount > 0& Then
            bSuccess = WriteImageLists(fileNum, SaveStringsANSI)
        Else
            bSuccess = True
        End If
        If bSuccess Then
            offsetHDR = SetFilePointer(fileNum, 0&, 0&, 1&)
            If Me.IncludeColumnHeaders = True And lvObject.ColumnHeaders.Count > 0& Then
                bSuccess = WriteHeaders(fileNum, SaveStringsANSI)
            Else
                bSuccess = True
            End If
            If bSuccess Then
                offsetLI = SetFilePointer(fileNum, 0&, 0&, 1&)
                If lvObject.ListItems.Count > 0& Then
                    bSuccess = WriteListItems(fileNum, SaveStringsANSI)
                Else
                    bSuccess = True
                End If
                If bSuccess Then
                    SetFilePointer fileNum, 8&, 0&, 0&
                    WriteFile fileNum, offsetIL, 4&, vWritten
                    WriteFile fileNum, offsetHDR, 4&, vWritten
                    WriteFile fileNum, offsetLI, 4&, vWritten
                End If
            End If
        End If
    End If
EH:
    If Not fileNum = INVALID_HANDLE_VALUE Then CloseHandle fileNum
    If Err Then
        MsgBox Err.Description, vbExclamation + vbOKOnly, "Failed to Export ListView Contents"
        Err.Clear
    ElseIf bSuccess Then
        ExportToFile = bSuccess
    Else
        MsgBox "Failed to Export the ListView Contents", vbExclamation + vbOKOnly, "Error"
    End If
    Set lvObject = Nothing
    
End Function

Public Function ImportFromFile(ListView As MSComctlLib.ListView, ByVal FileName As String) As Boolean

    If ListView Is Nothing Then Exit Function
    If FileName = vbNullString Then Exit Function
    
    Dim fileNum As Long, lLoc As Long
    Dim bSuccess As Boolean, useUnicode As Boolean
    Dim udtFile As StructureFileHeader
    
    If IsWindowUnicode(GetDesktopWindow) Then useUnicode = True
    fileNum = GetFileHandle(FileName, True, useUnicode)
    If fileNum = INVALID_HANDLE_VALUE Then
        If FileExists(FileName, useUnicode) Then
            MsgBox "Cannot access that file", vbExclamation + vbOKOnly, "Failed To Import ListView"
        Else
            MsgBox "The file: " & FileName & " does not exist", vbExclamation + vbOKOnly, "Failed to Import ListView"
        End If
        Exit Function
    End If
    
    On Error GoTo EH
    If ReadFileHeader(fileNum, udtFile) Then
        ' prompt user with enough info to modify destination listview, set class properties, create imagelists, etc
        RaiseEvent ImportSummary((udtFile.Columns > 0&), udtFile.ImageLists, udtFile.Items, bSuccess)
        
        If bSuccess Then
            Set lvObject = ListView
            If Me.IncludeControlFormatting = True And ((udtFile.Flags And fgHasCtrlFormat) = fgHasCtrlFormat) Then
                bSuccess = ReadObjAttributes(fileNum, udtFile)
            Else
                bSuccess = True
            End If
            If bSuccess Then
                If Me.IncludeImageLists = True And udtFile.ImageLists > 0& Then
                    lLoc = SetFilePointer(fileNum, 0&, 0&, 1&)
                    SetFilePointer fileNum, udtFile.ImageListOffset - lLoc, 0&, 1&
                    bSuccess = ReadImageLists(fileNum, udtFile)
                Else
                    bSuccess = True
                End If
                If bSuccess Then
                    If udtFile.Columns > 0& And Me.IncludeColumnHeaders = True Then
                        lLoc = SetFilePointer(fileNum, 0&, 0&, 1&)
                        SetFilePointer fileNum, udtFile.HeadersOffset - lLoc, 0&, 1&
                        bSuccess = ReadHeaders(fileNum, udtFile)
                    Else
                        bSuccess = True
                    End If
                    If bSuccess Then
                        If udtFile.Items > 0& Then
                            lLoc = SetFilePointer(fileNum, 0&, 0&, 1&)
                            SetFilePointer fileNum, udtFile.ListItemOffset - lLoc, 0&, 1&
                            bSuccess = ReadListItems(fileNum, udtFile)
                        Else
                            bSuccess = True
                        End If
                        If (udtFile.Flags And 1024&) Then lvObject.Sorted = True
                    End If
                End If
            End If
        End If
    End If
EH:
    CloseHandle fileNum
    Set lvObject = Nothing
    If Err Then Err.Clear
    ImportFromFile = bSuccess
    If Not bSuccess Then MsgBox "Failed to Import Listview Contents", vbExclamation + vbOKOnly, "Failure"
End Function

Private Function WriteObjAttributes(fileNum As Long, SaveStringsANSI As Boolean, imgListExported As Long) As Boolean
    
    ' Purpose: Write the overall listview attributes/properties
    ' File is written/read in the following order
    ' :: File Header
    ' :: Font
    ' :: Picture Property
    ' :: Tag property
    ' :: Control Properties
    
    Dim udtAttr As StructureListView
    Dim udtFont As StructureFont
    Dim udtFile As StructureFileHeader
    Dim vWritten As Long
    
    On Error GoTo EH
    With udtFile
        .MagicNumber = 1447843404
        If SaveStringsANSI Then .Version = 1! Else .Version = 1.1!
        If Me.IncludeColumnHeaders Then .Columns = lvObject.ColumnHeaders.Count
        .Items = lvObject.ListItems.Count
        .Flags = c_Flags
        .ListViewStructSize = 32&
        If Me.IncludeImageLists Then
            If Not lvObject.ColumnHeaderIcons Is Nothing Then .ImageLists = 1&
            If Not lvObject.Icons Is Nothing Then
                If Not lvObject.Icons Is lvObject.ColumnHeaderIcons Then .ImageLists = .ImageLists + 1&
            End If
            If Not lvObject.SmallIcons Is Nothing Then
                If Not lvObject.SmallIcons Is lvObject.Icons Then
                    If Not lvObject.SmallIcons Is lvObject.ColumnHeaderIcons Then .ImageLists = .ImageLists + 1&
                End If
            End If
        End If
        imgListExported = .ImageLists
    End With
    SetFilePointer fileNum, 0&, 0&, 0&
    WriteFile fileNum, udtFile, 40&, vWritten
    If vWritten = 40& Then
        
        If Me.IncludeControlFormatting Then
            On Error Resume Next
            ' extract the listview properties; not all listview versions will have same properties,
            ' therefore, we use the On Error statement
            With udtAttr
                .AllowColumnReorder = Abs(lvObject.AllowColumnReorder)
                .AppearanceFlat = Abs(lvObject.Appearance = ccFlat)
                .Arrange = lvObject.Arrange
                .BackColor = lvObject.BackColor
                .BorderStyleNone = Abs(lvObject.BorderStyle = ccNone)
                .Checkboxes = Abs(lvObject.Checkboxes)
                .FlatScrollBar = Abs(lvObject.FlatScrollBar)
                .ForeColor = lvObject.ForeColor
                .FullRowSelect = Abs(lvObject.FullRowSelect)
                .GridLines = Abs(lvObject.GridLines)
                .HideColumnHeaders = Abs(lvObject.HideColumnHeaders)
                .HotTracking = Abs(lvObject.HotTracking)
                .HoverSelection = Abs(lvObject.HoverSelection)
                .LabelEditManual = Abs(lvObject.LabelEdit = lvwManual)
                .LabelWrap = Abs(lvObject.LabelWrap)
                .MultiSelect = Abs(lvObject.MultiSelect)
                .PictureAlignment = lvObject.PictureAlignment
                .Sorted = Abs(lvObject.Sorted)
                .SortKey = lvObject.SortKey
                .SortOrderDesc = Abs(lvObject.SortOrder = lvwDescending)
                .TextBkgOpaque = Abs(lvObject.TextBackground = lvwOpaque)
                .View = lvObject.View
            End With
            ' start writing everything
            ' if an error occurs here, we abort the routine
            On Error GoTo EH
            With lvObject.Font       ' write the font used
                udtFont.Attributes = Abs(.Bold = True)
                udtFont.Attributes = udtFont.Attributes Or Abs(.Italic) * &H2
                udtFont.Attributes = udtFont.Attributes Or Abs(.Strikethrough) * &H4
                udtFont.Attributes = udtFont.Attributes Or Abs(.Underline) * &H8
                udtFont.CharSet = .CharSet
                udtFont.Size = .Size
                udtFont.Weight = .Weight
            End With
            WriteFile fileNum, udtFont, 16&, vWritten
            If vWritten = 16& Then  ' write the font name
                If WriteString(fileNum, lvObject.Font.Name, False) Then
                    ' save the picture property & .tag property
                     If WritePicture(fileNum, lvObject.Picture) Then
                        If WriteString(fileNum, lvObject.Tag, SaveStringsANSI) Then
                            ' write the listview attributes
                            WriteFile fileNum, udtAttr, udtFile.ListViewStructSize, vWritten
                            WriteObjAttributes = (vWritten = udtFile.ListViewStructSize)
                        End If
                     End If
                End If
            End If
        Else
            WriteObjAttributes = True
        End If
    End If
EH:
    If Err Then Err.Clear
End Function

Private Function ReadFileHeader(fileNum As Long, udtFile As StructureFileHeader) As Boolean
    ' Read the file header only
    On Error GoTo EH
    Dim vRead As Long
    SetFilePointer fileNum, 0&, 0&, 0&
    ReadFile fileNum, udtFile, 40&, vRead
    If vRead = 40& Then ReadFileHeader = (udtFile.MagicNumber = 1447843404)
EH:
    If Err Then Err.Clear
End Function

Private Function ReadObjAttributes(fileNum As Long, udtFile As StructureFileHeader) As Boolean
    
    ' Section is read in the following order
    ' :: File Header, read in ReadFileHeader routine
    ' :: Font
    ' :: Picture Property
    ' :: Tag property
    ' :: Control Properties
    
    Dim udtAttr As StructureListView
    Dim udtFont As StructureFont
    Dim vRead As Long
    Dim lvFont As StdFont, lvPic As StdPicture
    
    On Error GoTo EH
    ReadFile fileNum, udtFont, 16&, vRead ' read the font information
    If vRead = 16& Then
        Set lvFont = New StdFont
        With udtFont ' create the font
            lvFont.Bold = (.Attributes And &H1)
            lvFont.Italic = ((.Attributes \ &H2) And &H1)
            lvFont.Strikethrough = ((.Attributes \ &H4) And &H1)
            lvFont.Underline = (.Attributes \ &H8)
            lvFont.CharSet = .CharSet
            lvFont.Size = .Size
            lvFont.Weight = .Weight
        End With
        ' read the font name & set the font
        If ReadString(fileNum, False, udtFont.Name) Then
            lvFont.Name = udtFont.Name
            ' read the picture property data
            If ReadPicture(fileNum, lvPic) Then
                ' read the .Tag property
                If ReadString(fileNum, (udtFile.Version = 1!), udtAttr.Tag) Then
                    ReadFile fileNum, udtAttr, 32&, vRead ' read the v1 structure size
                    If vRead = 32& Then
                        On Error Resume Next
                        With lvObject ' update the listview properties with what was read
                            .AllowColumnReorder = udtAttr.AllowColumnReorder
                            If udtAttr.AppearanceFlat Then .Appearance = ccFlat
                            .Arrange = udtAttr.Arrange
                            .BackColor = udtAttr.BackColor
                            If udtAttr.BorderStyleNone Then .BorderStyle = ccNone
                            .Checkboxes = udtAttr.Checkboxes
                            .FlatScrollBar = udtAttr.FlatScrollBar
                            .ForeColor = udtAttr.ForeColor
                            .FullRowSelect = udtAttr.FullRowSelect
                            .GridLines = udtAttr.GridLines
                            .HideColumnHeaders = udtAttr.HideColumnHeaders
                            .HotTracking = udtAttr.HotTracking
                            .HoverSelection = udtAttr.HoverSelection
                            If udtAttr.LabelEditManual Then .LabelEdit = lvwManual
                            .LabelWrap = udtAttr.LabelWrap
                            .MultiSelect = udtAttr.MultiSelect
                            .PictureAlignment = udtAttr.PictureAlignment
                            .SortKey = udtAttr.SortKey
                            If udtAttr.SortOrderDesc Then .SortOrder = lvwDescending
                            If udtAttr.Sorted Then udtFile.Flags = udtFile.Flags Or 1024& ' re-sort after list items loaded
                            lvObject.Sorted = False
                            If udtAttr.TextBkgOpaque Then .TextBackground = lvwOpaque
                            .View = udtAttr.View
                            .Tag = udtAttr.Tag
                            Set lvObject.Font = lvFont
                            Set lvObject.Picture = lvPic
                            ReadObjAttributes = True
                        End With
                    End If
                End If
            End If
        End If
    End If
EH:
    If Err Then Err.Clear
End Function

Private Function WriteHeaders(fileNum As Long, StringsAsAnsi As Boolean) As Boolean

    ' Section is written in the following order
    ' :: Properties
    ' :: Icon reference
    ' :: Tag
    
    Dim h As Long, vWritten As Long
    Dim udtHdr As StructureHeaders
    
    On Error GoTo EH
    For h = 1 To lvObject.ColumnHeaders.Count
        With lvObject.ColumnHeaders(h)
            udtHdr.Alignment = .Alignment
            udtHdr.Position = .Position
            udtHdr.Left = .Left
            udtHdr.Width = .Width
            Select Case VarType(.Icon)
            Case vbInteger
                udtHdr.IconType = vbInteger
            Case vbString
                If .Icon = vbNullString Then udtHdr.IconType = 0 Else udtHdr.IconType = 0
            Case Else
                udtHdr.IconType = 0
            End Select
            If Me.IncludeHeaderTags Then
                Select Case VarType(.Tag)
                Case vbString, vbLong, vbInteger, vbByte, vbDate, vbCurrency, vbSingle, vbDouble, vbBoolean
                    udtHdr.TagType = VarType(.Tag)
                Case Else
                    udtHdr.TagType = 0
                End Select
            End If
        End With
        WriteFile fileNum, udtHdr, 18&, vWritten
        If vWritten <> 18& Then Exit For
        Select Case udtHdr.IconType
        Case 0
        Case vbString
            If WriteString(fileNum, CStr(lvObject.ColumnHeaders(h).Icon), StringsAsAnsi) = False Then Exit For
        Case vbInteger
            WriteFile fileNum, CInt(lvObject.ColumnHeaders(h).Icon), 2, vWritten
            If vWritten <> 2 Then Exit For
        End Select
        If udtHdr.TagType Then
            If WriteVariantTag(fileNum, lvObject.ColumnHeaders(h).Tag, StringsAsAnsi) = False Then Exit For
        End If
        If WriteString(fileNum, lvObject.ColumnHeaders(h).Text, StringsAsAnsi) = False Then Exit For
        If WriteString(fileNum, lvObject.ColumnHeaders(h).Key, StringsAsAnsi) = False Then Exit For
    Next
    WriteHeaders = (h > lvObject.ColumnHeaders.Count)
EH:
    If Err Then Err.Clear
End Function

Private Function ReadHeaders(fileNum As Long, udtFile As StructureFileHeader) As Boolean

    ' Section is read in the following order
    ' :: Properties
    ' :: Icon reference
    ' :: Tag
    
    Dim h As Long, vRead As Long, tSize As Long
    Dim sIcon As String, iIcon As Integer
    Dim bApplyTags As Boolean, hasTags As Boolean
    Dim bApplyIcon As Boolean, StringsAsAnsi As Boolean
    
    Dim udtHdr() As StructureHeaders
    
    On Error Resume Next
    ReDim udtHdr(1 To udtFile.Columns)
    
    If (udtFile.Flags And fgHasHeaderTags) Then
        hasTags = True
        If Me.IncludeHeaderTags = True Then bApplyTags = True
    End If
    If (udtFile.Version = 1!) Then StringsAsAnsi = True
    If Not lvObject.ColumnHeaderIcons Is Nothing Then bApplyIcon = True
    
    For h = 1 To udtFile.Columns
        ReadFile fileNum, udtHdr(h), 18&, vRead
        If vRead <> 18 Then Exit For
        Select Case udtHdr(h).IconType
        Case vbString
            If ReadString(fileNum, StringsAsAnsi, sIcon) = False Then Exit For
            udtHdr(h).Icon = sIcon
        Case vbInteger
            ReadFile fileNum, iIcon, 2&, vRead
            If vRead <> 2 Then Exit For
            udtHdr(h).Icon = iIcon
        Case Else
            udtHdr(h).Icon = 0
        End Select
        If hasTags Then
            If ReadVariantTag(fileNum, udtHdr(h).TagType, StringsAsAnsi, udtHdr(h).Tag) = False Then Exit For
        End If
        If ReadString(fileNum, StringsAsAnsi, udtHdr(h).Text) = False Then Exit For
        If ReadString(fileNum, StringsAsAnsi, udtHdr(h).Key) = False Then Exit For
    Next
    If h > udtFile.Columns Then
        lvObject.ColumnHeaders.Clear
        For h = 1 To udtFile.Columns
            lvObject.ColumnHeaders.Add , udtHdr(h).Key, udtHdr(h).Text, udtHdr(h).Width, udtHdr(h).Alignment
            If (Not udtHdr(h).Icon = 0) And bApplyIcon = True Then lvObject.ColumnHeaders(h).Icon = udtHdr(h).Icon
            If bApplyTags Then lvObject.ColumnHeaders(h).Tag = udtHdr(h).Tag
        Next
        For h = 1 To udtFile.Columns
            If udtHdr(h).Position <> h Then lvObject.ColumnHeaders(h).Position = udtHdr(h).Position
        Next
        ReadHeaders = True
    End If
    
End Function

Private Function ReadImageLists(fileNum As Long, udtFile As StructureFileHeader) As Boolean
    
    ' Section is read in the following order
    ' :: Properties
    ' :: Image Count
    ' :: Image Data per image
    ' :: Tag per image
    ' :: Key per image
    
    Dim vTag As Variant, vRead As Long
    Dim lValue As Long, sKey As String
    Dim i As Long, L As Long, ansiStrings As Boolean
    
    Dim udtIL As StructureImageList
    Dim theList As MSComctlLib.ImageList
    Dim tPic As StdPicture, cImage As MSComctlLib.ListImage
    
    Set lvObject.ColumnHeaderIcons = Nothing
    Set lvObject.Icons = Nothing
    Set lvObject.SmallIcons = Nothing
    
    If udtFile.Version = 1! Then ansiStrings = True
    
    For L = 1 To 3
        On Error GoTo EH
        ReadFile fileNum, lValue, 4&, vRead
        If vRead <> 4& Then Exit For
        
        If lValue > 17& Then
            On Error Resume Next
            RaiseEvent SetImageList(L, theList)
            If theList Is Nothing Then
                ReadImageLists = True   ' user is not supplying an imagelist, abort
                Exit Function           ' but continue importing
            Else
                theList.ListImages.Clear
                If Err Then     ' we can't clear an imagelist if it is assigned to another control
                    Err.Clear   ' therefore, we will allow importing to continue, but without imagelists
                    ReadImageLists = True
                    Exit Function
                End If
            End If
            On Error GoTo EH
        ElseIf lValue Then
            If L = ilLargeIcons Then  ' listimage is same as ColumnHeaders
                Set lvObject.Icons = lvObject.ColumnHeaderIcons
            ElseIf lValue = ilColumnHeaders Then ' listimage is same as ColumnHeaders
                Set lvObject.SmallIcons = lvObject.ColumnHeaderIcons
            Else ' listimage is same as large icons
                Set lvObject.SmallIcons = lvObject.Icons
            End If
            lValue = 0& ' we won't be importing any images
        End If
        
        If lValue Then
            lValue = 18&
            ReadFile fileNum, udtIL, lValue, vRead
            If vRead <> lValue Then Exit For
            With theList
                .BackColor = udtIL.BackColor
                .ImageHeight = udtIL.ImageHeight
                .ImageWidth = udtIL.ImageWidth
                .MaskColor = udtIL.MaskColor
                .UseMaskColor = udtIL.UseMaskColor
            End With
            If ReadString(fileNum, ansiStrings, sKey) Then theList.Tag = sKey
            
            For i = 1 To udtIL.Count
                If ReadPicture(fileNum, tPic) = False Then Exit For
                ReadFile fileNum, lValue, 4&, vRead
                If vRead <> 4& Then Exit For
                If ReadVariantTag(fileNum, lValue, ansiStrings, vTag) = False Then Exit For
                If ReadString(fileNum, ansiStrings, sKey) = False Then Exit For
                Set cImage = theList.ListImages.Add(, sKey, tPic)
                cImage.Tag = vTag
            Next
            If i <= udtIL.Count Then Exit For
            Select Case L
            Case 1: Set lvObject.ColumnHeaderIcons = theList
            Case 2: Set lvObject.Icons = theList
            Case 3: Set lvObject.SmallIcons = theList
            End Select
        End If
    Next
    ReadImageLists = (L = 4)
EH:
    If Err Then Err.Clear
End Function

Private Function WriteImageLists(fileNum As Long, ansiStrings As Boolean) As Boolean

    ' Section is written in the following order
    ' :: Properties
    ' :: Image Count
    ' :: Image Data per image
    ' :: Tag per image
    ' :: Key per image
    
    Dim i As Long, vWritten As Long, L As Long
    Dim theList As MSComctlLib.ImageList
    Dim tPic As StdPicture, TagType As VbVarType
    Dim udtIL As StructureImageList
    
    On Error GoTo EH
    For L = 1 To 3
        Select Case L
        Case 1: Set theList = lvObject.ColumnHeaderIcons
            If theList Is Nothing Then
                WriteFile fileNum, 0&, 4&, vWritten
                If vWritten <> 4& Then Exit For
            End If
        Case 2: Set theList = lvObject.Icons
            If theList Is Nothing Then
                WriteFile fileNum, 0&, 4&, vWritten
                If vWritten <> 4& Then Exit For
            Else
                If theList Is lvObject.ColumnHeaderIcons Then
                    Set theList = Nothing ' uses same list as ColumnHeaderIcons
                    WriteFile fileNum, 1&, 4&, vWritten
                    If vWritten <> 4& Then Exit For
                End If
            End If
        Case 3: Set theList = lvObject.SmallIcons
            If theList Is Nothing Then
                WriteFile fileNum, 0&, 4&, vWritten
                If vWritten <> 4& Then Exit For
            ElseIf theList Is lvObject.ColumnHeaderIcons Then
                Set theList = Nothing ' uses same list as ColumnHeaderIcons
                WriteFile fileNum, 1&, 4&, vWritten
                If vWritten <> 4& Then Exit For
            ElseIf theList Is lvObject.Icons Then
                Set theList = Nothing ' uses same list as Icons
                WriteFile fileNum, 2&, 4&, vWritten
                If vWritten <> 4& Then Exit For
            End If
        End Select
    
        If Not theList Is Nothing Then
            With udtIL
                .BackColor = theList.BackColor
                .Count = theList.ListImages.Count
                .ImageHeight = theList.ImageHeight
                .ImageWidth = theList.ImageWidth
                .MaskColor = theList.MaskColor
                .UseMaskColor = theList.UseMaskColor
            End With
            WriteFile fileNum, 18&, 4&, vWritten
            If vWritten <> 4& Then Exit For
            WriteFile fileNum, udtIL, 18&, vWritten
            If vWritten <> 18& Then Exit For
            If WriteString(fileNum, theList.Tag, ansiStrings) = False Then Exit For
            For i = 1 To udtIL.Count
                
                Set tPic = theList.ListImages(i).Picture
                If WritePicture(fileNum, tPic) = False Then Exit For
                
                WriteFile fileNum, CLng(VarType(theList.ListImages(i).Tag)), 4&, vWritten
                If vWritten <> 4& Then Exit For
                If WriteVariantTag(fileNum, theList.ListImages(i).Tag, ansiStrings) = False Then Exit For
                If WriteString(fileNum, theList.ListImages(i).Key, ansiStrings) = False Then Exit For
            Next
            If i <= theList.ListImages.Count Then Exit For
        End If
    Next
    WriteImageLists = (L = 4)
EH:
    If Err Then Err.Clear
End Function

Private Function WriteListItems(fileNum As Long, ansiStrings As Boolean) As Boolean

    ' Section is written in the following order
    ' :: Properties
    ' :: ListSubItems Count
    ' :: Icon reference type & reference
    ' :: Text
    ' :: Key
    ' :: ToolTipText
    ' -- for each sub item
    ' :: Properties
    ' :: Text
    ' :: Key
    ' :: ToolTipText
    ' :: ReportIcon reference type & reference (ListSubItems property)

    Dim udtLI As StructureListItem
    Dim udtLIsub As StructureListSubItem
    Dim vWritten As Long
    Dim L As Long, S As Long
    
    For L = 1 To lvObject.ListItems.Count
        With lvObject.ListItems(L)
            If Me.IncludeTextFormatting Then
                udtLI.BoldChecked = Abs(.Bold) Or Abs(.Checked) * 2&
                udtLI.ForeColor = .ForeColor
            Else
                udtLI.BoldChecked = Abs(Checked) * 2&
            End If
            Select Case VarType(.Icon)
            Case vbString
                If .Icon = vbNullString Then udtLI.IconTypes = 0 Else udtLI.IconTypes = vbString
            Case vbInteger
                udtLI.IconTypes = vbInteger
            Case Else
                udtLI.IconTypes = 0&
            End Select
            Select Case VarType(.SmallIcon)
            Case vbString
                If Not .SmallIcon = vbNullString Then
                    udtLI.IconTypes = udtLI.IconTypes Or (vbString * &H10000)
                End If
            Case vbInteger
                udtLI.IconTypes = udtLI.IconTypes Or (vbInteger * &H10000)
            End Select
            If Me.IncludeListItemTags Then
                Select Case VarType(.Tag)
                Case vbString, vbLong, vbInteger, vbByte, vbDate, vbCurrency, vbSingle, vbDouble, vbBoolean
                    udtLI.TagType = VarType(.Tag)
                Case Else
                    udtLI.TagType = 0
                End Select
            End If
            udtLI.SubCount = .ListSubItems.Count
            WriteFile fileNum, udtLI, 16&, vWritten
            If vWritten <> 16& Then Exit For
            If WriteString(fileNum, .Text, ansiStrings) = False Then Exit For
            If WriteString(fileNum, .Key, ansiStrings) = False Then Exit For
            If WriteString(fileNum, .ToolTipText, ansiStrings) = False Then Exit For
            Select Case (udtLI.IconTypes And &HFFFF&)
            Case 0
            Case vbString
                If WriteString(fileNum, CStr(.Icon), ansiStrings) = False Then Exit For
            Case vbInteger
                WriteFile fileNum, CInt(.Icon), 2, vWritten
                If vWritten <> 2 Then Exit For
            End Select
            Select Case (udtLI.IconTypes \ &H10000)
            Case 0
            Case vbString
                If WriteString(fileNum, CStr(.SmallIcon), ansiStrings) = False Then Exit For
            Case vbInteger
                WriteFile fileNum, CInt(.SmallIcon), 2, vWritten
                If vWritten <> 2 Then Exit For
            End Select
            If Me.IncludeListItemTags Then
                If WriteVariantTag(fileNum, .Tag, ansiStrings) = False Then Exit For
            End If
        End With
        For S = 1 To udtLI.SubCount
            With lvObject.ListItems(L).ListSubItems(S)
                udtLIsub.Bold = Abs(.Bold)
                udtLIsub.ForeColor = .ForeColor
                Select Case VarType(.ReportIcon)
                Case vbString, vbInteger
                    udtLIsub.IconType = VarType(.ReportIcon)
                Case Else
                    udtLIsub.IconType = 0&
                End Select
                If Me.IncludeListItemTags Then
                    Select Case VarType(.Tag)
                    Case vbString, vbLong, vbInteger, vbByte, vbDate, vbCurrency, vbSingle, vbDouble, vbBoolean
                        udtLIsub.TagType = VarType(.Tag)
                    Case Else
                        udtLIsub.TagType = 0
                    End Select
                End If
                WriteFile fileNum, udtLIsub, 9&, vWritten
                If vWritten <> 9& Then Exit For
                If WriteString(fileNum, .Text, ansiStrings) = False Then Exit For
                If WriteString(fileNum, .Key, ansiStrings) = False Then Exit For
                If WriteString(fileNum, .ToolTipText, ansiStrings) = False Then Exit For
                Select Case udtLIsub.IconType
                Case 0
                Case vbString
                    If WriteString(fileNum, CStr(.ReportIcon), ansiStrings) = False Then Exit For
                Case vbInteger
                    WriteFile fileNum, CInt(.ReportIcon), 2, vWritten
                    If vWritten <> 2 Then Exit For
                End Select
                If Me.IncludeListItemTags Then
                    If WriteVariantTag(fileNum, .Tag, ansiStrings) = False Then Exit For
                End If
            End With
        Next
        If S <= udtLI.SubCount Then Exit For
    Next
    WriteListItems = (L > lvObject.ListItems.Count)

End Function

Private Function ReadListItems(fileNum As Long, udtFile As StructureFileHeader) As Boolean

    ' Section is read in the following order
    ' :: Properties
    ' :: ListSubItems Count
    ' :: Icon reference type & reference
    ' :: Text
    ' :: Key
    ' :: ToolTipText
    ' -- for each sub item
    ' :: Properties
    ' :: Text
    ' :: Key
    ' :: ToolTipText
    ' :: ReportIcon reference type & reference (ListSubItems property)
    
    Dim udtLI As StructureListItem
    Dim udtLIsub As StructureListSubItem
    Dim vRead As Long, sValue As String
    Dim L As Long, S As Long, iValue As Integer
    Dim bApplyTag As Boolean, hasTags As Boolean
    Dim bApplyIconLG As Boolean, bApplyIconSM As Boolean
    Dim bApplyFormat As Boolean, ansiStrings As Boolean
    Dim xItm As MSComctlLib.ListItem
    Dim xSub As MSComctlLib.ListSubItem
    
    On Error Resume Next
    ' Why resume next?  Icon references, when they exist...
    ' When icon references exist, here are some errors that are expected
    ' 1. You did not provide an imagelist control when prompted: error will be no imagelist is initialized
    ' 2. You did not provide an imagelist but set one to the listview before this is called: error is index/key does not exist
    ' 3. The imagelist was saved but you opted not to import it: error can be one of the two above
    
    If (udtFile.Flags And fgHasItemTags) Then
        hasTags = True
        If Me.IncludeListItemTags = True Then bApplyTag = True
    End If
    If (udtFile.Flags And fgHasTextFormat) Then
        If Me.IncludeTextFormatting Then bApplyFormat = True
    End If
    If udtFile.Version = 1! Then ansiStrings = True
    If Not lvObject.Icons Is Nothing Then bApplyIconLG = True
    If Not lvObject.SmallIcons Is Nothing Then bApplyIconSM = True
    
    For L = 1 To udtFile.Items
        ReadFile fileNum, udtLI, 16&, vRead
        If vRead <> 16& Then Exit For
        If ReadString(fileNum, ansiStrings, udtLI.Text) = False Then Exit For
        If ReadString(fileNum, ansiStrings, udtLI.Key) = False Then Exit For
        If ReadString(fileNum, ansiStrings, udtLI.Tooltip) = False Then Exit For
        Select Case (udtLI.IconTypes And &HFFFF&)
        Case 0
        Case vbString
            If ReadString(fileNum, ansiStrings, sValue) = False Then Exit For
            udtLI.IconLg = sValue
        Case vbInteger
            ReadFile fileNum, iValue, 2&, vRead
            If vRead <> 2& Then Exit For
            udtLI.IconLg = iValue
        End Select
        Select Case (udtLI.IconTypes \ &H10000)
        Case 0
        Case vbString
            If ReadString(fileNum, ansiStrings, sValue) = False Then Exit For
            udtLI.IconSm = sValue
        Case vbInteger
            ReadFile fileNum, iValue, 2&, vRead
            If vRead <> 2& Then Exit For
            udtLI.IconSm = iValue
        End Select
        If hasTags Then
            If ReadVariantTag(fileNum, udtLI.TagType, ansiStrings, udtLI.Tag) = False Then Exit For
        End If
        Set xItm = lvObject.ListItems.Add(, udtLI.Key, udtLI.Text)
        If (udtLI.IconTypes And &HFFFF&) Then
            If bApplyIconLG Then xItm.Icon = udtLI.IconLg
        End If
        If (udtLI.IconTypes \ &H10000) Then
            If bApplyIconSM Then xItm.SmallIcon = udtLI.IconSm
        End If
        xItm.Checked = CBool(udtLI.BoldChecked \ 2)
        xItm.ToolTipText = udtLI.Tooltip
        If bApplyFormat Then
            xItm.Bold = CBool(udtLI.BoldChecked And 1)
            xItm.ForeColor = udtLI.ForeColor
        End If
        If bApplyTag Then xItm.Tag = udtLI.Tag
        
        For S = 1 To udtLI.SubCount
            ReadFile fileNum, udtLIsub, 9&, vRead
            If vRead <> 9& Then Exit For
            If ReadString(fileNum, ansiStrings, udtLIsub.Text) = False Then Exit For
            If ReadString(fileNum, ansiStrings, udtLIsub.Key) = False Then Exit For
            If ReadString(fileNum, ansiStrings, udtLIsub.Tooltip) = False Then Exit For
            Select Case udtLIsub.IconType
            Case 0
            Case vbString
                If ReadString(fileNum, ansiStrings, sValue) = False Then Exit For
                udtLIsub.ReportIcon = sValue
            Case vbInteger
                ReadFile fileNum, iValue, 2&, vRead
                If vRead <> 2& Then Exit For
                udtLIsub.ReportIcon = iValue
            End Select
            If hasTags Then
                If ReadVariantTag(fileNum, udtLIsub.TagType, ansiStrings, udtLIsub.Tag) = False Then Exit For
            End If
            Set xSub = xItm.ListSubItems.Add(, udtLIsub.Key, udtLIsub.Text, , udtLIsub.Tooltip)
            If (Not udtLIsub.IconType = 0) Then xSub.ReportIcon = udtLIsub.ReportIcon
            If bApplyTag Then xSub.Tag = udtLIsub.Tag
            If bApplyFormat Then
                xSub.Bold = CBool(udtLIsub.Bold)
                xSub.ForeColor = udtLIsub.ForeColor
            End If
        Next
        If S <= udtLI.SubCount Then Exit For
    Next
    ReadListItems = (L > udtFile.Items)
    
End Function

Private Function ReadVariantTag(fileNum As Long, ByVal varTypo As VbVarType, StringsAsAnsi As Boolean, varReturn As Variant) As Boolean

    ' Converts file bytes to proper variable type
    ' Many ListView class Tags are Variant vs String

    Dim sTag As String, byTag As Byte, iTag As Integer
    Dim blTag As Boolean, lTag As Long, snTag As Single
    Dim dtTag As Date, dbTag As Double, crTag As Currency
    Dim tSize As Long, vRead As Long
        
    Select Case varTypo
    Case vbString
        If ReadString(fileNum, StringsAsAnsi, sTag) Then
            varReturn = sTag
            vRead = tSize
        Else
            vRead = tSize - 1&
        End If
    Case vbByte
        tSize = 1&
        ReadFile fileNum, byTag, tSize, vRead
        varReturn = byTag
    Case vbInteger
        tSize = 2&
        ReadFile fileNum, iTag, tSize, vRead
        varReturn = iTag
    Case vbBoolean
        tSize = 2&
        ReadFile fileNum, blTag, tSize, vRead
        varReturn = blTag
    Case vbLong
        tSize = 4&
        ReadFile fileNum, lTag, tSize, vRead
        varReturn = lTag
    Case vbSingle
        tSize = 4&
        ReadFile fileNum, snTag, tSize, vRead
        varReturn = snTag
    Case vbDate
        tSize = 8&
        ReadFile fileNum, dtTag, tSize, vRead
        varReturn = dtTag
    Case vbDouble
        tSize = 8&
        ReadFile fileNum, dbTag, tSize, vRead
        varReturn = dbTag
    Case vbCurrency
        tSize = 8&
        ReadFile fileNum, crTag, tSize, vRead
        varReturn = crTag
    Case Else
        vRead = tSize
        varReturn = vbNullString
    End Select
    If vRead = tSize Then
        varReturn = varReturn
        ReadVariantTag = True
    End If
    
End Function

Private Function WriteVariantTag(fileNum As Long, vTag As Variant, ansiStrings As Boolean) As Boolean
 
    ' Converts variable type to bytes for file writing
    ' Many ListView class Tags are Variant vs String

    Dim vWritten As Long
    Select Case VarType(vTag)
    Case vbString
        WriteVariantTag = WriteString(fileNum, CStr(vTag), ansiStrings)
    Case vbByte
        WriteFile fileNum, CByte(vTag), 1&, vWritten
        WriteVariantTag = (vWritten = 1&)
    Case vbInteger
        WriteFile fileNum, CInt(vTag), 2&, vWritten
        WriteVariantTag = (vWritten = 2&)
    Case vbBoolean
        WriteFile fileNum, CBool(vTag), 2&, vWritten
        WriteVariantTag = (vWritten = 2&)
    Case vbLong
        WriteFile fileNum, CLng(vTag), 4&, vWritten
        WriteVariantTag = (vWritten = 4&)
    Case vbSingle
        WriteFile fileNum, CSng(vTag), 4&, vWritten
        WriteVariantTag = (vWritten = 4&)
    Case vbDouble
        WriteFile fileNum, CDbl(vTag), 8&, vWritten
        WriteVariantTag = (vWritten = 8&)
    Case vbDate
        WriteFile fileNum, CDate(vTag), 8&, vWritten
        WriteVariantTag = (vWritten = 8&)
    Case vbCurrency
        WriteFile fileNum, CCur(vTag), 8&, vWritten
        WriteVariantTag = (vWritten = 8&)
    Case Else
        WriteVariantTag = True
    End Select

End Function

Private Function GetFileHandle(FileName As String, bOpen As Boolean, useUnicode As Boolean) As Long

    ' Function uses APIs to read/create files with unicode support

    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const GENERIC_WRITE As Long = &H40000000
    Const FILE_SHARE_WRITE As Long = &H2
    Const CREATE_ALWAYS As Long = 2
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    
    Dim Flags As Long, Access As Long, bUnicode As Boolean
    Dim Disposition As Long, Share As Long
    
    If bOpen Then
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    Else
        Access = GENERIC_READ Or GENERIC_WRITE
        Share = 0&
        If useUnicode Then
            Flags = GetFileAttributesW(StrPtr(FileName))
        Else
            Flags = GetFileAttributes(FileName)
        End If
        If Flags < 0& Then Flags = FILE_ATTRIBUTE_NORMAL
        ' CREATE_ALWAYS will delete previous file if necessary
        Disposition = CREATE_ALWAYS
    End If
    
    If useUnicode Then
        GetFileHandle = CreateFileW(StrPtr(FileName), Access, Share, ByVal 0&, Disposition, Flags, 0&)
    Else
        GetFileHandle = CreateFile(FileName, Access, Share, ByVal 0&, Disposition, Flags, 0&)
    End If

End Function

Private Function FileExists(FileName As String, useUnicode As Boolean) As Boolean
    ' test to see if a file exists
    Const INVALID_HANDLE_VALUE = -1&
    If useUnicode Then
        FileExists = Not (GetFileAttributesW(StrPtr(FileName)) = INVALID_HANDLE_VALUE)
    Else
        FileExists = Not (GetFileAttributes(FileName) = INVALID_HANDLE_VALUE)
    End If
End Function


Private Function WriteString(fileNum As Long, theString As String, asAnsi As Boolean) As Boolean
    
    ' Writes a string to file in either ANSI or Unicode
    Dim sData() As Byte, lLen As Long, vWritten As Long
    
    lLen = Len(theString)
    WriteFile fileNum, lLen, 4&, vWritten
    If lLen Then
        If asAnsi Then
            sData() = StrConv(theString, vbFromUnicode)
            WriteFile fileNum, sData(0), lLen, vWritten
        Else
            lLen = lLen + lLen
            WriteFile fileNum, ByVal StrPtr(theString), lLen, vWritten
        End If
        WriteString = (vWritten = lLen)
    Else
        WriteString = True
    End If

End Function

Private Function ReadString(fileNum As Long, asAnsi As Boolean, toString As String) As Boolean

    ' Reads either ANSI or Unicode string from file
    
    Dim lLen As Long, vRead As Long, sData() As Byte

    ReadFile fileNum, lLen, 4&, vRead
    If vRead = 4& Then
        If lLen Then
            If asAnsi Then  ' ANSI version
                ReDim sData(0 To lLen - 1)
                ReadFile fileNum, sData(0), lLen, vRead
                If vRead = lLen Then
                    toString = StrConv(sData(), vbUnicode)
                    ReadString = True
                End If
            Else
                toString = String$(lLen, vbNullChar)
                lLen = lLen + lLen
                ReadFile fileNum, ByVal StrPtr(toString), lLen, vRead
                If vRead = lLen Then ReadString = True
            End If
        Else
            toString = vbNullString
            ReadString = True
        End If
    End If

End Function

Private Function WritePicture(fileNum As Long, thePicture As StdPicture) As Boolean

    ' Writes a picture property to file, preceding it with byte length
    
    Dim o_hMem As Long, o_lpMem As Long
    Dim o_lngByteCount As Long
    Dim vWritten As Long, bSuccess As Integer
    
    ' In order to use the OLESaveToStream function, we need the pointer to the IPersistStream interface
    ' This portion will return that pointer if the interface is implemented
    
    Dim varRtn As Variant, pIPS As Long
    Dim aGUID(0 To 3) As Long, IFauxStream As IUnknown
    Dim pvTypes(0 To 1) As Integer, pvPtrs(0 To 1) As Long, pValues(0 To 1) As Variant
    
    Const IUnknownQueryInterface As Long = 0&   ' IUnknown vTable offset to Query implemented interfaces
    Const IUnknownRelease As Long = 8&          ' IUnkownn vTable offset to decrement reference count
    Const CC_STDCALL As Long = 4&
    
    aGUID(0) = &H109            ' GUID for IPersistStream (00000109-0000-0000-C000-000000000046)
    aGUID(2) = &HC0
    aGUID(3) = &H46000000
    
    If Not thePicture Is Nothing Then
        If thePicture.Handle Then
    
            On Error GoTo CATCH_EXCEPTION
            pvTypes(0) = vbLong: pValues(0) = VarPtr(aGUID(0)): pvPtrs(0) = VarPtr(pValues(0))
            pvTypes(1) = vbLong: pValues(1) = VarPtr(pIPS): pvPtrs(1) = VarPtr(pValues(1))
            ' See if the stdPicture implements IUnknownQueryInterface & it should
            Call DispCallFunc(ObjPtr(thePicture), IUnknownQueryInterface, CC_STDCALL, vbLong, 2&, VarPtr(pvTypes(0)), VarPtr(pvPtrs(0)), varRtn)
            If pIPS Then ' we have IPersistStream interface
                ' IUnknownQueryInterface ups the ref count, release the reference
                Call DispCallFunc(pIPS, IUnknownRelease, CC_STDCALL, vbLong, 0&, 0&, 0&, varRtn)
                
                ' now create a stream
                Call CreateStreamOnHGlobal(0, 1, IFauxStream)
                ' call API to write the stdPic to the fake IStream
                If OleSaveToStream(pIPS, IFauxStream) = 0& Then
                    On Error Resume Next
                    If GetHGlobalFromStream(ByVal ObjPtr(IFauxStream), o_hMem) = 0 Then
                        o_lngByteCount = GlobalSize(o_hMem)
                        If o_lngByteCount > 0 Then
                            o_lpMem = GlobalLock(o_hMem)
                            If o_lpMem <> 0 Then
                                WriteFile fileNum, o_lngByteCount, 4&, vWritten
                                If vWritten = 4& Then
                                    WriteFile fileNum, ByVal o_lpMem, o_lngByteCount, vWritten
                                    If vWritten = o_lngByteCount Then
                                        bSuccess = 1
                                    Else
                                        bSuccess = -1
                                    End If
                                Else
                                    bSuccess = -1
                                End If
                                GlobalUnlock o_hMem
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

CATCH_EXCEPTION:
    If Err Then Err.Clear
    If bSuccess = 1 Then
        WritePicture = (vWritten = o_lngByteCount)
    ElseIf bSuccess = 0 Then
        WriteFile fileNum, 0&, 4&, vWritten
        WritePicture = (vWritten = 4&)
    End If
    
End Function

Private Function ReadPicture(fileNum As Long, thePicture As StdPicture) As Boolean
    
    ' Reads binary data from file and converts it to a stdPicture
    
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim vRead As Long, o_ByteSize As Long, bSuccess As Integer
    Dim IIStream As IUnknown, IPic As IPicture
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    On Error Resume Next
    ReadFile fileNum, o_ByteSize, 4&, vRead
    If vRead = 4& Then
        If o_ByteSize Then
            o_ByteSize = o_ByteSize - 16&
            o_hMem = GlobalAlloc(&H2&, o_ByteSize)
            If Not o_hMem = 0& Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem Then
                    SetFilePointer fileNum, 16&, 0&, 1& ' 1st 16 bytes is a guid we don't need
                    ReadFile fileNum, ByVal o_lpMem, o_ByteSize, vRead
                    Call GlobalUnlock(o_hMem)
                    If vRead = o_ByteSize Then
                        If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                              Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), IPic)
                        End If
                        bSuccess = 1
                    Else
                        bSuccess = -1
                    End If
                End If
            End If
        End If
    Else
        bSuccess = -1
    End If
    
    If bSuccess = 1 Then
        Set thePicture = IPic
        ReadPicture = True
    Else
        Set thePicture = Nothing
        ReadPicture = (bSuccess = 0)
    End If

End Function

Private Sub Class_Initialize()
    ' start with these properites defaulted to True
    c_Flags = fgHasHeaders Or fgHasImageList Or fgHasTextFormat
End Sub


