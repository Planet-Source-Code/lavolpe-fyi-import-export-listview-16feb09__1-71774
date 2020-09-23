VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Import Example"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Report View"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Icon View"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5318
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      ForeColor       =   8388608
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Col1"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Col2"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "col3"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' No export example provided; it is too easy
' Read the comments at the top of the class and any comments in the Import & Export functions
' The class was not supported for version 5 of the listview/imagelist control. Supports version 6 only


' must use WithEvents if importing, not required for exporting
Private WithEvents cListViewExportImport  As clsLstVwExportImport
Attribute cListViewExportImport.VB_VarHelpID = -1

Private Sub cListViewExportImport_ImportSummary(ByVal ColumnHeadersIncluded As Boolean, ByVal ImageListsNeeded As Long, ByVal ListItems As Boolean, Continue As Boolean)
    ' See comments at top of the class
    ' This event is called so you can decide what is to be imported and what is not
    ' You also have the opportunity to set class properties and listview control properties at this time
    ' Importing will continue only if the Continue parameter is set to True
    ListView1.ListItems.Clear
    
    cListViewExportImport.IncludeControlFormatting = True
    
    cListViewExportImport.IncludeListItemTags = False
    cListViewExportImport.IncludeHeaderTags = False
    
    Continue = True
    
End Sub

Private Sub cListViewExportImport_SetImageList(ByVal ListViewSection As lvImportImageList, ImportToImageList As MSComctlLib.ImageList)
    ' See comments at top of the class
    ' This event is only called if you are importing image lists
    ' If not importing imagelists, then your listview should already be assigned imagelist(s) as needed,
    '   otherwise, no icons will be displayed
    Set ImportToImageList = ImageList1
End Sub

Private Sub Command1_Click()
    
    ' You will notice the imagelist in design view is empty & set to custom size of 48x48
    ' When the images are imported, the images will be uploaded to the imagelist control as 32x32
    ' The control will be displayed in the same manner as what was exported
    Set cListViewExportImport = New clsLstVwExportImport
    cListViewExportImport.ImportFromFile ListView1, App.Path & "\IconViewExample.lve"
    
End Sub

Private Sub Command2_Click()
        
    ' You will notice the imagelist in design view is empty & set to custom size of 48x48
    ' When the images are imported, the images will be uploaded to the imagelist control of 16x16
    ' The control will be displayed in the same manner as what was exported
    Set cListViewExportImport = New clsLstVwExportImport
    cListViewExportImport.ImportFromFile ListView1, App.Path & "\ReportViewExample.lve"

End Sub
