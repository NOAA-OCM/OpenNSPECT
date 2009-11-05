VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPrj 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7170
   ClientLeft      =   2355
   ClientTop       =   1575
   ClientWidth     =   9630
   Icon            =   "frmPrj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Water Quality Standard "
      Height          =   870
      Left            =   6390
      TabIndex        =   58
      Top             =   2520
      Width           =   3135
      Begin VB.ComboBox cboWQStd 
         Height          =   315
         Left            =   735
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   330
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         Height          =   240
         Index           =   6
         Left            =   135
         TabIndex        =   60
         Top             =   375
         Width           =   660
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Watershed Delineation "
      Height          =   870
      Left            =   3255
      TabIndex        =   55
      Top             =   2520
      Width           =   3030
      Begin VB.ComboBox cboWSDelin 
         Height          =   315
         Left            =   690
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   330
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Index           =   3
         Left            =   105
         TabIndex        =   57
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Precipitation Scenario  "
      Height          =   870
      Left            =   150
      TabIndex        =   52
      Top             =   2520
      Width           =   3030
      Begin VB.ComboBox cboPrecipScen 
         Height          =   315
         Left            =   705
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   330
         Width           =   2115
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         Height          =   270
         Index           =   7
         Left            =   135
         TabIndex        =   54
         Top             =   375
         Width           =   735
      End
   End
   Begin VB.Frame frm_raintype 
      Caption         =   "Miscellaneous "
      Height          =   1695
      Left            =   6915
      TabIndex        =   50
      Top             =   735
      Width           =   2625
      Begin VB.ComboBox cboSelectPoly 
         Enabled         =   0   'False
         Height          =   315
         Left            =   675
         TabIndex        =   62
         Top             =   795
         Width           =   1680
      End
      Begin VB.CheckBox chkSelectedPolys 
         Caption         =   "Selected Polygons Only"
         Height          =   330
         Left            =   105
         TabIndex        =   61
         ToolTipText     =   "Select to limit analysis to selected polygons from a map layer"
         Top             =   330
         Width           =   2205
      End
      Begin VB.CheckBox chkLocalEffects 
         Caption         =   "Local Effects Only"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         ToolTipText     =   "Select for analysis of local effects only"
         Top             =   1290
         Width           =   1965
      End
      Begin VB.Label lblLayer 
         Caption         =   "Layer: "
         Enabled         =   0   'False
         Height          =   225
         Left            =   90
         TabIndex        =   63
         Top             =   840
         Width           =   660
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Project Information "
      Height          =   615
      Left            =   135
      TabIndex        =   37
      Top             =   45
      Width           =   9375
      Begin VB.CommandButton cmdOpenWS 
         Height          =   300
         Left            =   8550
         MaskColor       =   &H00000000&
         Picture         =   "frmPrj.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   210
         Width           =   345
      End
      Begin VB.TextBox txtOutputWS 
         Height          =   300
         Left            =   5595
         TabIndex        =   39
         Top             =   225
         Width           =   2880
      End
      Begin VB.TextBox txtProjectName 
         Height          =   300
         Left            =   945
         TabIndex        =   38
         Top             =   225
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Working Directory:"
         Height          =   270
         Left            =   4050
         TabIndex        =   42
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label4 
         Caption         =   "Name:"
         Height          =   240
         Left            =   165
         TabIndex        =   41
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Soils"
      Height          =   1695
      Left            =   3600
      TabIndex        =   9
      Top             =   735
      Width           =   3210
      Begin VB.ComboBox cboSoilsLayer 
         Height          =   315
         Left            =   105
         TabIndex        =   11
         Top             =   555
         Width           =   2115
      End
      Begin VB.Label Label6 
         Caption         =   "Hydrologic Soils Data Set:"
         Height          =   240
         Left            =   105
         TabIndex        =   47
         Top             =   945
         Width           =   2355
      End
      Begin VB.Label lblSoilsHyd 
         Height          =   315
         Left            =   105
         TabIndex        =   46
         Top             =   1185
         Width           =   2970
      End
      Begin VB.Label Label2 
         Caption         =   "Soils Definition:"
         Height          =   255
         Left            =   105
         TabIndex        =   10
         Top             =   285
         Width           =   2040
      End
   End
   Begin VB.Frame fraLC 
      Caption         =   "Land Cover"
      Height          =   1695
      Left            =   135
      TabIndex        =   5
      Top             =   735
      Width           =   3390
      Begin VB.ComboBox cboLCUnits 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmPrj.frx":0E00
         Left            =   1035
         List            =   "frmPrj.frx":0E0A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   765
         Width           =   2115
      End
      Begin VB.ComboBox cboLCLayer 
         Height          =   315
         Left            =   1035
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Choose a Land Cover GRID from current layers in map view"
         Top             =   345
         Width           =   2115
      End
      Begin VB.ComboBox cboLCType 
         Height          =   315
         ItemData        =   "frmPrj.frx":0E1C
         Left            =   1035
         List            =   "frmPrj.frx":0E1E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Grid Units:"
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   8
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Grid:"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   7
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Type:"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   1215
         Width           =   540
      End
   End
   Begin VB.ComboBox cboAreaLayer 
      Height          =   315
      Left            =   3465
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      Left            =   1650
      TabIndex        =   32
      Text            =   "cboClass"
      Top             =   6765
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Left            =   9435
      Top             =   4905
   End
   Begin VB.ComboBox cboCoeffSet 
      Height          =   315
      Left            =   1500
      TabIndex        =   31
      Text            =   "cboCoeffSet"
      Top             =   6750
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboCoeff 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmPrj.frx":0E20
      Left            =   1410
      List            =   "frmPrj.frx":0E30
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   6780
      Visible         =   0   'False
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3045
      Left            =   240
      TabIndex        =   12
      Top             =   3480
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   5371
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Pollutants"
      TabPicture(0)   =   "frmPrj.frx":0E54
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdCoeffs"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Erosion"
      TabPicture(1)   =   "frmPrj.frx":0E70
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblErodFactor"
      Tab(1).Control(1)=   "lblKFactor"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "chkCalcErosion"
      Tab(1).Control(5)=   "frameRainFall"
      Tab(1).Control(6)=   "cboErodFactor"
      Tab(1).Control(7)=   "cboSoilAttribute"
      Tab(1).Control(8)=   "frmSDR"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Land Uses"
      TabPicture(2)   =   "frmPrj.frx":0E8C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdLU"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Management Scenarios"
      TabPicture(3)   =   "frmPrj.frx":0EA8
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "grdLCChanges"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame frmSDR 
         Caption         =   "Frame7"
         Height          =   855
         Left            =   -70560
         TabIndex        =   64
         Top             =   1320
         Width           =   4455
         Begin VB.CommandButton cmdOpenSDR 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3840
            Picture         =   "frmPrj.frx":0EC4
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   360
            Width           =   345
         End
         Begin VB.TextBox txtSDRGRID 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   3495
         End
         Begin VB.CheckBox chkSDR 
            Caption         =   "Sediment Delivery Ratio GRID (optional)"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   0
            Width           =   3255
         End
      End
      Begin VB.ComboBox cboSoilAttribute 
         Height          =   315
         Left            =   -72480
         TabIndex        =   44
         Top             =   2640
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.ComboBox cboErodFactor 
         Height          =   315
         Left            =   -70440
         TabIndex        =   29
         Top             =   2640
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Frame frameRainFall 
         Caption         =   "Rainfall Factor "
         Height          =   1335
         Left            =   -74805
         TabIndex        =   22
         Top             =   1290
         Width           =   4020
         Begin VB.TextBox txtRainValue 
            Height          =   285
            Left            =   2085
            TabIndex        =   26
            ToolTipText     =   "Functionality to be implemented in Alpha2"
            Top             =   855
            Width           =   1260
         End
         Begin VB.ComboBox cboRainGrid 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1500
            TabIndex        =   25
            Text            =   "cboRainGRID"
            Top             =   375
            Width           =   1830
         End
         Begin VB.OptionButton optUseValue 
            Caption         =   "Use Constant Value: "
            Height          =   315
            Left            =   285
            TabIndex        =   24
            ToolTipText     =   "Functionality to be implemented in Alpha2"
            Top             =   855
            Width           =   1890
         End
         Begin VB.OptionButton optUseGRID 
            Caption         =   "Use GRID: "
            Enabled         =   0   'False
            Height          =   195
            Left            =   285
            TabIndex        =   23
            Top             =   450
            Value           =   -1  'True
            Width           =   1365
         End
      End
      Begin VB.CheckBox chkCalcErosion 
         Caption         =   "Calculate Erosion for Annual Type Precipitation Scenario"
         Height          =   285
         Left            =   -74790
         TabIndex        =   15
         Top             =   555
         Width           =   4965
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLU 
         Height          =   2295
         Left            =   -74820
         TabIndex        =   13
         Top             =   480
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   4
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCoeffs 
         Height          =   2295
         Left            =   -74820
         TabIndex        =   27
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   7
         BackColorSel    =   12648447
         HighLight       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdLCChanges 
         Height          =   2295
         Left            =   180
         TabIndex        =   34
         Top             =   480
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   4
         HighLight       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.Label Label7 
         Caption         =   "K Factor Dataset:"
         Height          =   315
         Left            =   -74760
         TabIndex        =   48
         Top             =   915
         Width           =   1260
      End
      Begin VB.Label Label3 
         Caption         =   "Hydrologic Soil Group Attribute:"
         Height          =   270
         Left            =   -74760
         TabIndex        =   45
         Top             =   2640
         Visible         =   0   'False
         Width           =   2700
      End
      Begin VB.Label lblKFactor 
         Height          =   315
         Left            =   -73425
         TabIndex        =   43
         Top             =   915
         Visible         =   0   'False
         Width           =   3810
      End
      Begin VB.Label lblErodFactor 
         Caption         =   "Erodibility Factor Attribute: "
         Height          =   225
         Left            =   -68400
         TabIndex        =   28
         Top             =   2760
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6885
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6705
      Width           =   1080
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8055
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6705
      Width           =   1080
   End
   Begin MSComDlg.CommonDialog dlgXML 
      Left            =   9375
      Top             =   4410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkIgnore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   1155
      TabIndex        =   30
      Top             =   5385
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CheckBox chkIgnoreMgmt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   8025
      TabIndex        =   35
      Top             =   5220
      Width           =   225
   End
   Begin VB.CheckBox chkIgnoreLU 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   7005
      TabIndex        =   36
      Top             =   5400
      Width           =   225
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   3225
      Left            =   5040
      ScaleHeight     =   3225
      ScaleWidth      =   4575
      TabIndex        =   49
      Top             =   720
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Results Output "
      Height          =   780
      Left            =   180
      TabIndex        =   16
      Top             =   5745
      Visible         =   0   'False
      Width           =   9210
      Begin VB.TextBox txtThemeName 
         Height          =   285
         Left            =   6330
         MaxLength       =   30
         TabIndex        =   19
         Top             =   330
         Width           =   2310
      End
      Begin VB.CommandButton cmdOutputBrowse 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4470
         Picture         =   "frmPrj.frx":173A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   315
         Width           =   345
      End
      Begin VB.TextBox txtOutputFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1500
         TabIndex        =   17
         Text            =   "C:\Temp\test.shp"
         ToolTipText     =   "Functionality to be finalized in Alpha2"
         Top             =   330
         Width           =   2940
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Layer Name:"
         Height          =   210
         Index           =   11
         Left            =   5235
         TabIndex        =   21
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Output Shapefile:"
         Height          =   285
         Index           =   12
         Left            =   90
         TabIndex        =   20
         ToolTipText     =   "Functionality to be finalized in Alpha2"
         Top             =   330
         Width           =   1500
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuBigHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGeneralHelp 
         Caption         =   "N-SPECT Help..."
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuLUAdd 
         Caption         =   "Add Scenario..."
      End
      Begin VB.Menu mnuLUEdit 
         Caption         =   "Edit Scenario..."
      End
      Begin VB.Menu mnuLUDelete 
         Caption         =   "Delete Scenario..."
      End
   End
   Begin VB.Menu mnuManagement 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuManAppen 
         Caption         =   "Append Row"
      End
      Begin VB.Menu mnuManInsert 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuManDelete 
         Caption         =   "Delete Current Row"
      End
   End
End
Attribute VB_Name = "frmPrj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmPrj
' *************************************************************************************
' *  Description:  The big one.  The main form called from Analysis... menu choice.
' *
' *
' *  Called By:  frmPrj
' *************************************************************************************

Option Explicit

Private m_strFileName As String                 'Name of Open doc
Private m_XMLPrjParams As clsXMLPrjFile         'xml doc that holds inputs
Private m_bolFirstLoad As Boolean               'Is initial Load event
Private m_booNew As Boolean                     'New
Private m_booExists As Boolean                  'Has file been saved
Private m_strOpenFileName As String             'String to hold open file name, if they change name, prompt to 'save as'
Private m_strWorkspace As String                'String holding workspace, set it
Private m_booAnnualPrecip As Boolean            'Is the precip scenario annual, if so = TRUE
                            
Private m_intCount As Integer
Private m_intMgmtCount As Integer               'Count for management scenarios
Private m_intLUCount As Integer                 'Count for Land Use grid
Private m_intPollCount As Integer               'Count for Pollutants in Pollutants tab
Private m_intPollRow As Integer                 'Row for Pollutant grid
Private m_intPollCol As Integer                 'Col for Pollutant grid
Private m_intLCRow As Integer                   'Row for LCChange Grid
Private m_intLCCol As Integer                   'Col for LCChange Grid
Private m_intLURow As Integer                   'Row for mgmt scenarios
Private m_intLUCol As Integer                   'col for mgmt scenarios

Private m_strType As String                     'Flag for deletion/creation of checkboxes
Private m_strPrecipFile As String
Private m_strWShed As String                    'String

'Font DPI API
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

' Win 32 Constant Declarations
Private Const LOGPIXELSX = 88                   'Logical pixels/inch in X

'ArcMap stuff
Private m_pMap As IMap                          'Ref to ArcMap.FocusMap
Private m_pMxDoc As IMxDocument
Private m_ParentHWND As Long                    'Set this to get correct parenting of Error handler forms

'Public stuff
Public m_App As IApplication                    'Ref to ArcMap
Public WithEvents m_pActiveViewEvents As Map    'Active View Events for tracking added data layers
Attribute m_pActiveViewEvents.VB_VarHelpID = -1

' The code file for "cBrowseForFolder" is copied from vbAccelerator.
' Link: http://vbaccelerator.com/zip.asp?id=5160
'Public WithEvents dlgBrowser As cBrowseForFolder 'Browse for folder add in: sets output directory

' Constant used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmPrj.frm"

' The following is a workaround for modeless VB forms.
' See: http://resources.esri.com/help/9.3/arcgisdesktop/com/COM/VB6/ModelessVBDialogs.htm
Private m_Frame As IModelessFrame
Public Function Frame() As IModelessFrame
   If m_Frame Is Nothing Then
     Set m_Frame = New ModelessFrame
     m_Frame.Create Me
     m_Frame.Caption = Me.Caption
   End If
   Set Frame = m_Frame
End Function

Public Sub init(ByVal pApp As IApplication)
  On Error GoTo ErrorHandler

    'Called before form loads, means to initialize pMap and the ActiveView events
    
    Set m_App = pApp
    Set m_pMxDoc = m_App.Document
    Set m_pMap = m_pMxDoc.FocusMap
    Set m_pActiveViewEvents = m_pMap  'Set up to catch add layer event
    
  Exit Sub
ErrorHandler:
  HandleError True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboAreaLayer_Click()
'The floating combobox used in the Management scenarios, is populated with all polygon layers in the current map
  On Error GoTo ErrorHandler
    
    grdLCChanges.TextMatrix(m_intLCRow, m_intLCCol) = cboAreaLayer.Text
    cboAreaLayer.Visible = False
    
    grdLCChanges.TextMatrix(0, 0) = ""
  
  Exit Sub

ErrorHandler:
  HandleError True, "cboAreaLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboClass_Click()
'The floating comboxbox used in the Management Scenarios grid, is populated with all polygon layers
  On Error GoTo ErrorHandler

    grdLCChanges.TextMatrix(m_intLCRow, m_intLCCol) = cboClass.Text
    cboClass.Visible = False

  Exit Sub

ErrorHandler:
  HandleError True, "cboClass_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboCoeff_Click()
'The floating combobox used in the Pollutant tab for Coefficient
  On Error GoTo ErrorHandler
        
    grdCoeffs.TextMatrix(m_intPollRow, m_intPollCol) = cboCoeff.Text
    cboCoeff.Visible = False
    
    If cboCoeff.ListIndex = 4 Then
        
        g_intCoeffRow = m_intPollRow 'Global set up to hold what row we're on
        g_strCoeffCalc = grdCoeffs.TextMatrix(g_intCoeffRow, 6)
              
        frmPrjCalc.init m_App
        frmPrjCalc.Show vbModal, Me
        
    End If

  Exit Sub

ErrorHandler:
  HandleError True, "cboCoeff_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboCoeffSet_Click()
'The floating combobox used in the Pollutant tab for Coefficient Set
  On Error GoTo ErrorHandler

    'Set the text of column 3 to the selected cbo Text
    grdCoeffs.TextMatrix(m_intPollRow, m_intPollCol) = cboCoeffSet.Text
    
    'Added 8/17/04; set Type to 'Type 1' as default
    grdCoeffs.TextMatrix(m_intPollRow, m_intPollCol + 1) = "Type 1"
    
    cboCoeffSet.Visible = False

  Exit Sub

ErrorHandler:
  HandleError True, "cboCoeffSet_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboLCLayer_Click()
'Landclass layer combobox
  On Error GoTo ErrorHandler
    
    cboLCUnits.ListIndex = modUtil.GetRasterDistanceUnits(cboLCLayer.Text, m_App)

  Exit Sub

ErrorHandler:
  HandleError True, "cboLCLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboLCType_Click()
'Landclass type

  On Error GoTo ErrorHandler

    FillCboLCCLass

  Exit Sub

ErrorHandler:
  HandleError True, "cboLCType_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboPrecipScen_Click()
'Combobox for Precip Scenarios
  On Error GoTo ErrorHandler

    'Have to change Erosion tab based on Annual/Event driven rain event
    Dim rsEvent As New ADODB.Recordset
    Dim strEvent As String
    
    'If define, then open new window for new definition, else select from database
    If cboPrecipScen.Text = "New precipitation scenario..." Then
        Load frmNewPrecip
        frmNewPrecip.Show vbModal, Me
    Else
        strEvent = "SELECT * FROM PRECIPSCENARIO WHERE NAME LIKE '" & cboPrecipScen.Text & "'"
        rsEvent.Open strEvent, g_ADOConn, adOpenStatic
        
        Select Case rsEvent!Type
            Case 0  'Annual
                frmSDR.Visible = True
                frameRainFall.Visible = True
                chkCalcErosion.Caption = "Calculate Erosion for Annual Type Precipitation Scenario"
                m_booAnnualPrecip = True   'Set flag
            Case 1  'Event
                frmSDR.Visible = False
                frameRainFall.Visible = False
                chkCalcErosion.Caption = "Calculate Erosion for Event Type Precipitation Scenario"
                m_booAnnualPrecip = False  'Set flag
        End Select
        
        m_strPrecipFile = rsEvent!PrecipFileName
        
        rsEvent.Close
        Set rsEvent = Nothing
    
    End If
    

  Exit Sub
ErrorHandler:
  HandleError True, "cboPrecipScen_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboSoilsLayer_Click()
'Soils layer combobox
  On Error GoTo ErrorHandler


    Dim rsSoils As New ADODB.Recordset
    
    Dim strSelect As String
    strSelect = "SELECT * FROM Soils WHERE NAME LIKE '" & cboSoilsLayer.Text & "'"
    
    rsSoils.Open strSelect, g_ADOConn, adOpenStatic
    
    lblKFactor.Caption = rsSoils!SoilsKFileName
    lblSoilsHyd.Caption = rsSoils!SoilsFileName

    'clean
    rsSoils.Close
    Set rsSoils = Nothing



  Exit Sub
ErrorHandler:
  HandleError True, "cboSoilsLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboWQStd_Click()
  On Error GoTo ErrorHandler

    Dim i As Integer
    Dim j As Integer
        
    If cboWQStd.Text = "New water quality standard..." Then
        Load frmAddWQStd
        frmAddWQStd.Show vbModal, Me
    Else
        Timer1.Interval = 10
        Timer1.Enabled = True
        PopPollutants
      
    End If
    
    For i = 1 To grdCoeffs.Rows - 1
        grdCoeffs.TextMatrix(i, 3) = ""
        grdCoeffs.TextMatrix(i, 4) = ""
        
    Next i
    

  Exit Sub
ErrorHandler:
  HandleError True, "cboWQStd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cboWSDelin_Click()
  On Error GoTo ErrorHandler

    
    If cboWSDelin.Text = "New watershed delineation..." Then
        
        g_boolNewWShed = True
        frmNewWSDelin.m_booProject = True
        frmNewWSDelin.init m_App
        frmNewWSDelin.Show vbModal, Me
        
    End If


  Exit Sub
ErrorHandler:
  HandleError True, "cboWSDelin_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub chkCalcErosion_Click()
  On Error GoTo ErrorHandler

    frameRainFall.Enabled = chkCalcErosion.Value
    cboErodFactor.Enabled = chkCalcErosion.Value
    lblErodFactor.Enabled = chkCalcErosion.Value
    optUseGRID.Enabled = chkCalcErosion.Value
    optUseValue.Enabled = True 'chkCalcErosion.Value
    lblKFactor.Visible = chkCalcErosion.Value
    Label7.Visible = chkCalcErosion.Value
    

  Exit Sub
ErrorHandler:
  HandleError True, "chkCalcErosion_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub chkIgnore_Click(Index As Integer)
  On Error GoTo ErrorHandler

    'Ignore column
    grdCoeffs.TextMatrix(Index, 1) = chkIgnore(Index).Value
    
  Exit Sub
ErrorHandler:
  HandleError True, "chkIgnore_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub chkIgnoreLU_Click(Index As Integer)
  On Error GoTo ErrorHandler

    'Ignore column value
    grdLU.TextMatrix(Index, 1) = chkIgnoreLU(Index).Value
    
  Exit Sub
ErrorHandler:
  HandleError True, "chkIgnoreLU_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub chkIgnoreMgmt_Click(Index As Integer)
  On Error GoTo ErrorHandler

    'Ignore column value
    grdLCChanges.TextMatrix(Index, 1) = chkIgnoreMgmt(Index).Value

  Exit Sub
ErrorHandler:
  HandleError True, "chkIgnoreMgmt_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub chkSDR_Click()
    If chkSDR.Value = 1 Then
        txtSDRGRID.Enabled = True
        cmdOpenSDR.Enabled = True
    Else
        txtSDRGRID.Enabled = False
        cmdOpenSDR.Enabled = False
    End If
End Sub

Private Sub chkSelectedPolys_Click()
  On Error GoTo ErrorHandler

    If chkSelectedPolys.Value = 1 Then
        cboSelectPoly.Enabled = True
        lblLayer.Enabled = True
    Else
        cboSelectPoly.Enabled = False
        lblLayer.Enabled = False
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "chkSelectedPolys_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmdOpenSDR_Click()
    
    Dim pRasterDataset As IRasterDataset
    
    On Error GoTo ErrHandler:
    Set pRasterDataset = AddInputFromGxBrowserText(txtSDRGRID, "Choose SDR GRID", frmPrj, 0)

Exit Sub
ErrHandler:
    MsgBox "There was an error opening the selected file.", vbCritical, "Error Opening File"
End Sub

Private Sub cmdRun_Click()
  On Error GoTo ErrorHandler

    Dim rsWaterShed As New ADODB.Recordset          'RS to get WaterShed information
    Dim strWaterShed As String                      'Connection string
    Dim rsPrecip As New ADODB.Recordset             'RS to get Precip info
    Dim strPrecip As String                         'Connection String
    Dim pTempRaster As IRaster                      'Temp raster used to delete all globals at completion
    Dim pSelectedPolyLayer As ILayer                'Selected polygon layer.
    Dim lngWShedLayerIndex As Long                  'Watershed layer index
    
    Dim booLUItems As Boolean                       'Are there Landuse Scenarios???
    Dim dictPollutants As New Dictionary            'Dict to hold all pollutants
    Dim i As Long
    Dim strProjectInfo As String                    'String that will hold contents of prj file for inclusion in metatdata

    Screen.MousePointer = vbHourglass
    
    'STEP 1: Save file, populate xml params: -------------------------------------------------------------------------
    If Not SaveXMLFile Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    'Init your global dictionary to hold the metadata records as well as the global xml prj file
    Set g_dicMetadata = New Dictionary
    Set g_clsXMLPrjFile = m_XMLPrjParams
    'END STEP 1: -----------------------------------------------------------------------------------------------------
    
    
    'STEP 2: Identify if local effects are being used : --------------------------------------------------------------
    'Local Effects Global
    If m_XMLPrjParams.intLocalEffects = 1 Then
        g_booLocalEffects = True
    Else
        g_booLocalEffects = False
    End If
    'END STEP 2: -----------------------------------------------------------------------------------------------------
    
    'STEP 3: Find out if user is making use of only the selected Sheds -----------------------------------------------
    'Selected Sheds only
    If m_XMLPrjParams.intSelectedPolys = 1 Then
        g_booSelectedPolys = True
        lngWShedLayerIndex = modUtil.GetLayerIndex(cboSelectPoly.Text, m_App)
        Set pSelectedPolyLayer = m_pMap.Layer(lngWShedLayerIndex)
    Else
        g_booSelectedPolys = False
    End If
    'END STEP 3: ---------------------------------------------------------------------------------------------------------
    
    'Check for Spatial Analyst Extension, doing it once here to eliminate multiple checks down the road
    If Not CheckSpatialAnalystLicense Then
        MsgBox "Spatial Analyst Extension is required for N-SPECT and a license is not available.", vbCritical, "Spatial Analyst Required"
        Exit Sub
    End If
    
    'STEP 4: Get the Management Scenarios: ------------------------------------------------------------------------------------
    'If they're using, we send them over to modMgmtScen to implement
    If m_XMLPrjParams.clsMgmtScenHolder.Count > 0 Then
        modMgmtScen.MgmtScenSetup m_XMLPrjParams.clsMgmtScenHolder, m_XMLPrjParams.strLCGridType, m_XMLPrjParams.strLCGridFileName, m_XMLPrjParams.strProjectWorkspace
    End If
    'END STEP 4: ---------------------------------------------------------------------------------------------------------
    
    'STEP 5: Pollutant Dictionary creation, needed for Landuse -----------------------------------------------------------
    'Go through and find the pollutants, if they're used and what the CoeffSet is
    'We're creating a dictionary that will hold Pollutant, Coefficient Set for use in the Landuse Scenarios
    For i = 1 To m_XMLPrjParams.clsPollItems.Count
        If m_XMLPrjParams.clsPollItems.Item(i).intApply = 1 Then
            dictPollutants.Add m_XMLPrjParams.clsPollItems.Item(i).strPollName, m_XMLPrjParams.clsPollItems.Item(i).strCoeffSet
        End If
    Next i
    'END STEP 5: ---------------------------------------------------------------------------------------------------------
    
    'STEP 6: Landuses sent off to modLanduse for processing -----------------------------------------------------
    For i = 1 To m_XMLPrjParams.clsLUItems.Count
        If m_XMLPrjParams.clsLUItems.Item(i).intApply = 1 Then
            booLUItems = True
            modLanduse.Begin m_XMLPrjParams.strLCGridType, m_XMLPrjParams.clsLUItems, dictPollutants, m_XMLPrjParams.strLCGridFileName, m_pMap, m_XMLPrjParams.strProjectWorkspace
            Exit For
        Else
            booLUItems = False
        End If
    Next i
    'END STEP 6: ---------------------------------------------------------------------------------------------------------
    
    'STEP 7: ---------------------------------------------------------------------------------------------------------
    'Obtain Watershed values
    
    strWaterShed = "Select * from WSDelineation Where Name like '" & m_XMLPrjParams.strWaterShedDelin & "'"
    rsWaterShed.Open strWaterShed, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'END STEP 7: -----------------------------------------------------------------------------------------------------

    'STEP 8: ---------------------------------------------------------------------------------------------------------
    'Set the Analysis Environment and globals for output workspace
    
    modMainRun.SetGlobalEnvironment rsWaterShed, m_XMLPrjParams.strProjectWorkspace, m_pMap, pSelectedPolyLayer
    
    'END STEP 8: -----------------------------------------------------------------------------------------------------

    'STEP 8a: --------------------------------------------------------------------------------------------------------
    'Added 1/08/2007 to account for non-adjacent polygons
    If m_XMLPrjParams.intSelectedPolys = 1 Then
        Dim pPolygon As IPolygon4
        Set pPolygon = g_pSelectedPolyClip
        If modMainRun.CheckMultiPartPolygon(pPolygon) Then
            MsgBox "Warning: Your selected polygons are not adjacent.  Please select only polygons that are adjacent.", vbCritical, "Non-adjacent Polygons Detected"
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
    End If

    'STEP 9: ---------------------------------------------------------------------------------------------------------
    'Create the runoff GRID
    'Get the precip scenario stuff
    strPrecip = "Select * from PrecipScenario where name like '" & m_XMLPrjParams.strPrecipScenario & "'"
    rsPrecip.Open strPrecip, g_ADOConn, adOpenDynamic, adLockOptimistic

    'If there has been a land use added, then a new LCType has been created, hence we get it from g_strLCTypename
    Dim strLCType As String
    If booLUItems Then
        strLCType = modLanduse.g_strLCTypeName
    Else
        strLCType = m_XMLPrjParams.strLCGridType
    End If

    'Added 6/04 to account for different PrecipTypes
    modMainRun.g_intPrecipType = rsPrecip!PrecipType
    
    If Not modRunoff.CreateRunoffGrid(m_XMLPrjParams.strLCGridFileName, strLCType, rsPrecip, _
                                        m_XMLPrjParams.strSoilsHydFileName) Then
        Exit Sub
    End If
    'END STEP 9: -----------------------------------------------------------------------------------------------------
    
    'STEP 10: ---------------------------------------------------------------------------------------------------------
    'Process pollutants
    For i = 1 To m_XMLPrjParams.clsPollItems.Count
        If m_XMLPrjParams.clsPollItems.Item(i).intApply = 1 Then
            'If user is NOT ignoring the pollutant then send the whole item over along with LCType
            If Not modPollutantCalcs.PollutantConcentrationSetup(m_XMLPrjParams.clsPollItems.Item(i), m_XMLPrjParams.strLCGridType, m_XMLPrjParams.strWaterQuality) Then
                Exit Sub
            End If
        End If
    Next i
    'END STEP 10: -----------------------------------------------------------------------------------------------------
    
    'Step 11: Erosion -------------------------------------------------------------------------------------------------
    'Check that they have chosen Erosion
    If m_XMLPrjParams.intCalcErosion = 1 Then
        If m_booAnnualPrecip Then      'If Annual (0) then TRUE, ergo RUSLE
            If m_XMLPrjParams.intRainGridBool Then
                If Not modRusle.RUSLESetup(rsWaterShed!NibbleFileName, rsWaterShed!dem2bfilename, m_XMLPrjParams.strRainGridFileName, _
                        m_XMLPrjParams.strSoilsKFileName, m_XMLPrjParams.strSDRGridFileName, m_XMLPrjParams.strLCGridType) Then
                        Exit Sub
                End If
            ElseIf m_XMLPrjParams.intRainConstBool Then
                If Not modRusle.RUSLESetup(rsWaterShed!NibbleFileName, rsWaterShed!dem2bfilename, m_XMLPrjParams.strRainGridFileName, _
                    m_XMLPrjParams.strSoilsKFileName, m_XMLPrjParams.strSDRGridFileName, m_XMLPrjParams.strLCGridType, m_XMLPrjParams.dblRainConstValue) Then
                    Exit Sub
                End If
            End If
                
         Else   'If event (1) then False, ergo MUSLE
            If Not modMUSLE.MUSLESetup(m_XMLPrjParams.strSoilsDefName, m_XMLPrjParams.strSoilsKFileName, m_XMLPrjParams.strLCGridType) Then
                Exit Sub
            End If
        End If
    End If
    'STEP 11: ----------------------------------------------------------------------------------------------------------
        
    'STEP 12 : Cleanup any temp critters -------------------------------------------------------------------------------
    'g_DictTempNames holds the names of all temporary landuses and/or coefficient sets created during the Landuse scenario
    'portion of our program, for example CCAP1, or NitSet1.  We now must eliminate them from the database if they exist.
    If g_DictTempNames.Count > 0 Then
        If booLUItems Then
            modLanduse.Cleanup g_DictTempNames, m_XMLPrjParams.clsPollItems, m_XMLPrjParams.strLCGridType
        End If
    End If
    'END STEP 12: -------------------------------------------------------------------------------------------------------
    
    'STEP 13: -----------------------------------------------------------------------------------------------------------
    'g_pGroupLayer has been created earlier and has been taken on GRIDs since.  Now lets add it
    'Add the group layer.
    With g_pGroupLayer
        .Expanded = True                        'Are going to 'expand' it
        .Name = m_XMLPrjParams.strProjectName   'The name equals whatever the user entered
    End With

    m_pMap.AddLayer g_pGroupLayer
    'END STEP 13: -------------------------------------------------------------------------------------------------------
    
    'STEP 14 save out group layer ---------------------------------------------------------------------------------------
    modUtil.ExportLayerToPath g_pGroupLayer, m_XMLPrjParams.strProjectWorkspace & "\" & m_XMLPrjParams.strProjectName & ".lyr"
    'END STEP 14: -------------------------------------------------------------------------------------------------------
    
    'STEP 15: create string describing project parameters ---------------------------------------------------------------
    strProjectInfo = modUtil.ParseProjectforMetadata(m_XMLPrjParams, m_strFileName)
    'END STEP 15: -------------------------------------------------------------------------------------------------------
    
    'STEP 16: Apply the metadata to each of the rasters in the group layer ----------------------------------------------
    m_App.StatusBar.Message(0) = "Creating metadata for the N-SPECT group layer..."
    modUtil.CreateMetadata g_pGroupLayer, strProjectInfo
    'END STEP 16: -------------------------------------------------------------------------------------------------------
    
    'Cleanup ------------------------------------------------------------------------------------------------------------
    m_App.StatusBar.Message(0) = "Deleting temporary files..."
    rsWaterShed.Close
    rsPrecip.Close
    Set rsWaterShed = Nothing
    Set rsPrecip = Nothing
    
    'Go into workspace and rid it of all rasters
    modUtil.CleanGlobals
    modUtil.CleanupRasterFolder m_XMLPrjParams.strProjectWorkspace
    
    Set g_pGroupLayer = Nothing
    Set g_DictTempNames = Nothing
    Set g_dicMetadata = Nothing
    
    Screen.MousePointer = vbNormal
    
    m_App.StatusBar.Message(0) = "N-SPECT processing complete!"
    
    Unload frmPrj
    
Exit Sub
    
UserCancel:
    modProgDialog.KillDialog
    Screen.MousePointer = vbNormal
    MsgBox "Processing has been stopped.", vbInformation, "Analysis Stopped"
    
ErrorHandler:
    HandleError True, "cmdRun_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    Screen.MousePointer = vbNormal
End Sub

Private Sub cmdSave_Click()
  On Error GoTo ErrorHandler

    dlgXML.FileName = Empty
    With dlgXML
     
     .Filter = MSG8
     .DialogTitle = MSG2
     .FilterIndex = 1
     .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
     .ShowSave
    
    End With
   
    If Len(dlgXML.FileName) > 0 Then
        m_strFileName = Trim(dlgXML.FileName)
        m_XMLPrjParams.SaveFile m_strFileName
    End If

  Exit Sub
ErrorHandler:
  HandleError False, "cmdSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub LoadXMLFile()
  'On Error GoTo ErrorHandler

     'browse...get output filename
    Dim fso As FileSystemObject
    Dim strFolder As String
    
    Set fso = New FileSystemObject
    strFolder = modUtil.g_nspectDocPath & "\projects"
    If Not fso.FolderExists(strFolder) Then
        MkDir strFolder
    End If

   dlgXML.FileName = Empty
   With dlgXML
     .Filter = MSG8
     .InitDir = strFolder
     .DialogTitle = "Open N-SPECT Project File"
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .ShowOpen
   End With
   
   If Len(dlgXML.FileName) > 0 Then
      m_strFileName = Trim(dlgXML.FileName)
      m_XMLPrjParams.XML = m_strFileName
      FillForm
   Else
      Exit Sub
   End If
   
   'Pop this string with the incoming name, if they change, we'll prompt to 'save as'.
   m_strOpenFileName = txtProjectName.Text
   
   Set fso = Nothing

  'Exit Sub
'ErrorHandler:
'  HandleError False, "LoadXMLFile " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Function SaveXMLFile() As Boolean
    
On Error GoTo ErrHandler

    Dim strFolder As String
    Dim intvbYesNo As Integer
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    
    strFolder = modUtil.g_nspectDocPath & "\projects"
    If Not fso.FolderExists(strFolder) Then
        MkDir strFolder
    End If
    
    If Not ValidateData Then 'check form inputs
        SaveXMLFile = False
        Exit Function
    End If
        
    'If it does not already exist, open Save As... dialog
    If Not m_booExists Then
        With dlgXML
          .Filter = MSG8
          .DialogTitle = "Save Project File As..."
          .InitDir = strFolder
          .FilterIndex = 1
          .Flags = cdlOFNOverwritePrompt + cdlOFNFileMustExist + cdlOFNPathMustExist
          .CancelError = True
          .FileName = txtProjectName.Text
          .ShowSave
        End With
        'check to make sure filename length is greater than zeros
        If Len(dlgXML.FileName) > 0 Then
            m_strFileName = Trim(dlgXML.FileName)
            m_booExists = True
            m_XMLPrjParams.SaveFile m_strFileName
            SaveXMLFile = True
        Else
            SaveXMLFile = False
            Exit Function
        End If
    
    Else
        'Now check to see if the name changed
        If m_strOpenFileName <> txtProjectName.Text Then
            intvbYesNo = MsgBox("You have changed the name of this project.  Would you like to save your settings as a new file?" & vbNewLine & _
                vbTab & "Yes" & vbTab & " -    Save as new N-SPECT project file" & vbNewLine & _
                vbTab & "No" & vbTab & " -    Save changes to current N-SPECT project file" & vbNewLine & _
                vbTab & "Cancel" & vbTab & " -    Return to the project window", vbQuestion + vbYesNoCancel, "N-SPECT")
            
            If intvbYesNo = vbYes Then
                With dlgXML
                    .Filter = MSG8
                    .DialogTitle = "Save Project File As..."
                    .InitDir = strFolder
                    .FilterIndex = 1
                    .Flags = cdlOFNOverwritePrompt + cdlOFNFileMustExist + cdlOFNPathMustExist
                    .CancelError = True
                    .FileName = txtProjectName.Text
                    .ShowSave
                End With
                'check to make sure filename length is greater than zeros
                If Len(dlgXML.FileName) > 0 Then
                    m_strFileName = Trim(dlgXML.FileName)
                    m_booExists = True
                    m_XMLPrjParams.SaveFile m_strFileName
                    SaveXMLFile = True
                Else
                    SaveXMLFile = False
                    Exit Function
                End If
            ElseIf intvbYesNo = vbNo Then
                m_XMLPrjParams.SaveFile m_strFileName
                m_booExists = True
                SaveXMLFile = True
            Else
                SaveXMLFile = False
                Exit Function
            End If
        Else
            m_XMLPrjParams.SaveFile m_strFileName
            m_booExists = True
            SaveXMLFile = True
            
        End If
    
    End If
           
    Set fso = Nothing
    
Exit Function
    
ErrHandler:
    
    If Err.Number = 32755 Then
        SaveXMLFile = False
        Exit Function
    Else
        MsgBox Err.Number & " " & Err.Description
        SaveXMLFile = False
    End If

End Function

Private Sub FillForm()
  'On Error GoTo ErrorHandler
On Error Resume Next
    Dim rsCurrWShed As New ADODB.Recordset
    Dim strCurrWShed As String
    Dim pCurrWShedPolyLayer As ILayer
    Dim lngCurrWshedPolyIndex As Long
    Dim intYesNo As Integer
    Dim pDBasinFeatureLayer As IFeatureLayer
    Dim pDBasinDataset As IDataset
    Dim strBasinPoly As String
    Dim i As Integer
    Dim z As Integer
    Dim booNameMatch
    Dim fso As New FileSystemObject
    
    Screen.MousePointer = vbHourglass
    
    txtProjectName.Text = m_XMLPrjParams.strProjectName
    txtOutputWS.Text = m_XMLPrjParams.strProjectWorkspace
    
    'Step 1:  LandCoverGrid
    'Check to see if the LC cover is in the map, if so, set the combobox
    If modUtil.LayerInMap(m_XMLPrjParams.strLCGridName, m_pMap) Then
        cboLCLayer.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strLCGridName, cboLCLayer)
        cboLCLayer.Refresh
    Else
        If fso.FileExists(m_XMLPrjParams.strLCGridFileName & ".lyr") Then
        
            modUtil.AddLayerFileToMap m_XMLPrjParams.strLCGridFileName & ".lyr", m_pMap
            
        ElseIf modUtil.AddRasterLayerToMapFromFileName(m_XMLPrjParams.strLCGridFileName, m_pMap) Then
        
            With cboLCLayer
                '.AddItem m_XMLPrjParams.strLCGridName
                .Refresh
                '.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strLCGridName, cboLCLayer)
            End With
        
        Else
            intYesNo = MsgBox("Could not find the Land Cover dataset: " & m_XMLPrjParams.strLCGridFileName & ".  Would you like " & _
            "to browse for it?", vbYesNo, "Cannot Locate Dataset")
            If intYesNo = vbYes Then
                m_XMLPrjParams.strLCGridFileName = modUtil.AddInputFromGxBrowser(cboLCLayer, frmPrj, "Raster")
                If m_XMLPrjParams.strLCGridFileName <> "" Then
                    If modUtil.AddRasterLayerToMapFromFileName(m_XMLPrjParams.strLCGridFileName, m_pMap) Then
                        cboLCLayer.ListIndex = 0
                    End If
                Else
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
       End If
       
    End If
        
    cboLCUnits.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strLCGridUnits, cboLCUnits)
    cboLCUnits.Refresh
    
    cboLCType.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strLCGridType, cboLCType)
    cboLCType.Refresh
    
    'Step 2: Soils - same process, if in doc and map, OK, else send em looking
    If modUtil.RasterExists(m_XMLPrjParams.strSoilsHydFileName) Then
        cboSoilsLayer.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strSoilsDefName, cboSoilsLayer)
        cboSoilsLayer.Refresh
    Else
        MsgBox "Could not find soils dataset.  Please correct the soils definition in the Advanced Settings.", vbCritical, "Dataset Missing"
        Exit Sub
    End If
    
    'Step5: Precip Scenario
    cboPrecipScen.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strPrecipScenario, cboPrecipScen)
    
    'Step6: Watershed Delineation
    cboWSDelin.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strWaterShedDelin, cboWSDelin)
    
    'Add the basinpoly to the map
    strCurrWShed = "Select * from WSDelineation where Name Like '" & m_XMLPrjParams.strWaterShedDelin & "'"
    rsCurrWShed.Open strCurrWShed, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    
    '**********************************************************************************************
    '**********************************************************************************************
    'Check to see if the Drainage Basins poly is already in the map
'    For z = 0 To m_pMap.LayerCount - 1
'
'        If TypeOf m_pMap.Layer(z) Is IFeatureLayer Then
'            Set pDBasinFeatureLayer = m_pMap.Layer(z)
'            Set pDBasinDataset = pDBasinFeatureLayer
'            strBasinPoly = Trim(pDBasinDataset.Workspace.PathName & "\" & pDBasinFeatureLayer.FeatureClass.AliasName)
'
'            If strBasinPoly <> rsCurrWShed!wsfilename Then
'                booNameMatch = False
'            Else
'                booNameMatch = True
'                Exit For
'            End If
'        End If
'    Next z
    
    
    If Not modUtil.LayerInMapByFileName(rsCurrWShed!wsfilename, m_pMap) Then
        If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed!wsfilename, m_pMap, rsCurrWShed!Name & " " & "Drainage Basins") Then
            lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App)
            m_pMxDoc.ContentsView(0).Refresh m_pMap.Layer(lngCurrWshedPolyIndex)
        Else
            MsgBox "Could not find watershed layer: " & rsCurrWShed!wsfilename & " .  Please add the watershed layer to the map.", vbCritical, "File Not Found"
        End If
    End If
        
            
'    If Not booNameMatch Then
'        If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed!wsfilename, m_pMap, rsCurrWShed!Name & " " & "Drainage Basins") Then
'            lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App)
'            m_pMxDoc.ContentsView(0).Refresh m_pMap.Layer(lngCurrWshedPolyIndex)
'        End If
'    End If
    
   
    '**********************************************************************************************
    '**********************************************************************************************
'    If modUtil.LayerInMap(rsCurrWShed!Name & " " & "Drainage Basins", m_pMap) Then
'        Set pDBasinFeatureLayer = m_pMap.Layer(modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App))
'        Set pDBasinDataset = pDBasinFeatureLayer
'        strBasinPoly = Trim(pDBasinDataset.Workspace.PathName & "\" & pDBasinFeatureLayer.FeatureClass.AliasName)
'        'if it is not the same one then add it
'        If strBasinPoly <> rsCurrWShed!wsfilename Then
'            If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed!wsfilename, m_pMap, rsCurrWShed!Name & " " & "Drainage Basins") Then
'                lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App)
'                m_pMxDoc.ContentsView(0).Refresh m_pMap.Layer(lngCurrWshedPolyIndex)
'            End If
'        End If
'    Else
'        'if not in the map do straight add in
'        If modUtil.AddFeatureLayerToMapFromFileName(rsCurrWShed!wsfilename, m_pMap, rsCurrWShed!Name & " " & "Drainage Basins") Then
'            lngCurrWshedPolyIndex = modUtil.GetLayerIndex(rsCurrWShed!Name & " " & "Drainage Basins", m_App)
'            m_pMxDoc.ContentsView(0).Refresh m_pMap.Layer(lngCurrWshedPolyIndex)
'        End If
'    End If

    '**********************************************************************************************
    '**********************************************************************************************
    
    'Step7: Water Quality
    cboWQStd.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strWaterQuality, cboWQStd)
    
    'Step8: LocalEffects/Selected Watersheds
    chkLocalEffects.Value = m_XMLPrjParams.intLocalEffects
    chkSelectedPolys.Value = m_XMLPrjParams.intSelectedPolys
    
    If chkSelectedPolys.Value = 1 Then
        '1st see if it's in the map
        If modUtil.LayerInMapByFileName(m_XMLPrjParams.strSelectedPolyFileName, m_pMap) Then
            cboSelectPoly.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strSelectedPolyLyrName, cboSelectPoly)
            cboSelectPoly.Refresh
        Else
            'Not there then add it
            If modUtil.AddFeatureLayerToMapFromFileName(m_XMLPrjParams.strSelectedPolyFileName, m_pMap, _
                m_XMLPrjParams.strSelectedPolyLyrName) Then
                cboSelectPoly.Refresh
                cboSelectPoly.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strSelectedPolyLyrName, cboSelectPoly)
            Else
                'Can't find it, then send em searching
                intYesNo = MsgBox("Could not find the Selected Polygons file used to limit extent: " & _
                           m_XMLPrjParams.strSelectedPolyFileName & ".  Would you like to browse for it? ", _
                           vbYesNo, "Cannot Locate Dataset")
                If intYesNo = vbYes Then
                    'if they want to look for it then give em the browser
                    m_XMLPrjParams.strSelectedPolyFileName = modUtil.AddInputFromGxBrowser(cboSelectPoly, frmPrj, "Feature")
                    'if they actually find something, throw it in the map
                    If Len(m_XMLPrjParams.strSelectedPolyFileName) > 0 Then
                        If modUtil.AddFeatureLayerToMapFromFileName(m_XMLPrjParams.strSelectedPolyFileName, m_pMap) Then
                            cboSelectPoly.ListIndex = modUtil.GetCboIndex(modUtil.SplitFileName(m_XMLPrjParams.strSelectedPolyFileName), cboSelectPoly)
                        End If
                    End If
                Else
                    m_XMLPrjParams.intSelectedPolys = 0
                    chkSelectedPolys.Value = 0
                End If
            End If
        End If
    End If
                            
    'Step: Erosion Tab - Calc Erosion, Erosion Attribute
    If m_XMLPrjParams.intCalcErosion = 1 Then
        chkCalcErosion.Value = 1
    Else
        chkCalcErosion.Value = 0
    End If
    
    'Step: Erosion Tab - Precip
    'Either they use the GRID
    optUseGRID.Value = m_XMLPrjParams.intRainGridBool
    
    If optUseGRID Then
        If modUtil.LayerInMap(m_XMLPrjParams.strRainGridName, m_pMap) Then
          cboRainGrid.ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strRainGridName, cboRainGrid)
          cboRainGrid.Refresh
        Else
            If modUtil.AddRasterLayerToMapFromFileName(m_XMLPrjParams.strRainGridFileName, m_pMap) Then
             With cboRainGrid
                 .AddItem m_XMLPrjParams.strRainGridName
                 .Refresh
                 .ListIndex = modUtil.GetCboIndex(m_XMLPrjParams.strRainGridName, cboRainGrid)
             End With
            Else
                 intYesNo = MsgBox("Could not find Rainfall GRID: " & m_XMLPrjParams.strRainGridName & ".  Would you like " & _
                 "to browse for it?", vbYesNo, "Cannot Locate Dataset")
                 If intYesNo = vbYes Then
                     m_XMLPrjParams.strRainGridFileName = modUtil.AddInputFromGxBrowser(cboRainGrid, frmPrj, "Raster")
                     If modUtil.AddRasterLayerToMapFromFileName(m_XMLPrjParams.strRainGridFileName, m_pMap) Then
                         cboRainGrid.ListIndex = 0
                     End If
                 Else
                     Exit Sub
                 End If
            End If
        End If
    End If
        
    'Or they use a constant value
    optUseValue.Value = m_XMLPrjParams.intRainConstBool
    
    If optUseValue Then
        txtRainValue.Text = m_XMLPrjParams.dblRainConstValue
    End If
    
    'SDR GRID
    
    'If Not m_XMLPrjParams.intUseOwnSDR Is Nothing Then
    On Error GoTo VersionProblem:
        If m_XMLPrjParams.intUseOwnSDR = 1 Then
            chkSDR.Value = 1
            txtSDRGRID.Text = m_XMLPrjParams.strLCGridFileName
        Else
            chkSDR.Value = 0
            txtSDRGRID.Text = m_XMLPrjParams.strSDRGridFileName
        End If
    'End If

    'Step Pollutants
    m_intPollCount = m_XMLPrjParams.clsPollItems.Count
    
    If m_intPollCount > 0 Then
        grdCoeffs.Rows = m_intPollCount + 1
        For i = 1 To m_intPollCount
            With grdCoeffs
                .row = m_XMLPrjParams.clsPollItems.Item(i).intID
                .TextMatrix(.row, 1) = m_XMLPrjParams.clsPollItems.Item(i).intApply
                .TextMatrix(.row, 2) = m_XMLPrjParams.clsPollItems.Item(i).strPollName
                .TextMatrix(.row, 3) = m_XMLPrjParams.clsPollItems.Item(i).strCoeffSet
                .TextMatrix(.row, 4) = m_XMLPrjParams.clsPollItems.Item(i).strCoeff
                .TextMatrix(.row, 5) = CStr(m_XMLPrjParams.clsPollItems.Item(i).intThreshold)
                
                If Len(m_XMLPrjParams.clsPollItems.Item(i).strTypeDefXMLFile) > 0 Then
                    .TextMatrix(.row, 6) = CStr(m_XMLPrjParams.clsPollItems.Item(i).strTypeDefXMLFile)
                End If
            
            End With
         Next i
         
    ClearCheckBoxes True
    CreateCheckBoxes True, True
    End If
    
    'Step - Land Uses
    m_intLUCount = m_XMLPrjParams.clsLUItems.Count
          
    If m_intLUCount > 0 Then
        grdLU.Rows = m_intLUCount + 1
        For i = 1 To m_intLUCount
          With grdLU
            .row = m_XMLPrjParams.clsLUItems.Item(i).intID
            .TextMatrix(.row, 1) = m_XMLPrjParams.clsLUItems.Item(i).intApply
            .TextMatrix(.row, 2) = m_XMLPrjParams.clsLUItems.Item(i).strLUScenName
            .TextMatrix(.row, 3) = m_XMLPrjParams.clsLUItems.Item(i).strLUScenXMLFile
          End With
        Next i
       
    ClearLUCheckBoxes True
    CreateLUCheckBoxes True
    
    End If
    
    'Step Management Scenarios
    m_intMgmtCount = m_XMLPrjParams.clsMgmtScenHolder.Count
    
    If m_intMgmtCount > 0 Then
        grdLCChanges.Rows = m_intMgmtCount + 1
        For i = 1 To m_intMgmtCount
            With grdLCChanges
                .row = m_XMLPrjParams.clsMgmtScenHolder.Item(i).intID
                If modUtil.LayerInMap(m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName, m_pMap) Then
                  .TextMatrix(i, 1) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).intApply
                  .TextMatrix(i, 2) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName
                  .TextMatrix(i, 3) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass
                Else
                    If modUtil.AddFeatureLayerToMapFromFileName(m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName, m_pMap) Then
                        .TextMatrix(i, 1) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).intApply
                        .TextMatrix(i, 2) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName
                        .TextMatrix(i, 3) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass
                    Else
                         intYesNo = MsgBox("Could not find Management Sceario Area Layer: " & _
                         m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName & ".  Would you like " & _
                         "to browse for it?", vbYesNo, "Cannot Locate Dataset:" & m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName)
                         If intYesNo = vbYes Then
                             m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName = modUtil.AddInputFromGxBrowser(cboAreaLayer, frmPrj, "Feature")
                             If modUtil.AddFeatureLayerToMapFromFileName(m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaFileName, m_pMap) Then
                                  cboAreaLayer.ListIndex = 0
                                 .TextMatrix(i, 1) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).intApply
                                 .TextMatrix(i, 2) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).strAreaName
                                 .TextMatrix(i, 3) = m_XMLPrjParams.clsMgmtScenHolder.Item(i).strChangeToClass
                             End If
                         Else
                             Exit Sub
                         End If
                    End If
                End If
             End With
        Next i
    
    ClearMgmtCheckBoxes True, m_intMgmtCount
    CreateMgmtCheckBoxes True
    
    End If
    
    'Reset to first tab
    SSTab1.Tab = 0
    m_booExists = True
    Screen.MousePointer = vbNormal
    
    'Cleanup
    rsCurrWShed.Close
    Set rsCurrWShed = Nothing
    Set pDBasinFeatureLayer = Nothing
    Set pDBasinDataset = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError False, "FillForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
VersionProblem:
    MsgBox "Version Problem"
End Sub

Private Sub cmdOpenWS_Click()
    On Error GoTo ErrorHandler

    Dim initFolder As String
    
    Dim shlShell As Shell32.Shell
    Dim shlFolder As Shell32.Folder3
    Dim fso As FileSystemObject
    
    Set fso = New FileSystemObject
    initFolder = modUtil.g_nspectDocPath & "\workspace"
    If Not fso.FolderExists(initFolder) Then
        MkDir initFolder
    End If
    
    Set shlShell = New Shell32.Shell
    Set shlFolder = shlShell.BrowseForFolder(Me.hwnd, "Choose a directory for analysis output: ", &H1) ', initFolder)
    If Not shlFolder Is Nothing Then
        Me.txtOutputWS.Text = shlFolder.Self.path
        'MsgBox "Output folder: " & Me.txtOutputWS.Text
        m_strWorkspace = txtOutputWS.Text
    End If
    
    'Open workspace button
    'uses the addin folder browser; Reference: vbAccelerator folder browse library
    'dlgBrowser is initiated on Form Load
    'With dlgBrowser
    '    .hwndOwner = Me.hwnd
    '    .InitialDir = initFolder
    '    .FileSystemOnly = True
    '    .StatusText = True
    '    .EditBox = True
    '    .UseNewUI = True
    '    .Title = "Choose a Directory for Analysis Output:"
    'End With
    'txtOutputWS.Text = dlgBrowser.BrowseForFolder
    'm_strWorkspace = txtOutputWS.Text
    
    Exit Sub
  
ErrorHandler:
    HandleError True, "cmdOpenWS_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrorHandler

    'Enables Shift + F1 to bring up NSPECT help.  Regular F1 brings in the darn ArcMap help

    Dim shiftdown As Integer
    
    If (Shift And vbShiftMask) > 0 Then

        If KeyCode = vbKeyF1 Then
            HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "project_setup.htm"
        End If
        
        If KeyCode = Shift + vbKeyF7 Then
            frmAbout.Show vbModal
        End If

    End If

  Exit Sub
ErrorHandler:
  HandleError True, "Form_KeyDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler


    'Check the DPI setting: This was put in place because the alignment of the checkboxes will be
    'messed up if DPI settings are anything other than Normal
    Dim lngMapDC As Long
    Dim lngDPI As Long
    lngMapDC = GetDC(frmPrj.hwnd)
    lngDPI = GetDeviceCaps(lngMapDC, LOGPIXELSX)
    ReleaseDC frmPrj.hwnd, lngMapDC
    
    If lngDPI <> 96 Then
        MsgBox "Warning: N-SPECT requires your font size to be 96 DPI." & vbNewLine & _
        "Some controls may appear out of alignment on this form.", vbCritical, "Warning!"
    End If
    
    m_bolFirstLoad = True           'It's the first load
    m_booExists = False
    
    'Initialize the browse for folder vbAccelerator Reference
    'Set dlgBrowser = New cBrowseForFolder
    
    'define flexgrid for coefficients tab
    With grdCoeffs
      .col = .FixedCols
      .row = .FixedRows
      .Width = 7500
      .ColWidth(0) = 400
      .ColWidth(1) = 600
      .ColWidth(2) = 2200
      .ColWidth(3) = 2400
      .ColWidth(4) = 1800
      .ColWidth(5) = 0
      .ColWidth(6) = 0
      .row = 0
      .col = 1
      .Text = "Apply"
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "Pollutant Name"
      .CellAlignment = flexAlignCenterCenter
      .col = 3
      .Text = "Coefficient Set"
      .CellAlignment = flexAlignCenterCenter
      .col = 4
      .Text = "Which Coefficient"
      .CellAlignment = flexAlignCenterCenter
      
    End With
    
    'define flexgrid for land cover change scenarios on Management Scenarios tab
    With grdLCChanges
      .col = .FixedCols
      .row = .FixedRows
      .Width = 6700
      .ColWidth(0) = 400
      .ColWidth(1) = 600
      .ColWidth(2) = 2800
      .ColWidth(3) = 2800
      .row = 0
      .col = 1
      .Text = "Apply"
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "Change area layer"
      .CellAlignment = flexAlignCenterCenter
      .col = 3
      .Text = "Change to class"
      .CellAlignment = flexAlignCenterCenter
    End With
    
    'define flexgrid for Point Sources tab
    With grdLU
      .col = .FixedCols
      .row = .FixedRows
      .Width = 5200 + (.GridLineWidth * (.Cols + 1)) + 100
      .ColWidth(0) = 400
      .ColWidth(1) = 600
      .ColWidth(2) = 4200
      .ColWidth(3) = 0
      .row = 0
      .col = 1
      .Text = "Apply"
      .CellAlignment = flexAlignCenterCenter
      .col = 2
      .Text = "Land Use Scenario"
      .CellAlignment = flexAlignCenterCenter
    End With
    
    'Fill the Form
    
    'ComboBox::LandCover Type
    modUtil.InitComboBox cboLCType, "LCType"
    
    'ComboBox::Precipitation Scenarios
    modUtil.InitComboBox cboPrecipScen, "PrecipScenario"
    cboPrecipScen.AddItem "New precipitation scenario...", cboPrecipScen.ListCount
    
    'ComboBox::WaterShed Delineations
    modUtil.InitComboBox cboWSDelin, "WSDelineation"
    cboWSDelin.AddItem "New watershed delineation...", cboWSDelin.ListCount
    
    'ComboBox::WaterQuality Criteria
    modUtil.InitComboBox cboWQStd, "WQCriteria"
    cboWQStd.AddItem "New water quality standard...", cboWQStd.ListCount
    
    'Fill Land Cover cbo
    modUtil.AddRasterLayerToComboBox cboLCLayer, m_pMap
    
    'Fill Rain GRID cbo
    modUtil.AddRasterLayerToComboBox cboRainGrid, m_pMap
    
    'Soils, now a 'scenario', not just a datalayer
    modUtil.InitComboBox cboSoilsLayer, "Soils"
    
    'Fill area
    modUtil.AddFeatureLayerToComboBox cboAreaLayer, m_pMap, "poly"
    
    'Fill LandClass
    FillCboLCCLass
    
    m_intMgmtCount = grdLCChanges.Rows - 1   'Number of mgmt scens
    m_intLUCount = grdLU.Rows - 1            'Number of landuses
    
    'Initialize parameter file
    Set m_XMLPrjParams = New clsXMLPrjFile
    
    frmPrj.Caption = "Untitled"
    
    'Find out what the deal is
    cboSelectPoly.Clear
    modUtil.AddFeatureLayerToComboBox cboSelectPoly, m_pMap, "poly"
    
    chkSelectedPolys.Enabled = EnableChkWaterShed
    
    'Test workspace persistence
    If Len(m_strWorkspace) > 0 Then
        txtOutputWS.Text = m_strWorkspace
    End If
    
    
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub cmdOutputBrowse_Click()
On Error GoTo ErrHandler:

   'browse...get output filename
   dlgXML.FileName = Empty
   With dlgXML
     .Filter = MSG6
     .DialogTitle = MSG7
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .CancelError = True
     .ShowSave
   End With
   
   If Len(dlgXML.FileName) > 0 Then
      txtOutputFile.Text = Trim(dlgXML.FileName)
      txtThemeName.Text = modUtil.SplitFileName(txtOutputFile.Text)
   End If
   
ErrHandler:
    Exit Sub

End Sub

Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler

    Dim intvbYesNo As Integer
    
    intvbYesNo = MsgBox("Do you want to save changes you made to " & frmPrj.Caption & "?", vbYesNoCancel + vbExclamation, "N-SPECT")
    
    If intvbYesNo = vbYes Then
        If SaveXMLFile Then
            Unload frmPrj
        End If
    ElseIf intvbYesNo = vbNo Then
        Unload frmPrj
    ElseIf intvbYesNo = vbCancel Then
        Exit Sub
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorHandler

    'Cleanup
    Set m_pMap = Nothing
    Set m_pActiveViewEvents = Nothing
    Set m_pMxDoc = Nothing
    Set m_App = Nothing
    m_strOpenFileName = ""
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Unload " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub grdCoeffs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    m_intPollRow = grdCoeffs.row
    m_intPollCol = grdCoeffs.col
    
    
    If grdCoeffs.TextMatrix(m_intPollRow, 1) = "1" Then
        'We want to limit editing to only the coeff/coeff type columns
        If m_intPollCol = 4 And m_intPollRow >= 1 Then
            
            cboCoeffSet.Visible = False
            
            With cboCoeff
                .Visible = True
                
                .Move SSTab1.Left + grdCoeffs.Left + grdCoeffs.CellLeft, _
                SSTab1.Top + grdCoeffs.Top + grdCoeffs.CellTop
                
                .Width = grdCoeffs.CellWidth
            End With
            
        ElseIf m_intPollCol = 3 And m_intPollRow >= 1 Then
        
            cboCoeff.Visible = False
            
            With cboCoeffSet
                .Visible = True
                .Move SSTab1.Left + grdCoeffs.Left + grdCoeffs.CellLeft, _
                SSTab1.Top + grdCoeffs.Top + grdCoeffs.CellTop
                .Width = grdCoeffs.CellWidth
                FillCboCoeffSet
            End With
        Else
            cboCoeff.Visible = False
            cboCoeffSet.Visible = False
        End If
    Else
        cboCoeff.Visible = False
        cboCoeffSet.Visible = False
    End If
    

End Sub

    
Private Sub grdLCChanges_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

    
    If Button = 2 Then
      PopupMenu mnuManagement
    End If
    


  Exit Sub
ErrorHandler:
  HandleError True, "grdLCChanges_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub grdLCChanges_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m_intLCRow = grdLCChanges.row
    m_intLCCol = grdLCChanges.col
    
    'We
    If m_intLCCol = 3 And m_intLCRow >= 1 Then
        
        cboAreaLayer.Visible = False
        
        With cboClass
            
            .Visible = True
            .Move SSTab1.Left + grdLCChanges.Left + grdLCChanges.CellLeft, _
            SSTab1.Top + grdLCChanges.Top + grdLCChanges.CellTop
            
            .Width = grdLCChanges.CellWidth
        End With
        
    ElseIf m_intLCCol = 2 And m_intLCRow >= 1 Then
        
        cboClass.Visible = False
        
        With cboAreaLayer
            .Visible = True
            .Move SSTab1.Left + grdLCChanges.Left + grdLCChanges.CellLeft, _
            SSTab1.Top + grdLCChanges.Top + grdLCChanges.CellTop
            .Width = grdLCChanges.CellWidth
            
        End With
    
    Else
        
        cboClass.Visible = False
        cboAreaLayer.Visible = False
        
    End If

End Sub

Private Sub m_pActiveViewEvents_ItemAdded(ByVal Item As Variant)
  On Error GoTo ErrorHandler

'Necessary to track items added/removed in this case to the Map, so cbos and such can update them darn selves

    Dim strLCLayer As String
    Dim strAreaLayer As String
    Dim strSelectPolyLayer As String
    Dim strRainLayer As String
    
    'Find out the current LcLayer selection
    strLCLayer = cboLCLayer.Text
        
    'Fill Land Cover cbo
    modUtil.AddRasterLayerToComboBox cboLCLayer, m_pMap
    
    'Return the cboLCLayer to original selection, if there was one
    If Len(strLCLayer) <> 0 Then
        cboLCLayer.ListIndex = modUtil.GetCboIndex(strLCLayer, cboLCLayer)
    End If

    'Fill Rain GRID cbo
    If cboRainGrid.Visible = True Then
        strRainLayer = cboRainGrid.Text
        modUtil.AddRasterLayerToComboBox cboRainGrid, m_pMap
        'Again, check for prior selection, if there was one, return to it
        If Len(strRainLayer) > 0 Then
            cboRainGrid.ListIndex = modUtil.GetCboIndex(strRainLayer, cboRainGrid)
        End If
        
    End If
    
    'Fill area
    strAreaLayer = cboAreaLayer.Text
    modUtil.AddFeatureLayerToComboBox cboAreaLayer, m_pMap, "poly"
    
    If Len(strAreaLayer) <> 0 Then
        cboAreaLayer.ListIndex = modUtil.GetCboIndex(strAreaLayer, cboAreaLayer)
    End If
    
    'Fill SelectPolys
    strSelectPolyLayer = cboSelectPoly.Text
    modUtil.AddFeatureLayerToComboBox cboSelectPoly, m_pMap, "poly"
    cboSelectPoly.Enabled = False
        
    If Len(strSelectPolyLayer) <> 0 Then
        cboSelectPoly.ListIndex = modUtil.GetCboIndex(strSelectPolyLayer, cboSelectPoly)
    End If
        
    chkSelectedPolys.Enabled = EnableChkWaterShed
    
  Exit Sub
ErrorHandler:
  HandleError False, "m_pActiveViewEvents_ItemAdded " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub m_pActiveViewEvents_ItemDeleted(ByVal Item As Variant)
  On Error GoTo ErrorHandler

    Dim pLayer As esriCarto.ILayer
    Dim strLCLayer As String
    Dim strAreaLayer As String
    Dim strSelectPolyLayer As String
    Dim strRainLayer As String

    Set pLayer = Item
    
    'Find out the current LcLayer selection
    strLCLayer = cboLCLayer.Text
        
    'Fill Land Cover cbo
    modUtil.AddRasterLayerToComboBox cboLCLayer, m_pMap
    
    'Return the cboLCLayer to original selection, if there was one, have to make sure
    'however that the item removed isn't the selected item
    If Len(strLCLayer) <> 0 Then
        If Not pLayer.Name = strLCLayer Then
            cboLCLayer.ListIndex = modUtil.GetCboIndex(strLCLayer, cboLCLayer)
        Else
            cboLCLayer.ListIndex = -1
        End If
    End If

    'Fill Rain GRID cbo
    If cboRainGrid.Visible = True Then
        strRainLayer = cboRainGrid.Text
        modUtil.AddRasterLayerToComboBox cboRainGrid, m_pMap
        'Again, check for prior selection, if there was one, return to it
        If Len(strRainLayer) > 0 Then
            If Not pLayer.Name = strRainLayer Then
                cboRainGrid.ListIndex = modUtil.GetCboIndex(strRainLayer, cboRainGrid)
            Else
                cboRainGrid.ListIndex = -1
            End If
        End If
        
    End If
    
    'Fill area
    strAreaLayer = cboAreaLayer.Text
    modUtil.AddFeatureLayerToComboBox cboAreaLayer, m_pMap, "poly"
    
    If Len(strAreaLayer) <> 0 Then
        If Not pLayer.Name = strAreaLayer Then
            cboAreaLayer.ListIndex = modUtil.GetCboIndex(strAreaLayer, cboAreaLayer)
        Else
            cboAreaLayer.ListIndex = -1
        End If
    End If
    
    'Fill SelectPolys
    strSelectPolyLayer = cboSelectPoly.Text
    modUtil.AddFeatureLayerToComboBox cboSelectPoly, m_pMap, "poly"
        
    If Len(strSelectPolyLayer) <> 0 Then
        If Not pLayer.Name = strSelectPolyLayer Then
            cboSelectPoly.ListIndex = modUtil.GetCboIndex(strSelectPolyLayer, cboSelectPoly)
        Else
            cboSelectPoly.ListIndex = -1
        End If
    End If
        
    chkSelectedPolys.Enabled = EnableChkWaterShed
    cboSelectPoly.Enabled = chkSelectedPolys.Enabled
    
    'Clean
    Set pLayer = Nothing
    


  Exit Sub
ErrorHandler:
  HandleError False, "m_pActiveViewEvents_ItemDeleted " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub m_pActiveViewEvents_SelectionChanged()
  On Error GoTo ErrorHandler

    chkSelectedPolys.Enabled = EnableChkWaterShed
    
  Exit Sub
ErrorHandler:
  HandleError False, "m_pActiveViewEvents_SelectionChanged " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub mnuExit_Click()
  On Error GoTo ErrorHandler

    Dim intvbYesNo As Integer
    
    intvbYesNo = MsgBox("Do you want to save changes you made to " & frmPrj.Caption & "?", vbYesNoCancel + vbExclamation, "N-SPECT")
    
    If intvbYesNo = vbYes Then
        If SaveXMLFile Then
            Unload frmPrj
        End If
    ElseIf intvbYesNo = vbNo Then
        Unload frmPrj
    ElseIf intvbYesNo = vbCancel Then
        Exit Sub
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "mnuExit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub mnuGeneralHelp_Click()
  On Error GoTo ErrorHandler

    'API call to help
    HtmlHelp 0, modUtil.g_nspectPath & "\Help\nspect.chm", HH_DISPLAY_TOPIC, ByVal "project_setup.htm"
    
        
  Exit Sub
ErrorHandler:
  HandleError True, "mnuGeneralHelp_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub mnuLUDelete_Click()
  On Error GoTo ErrorHandler

    'delete current row
    Dim R%, C%, row%, col%
    
    With grdLU
      If .Rows > .FixedRows Then        'make sure we don't del header Rows
          For col% = 1 To .Cols - 1
              If ((Trim(.TextMatrix(.row, col%)) > "" And col% = 2) Or _
                 (.TextMatrix(.row, col%) <> "0" And col% = 1) Or _
                 (.TextMatrix(.row, col%) <> "0" And col% >= 3)) Then 'data?
                  C% = 1
                  Exit For
              End If
          Next col%
          If C% Then
              R% = MsgBox("There is data in Row" + Str$(.row) + "! Delete anyway?", vbYesNo, "Delete Row!")
          End If
          If C% = 0 Or R% = vbYes Then        'no exist. data or YES
              If .row = .Rows - 1 And .Rows = 2 Then  'last row?
                  'If .row = 1 And .Rows = 1 Then  'you want to leave 1 row, but empty it.
                    .TextMatrix(.row, 1) = 0
                    .TextMatrix(.row, 2) = ""
                    .TextMatrix(.row, 3) = ""
              Else
                    For row% = .row To .Rows - 2 'move data up 1 row
                      For col% = 1 To .Cols - 1
                          .TextMatrix(row%, col%) = .TextMatrix(row% + 1, col%)
                      Next col%
                    Next row%
                  .Rows = .Rows - 1     'del last row
                  End If
              
              End If
            
          End If
      
   End With
    
    m_intLUCount = grdLU.Rows
   
    m_intCount = grdLU.Rows
    ClearLUCheckBoxes True, m_intCount + 1
    CreateLUCheckBoxes True

  Exit Sub
ErrorHandler:
  HandleError True, "mnuLUDelete_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

'Add row...
Private Sub mnuManAppen_Click()
  On Error GoTo ErrorHandler

    'add row to end of  grid
    With grdLCChanges
       .Rows = .Rows + 1
       .row = .Rows - 1
       .TextMatrix(.row, 1) = "0"
    End With
    
    m_intMgmtCount = grdLCChanges.Rows
    
    CreateMgmtCheckBoxes False, grdLCChanges.row
    
  Exit Sub
ErrorHandler:
  HandleError True, "mnuManAppen_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub mnuManDelete_Click()
  On Error GoTo ErrorHandler

'delete current row

    Dim R%, C%, row%, col%
    
    With grdLCChanges
    
      If .Rows > .FixedRows Then        'make sure we don't del header Rows
          For col% = 1 To .Cols - 1
              If ((Trim(.TextMatrix(.row, col%)) > "" And col% = 2) Or _
                 (.TextMatrix(.row, col%) <> "0" And col% = 1) Or _
                 (.TextMatrix(.row, col%) <> "0" And col% >= 3)) Then 'data?
                  C% = 1
                  Exit For
              End If
          Next col%
          If C% Then
              R% = MsgBox("There is data in Row" + Str$(.row) + " ! Delete anyway?", vbYesNo, "Delete Row!")
          End If
          If C% = 0 Or R% = vbYes Then        'no exist. data or YES
              If .row = .Rows - 1 Then  'last row?
                  .row = .row - 1       'move active cell
                 
              Else
                  For row% = .row To .Rows - 2 'move data up 1 row
                      For col% = 1 To .Cols - 1
                          .TextMatrix(row%, col%) = .TextMatrix(row% + 1, col%)
                      Next col%
                  Next row%
              End If
               .Rows = .Rows - 1 'del last row
          Else
            Exit Sub
          End If
      End If
   End With
   
   m_intMgmtCount = grdLCChanges.Rows
   
   ClearMgmtCheckBoxes True, m_intMgmtCount
   CreateMgmtCheckBoxes True
   
  Exit Sub
ErrorHandler:
  HandleError True, "mnuManDelete_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub mnuManInsert_Click()
  On Error GoTo ErrorHandler

'Insert row above current row in grdLChanges- Thanks, Andrew

    Dim R%, row%, col%
    With grdLCChanges
      If .row < .FixedRows Then        'make sure we don't insert above header Rows
         mnuManAppen_Click
      Else
         R% = .row
         .Rows = .Rows + 1             'add a row
         
         For row% = .Rows - 1 To R% + 1 Step -1 'move data dn 1 row
             For col% = 1 To .Cols - 1
                 .TextMatrix(row%, col%) = .TextMatrix(row% - 1, col%)
             Next col%
         Next row%
         For col% = 1 To .Cols - 1       ' clear all cells in this row
            If (col% = 1) Then
               .TextMatrix(R%, col%) = "0"
            Else
               .TextMatrix(R%, col%) = ""
            End If
         Next col%
         
         
     End If
   End With
    
  
   m_intMgmtCount = grdLCChanges.Rows
   
   ClearMgmtCheckBoxes True, m_intMgmtCount - 2
   CreateMgmtCheckBoxes True
     

  Exit Sub
ErrorHandler:
  HandleError True, "mnuManInsert_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub MnuLUEdit_Click()
  On Error GoTo ErrorHandler

   mnuOptions = False
   
   g_intManScenRow = m_intLURow
   g_strLUScenFileName = grdLU.TextMatrix(m_intLURow, 3)
   
   With frmLUScen
    .init m_App, cboWQStd.Text
    .Caption = "Edit Land Use Scenario"
    .Show vbModal, Me
   End With
   
  Exit Sub
ErrorHandler:
  HandleError False, "MnuLUEdit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub MnuLUAdd_Click()
  On Error GoTo ErrorHandler

   Dim intRow As Integer
   
   g_intManScenRow = m_intLURow
   g_strLUScenFileName = grdLU.TextMatrix(m_intLURow, 3)
   
   'if there is data in the row....
   If g_strLUScenFileName <> "" Then

        grdLU.Rows = grdLU.Rows + 1 'Add a row
        g_strLUScenFileName = ""   'reset data
        g_intManScenRow = grdLU.Rows - 1 '
        intRow = g_intManScenRow
        CreateLUCheckBoxes False, intRow
   Else
        g_intManScenRow = m_intLURow
        g_strLUScenFileName = ""
   End If
   
   With frmLUScen
    .init m_App, cboWQStd.Text
    .Caption = "Add Land Use Scenario"
    .Show vbModal, Me
   End With


  Exit Sub
ErrorHandler:
  HandleError False, "MnuLUAdd_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub grdLU_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo ErrorHandler

    Dim strName As String
    
    m_intLURow = grdLU.MouseRow
    m_intLUCol = grdLU.MouseCol
    
    'Set the popup to proper functionality, if current row, all for add, delete, edit
    'If empty row, disable edit, delete
    If Button = 2 And m_intLURow > 0 Then
        strName = grdLU.TextMatrix(m_intLURow, 2)
        If Len(strName) = 0 Then
            mnuLUEdit.Enabled = False
            mnuLUDelete.Enabled = False
            PopupMenu mnuOptions
        Else
            mnuLUEdit.Enabled = True
            mnuLUDelete.Enabled = True
            PopupMenu mnuOptions
        End If
    End If
    
  Exit Sub
ErrorHandler:
  HandleError True, "grdLU_MouseDown " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub mnuNew_Click()
  On Error GoTo ErrorHandler

    Dim intvbYesNo As Integer
    
    intvbYesNo = MsgBox("Do you want to save changes you made to " & frmPrj.Caption & "?", vbYesNoCancel + vbExclamation, "N-SPECT")
    
    If intvbYesNo = vbYes Then
        If SaveXMLFile Then
            ClearForm
            Form_Load
        Else
            Exit Sub
        End If
    ElseIf intvbYesNo = vbNo Then
        ClearForm
        Form_Load
    ElseIf intvbYesNo = vbCancel Then
        Exit Sub
    End If
    
  Exit Sub
ErrorHandler:
  HandleError True, "mnuNew_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub ClearForm()
  On Error GoTo ErrorHandler

'Gotta clean up before new, clean form

    ClearCheckBoxes True
    ClearMgmtCheckBoxes True, m_intMgmtCount
    ClearLUCheckBoxes True
    
    'LandClass stuff
    cboLCLayer.Clear
    cboLCType.Clear
    
    'DBase scens
    cboPrecipScen.Clear
    cboWSDelin.Clear
    cboWQStd.Clear
    cboSoilsLayer.Clear
    
    'Text
    txtProjectName.Text = ""
    txtOutputWS.Text = ""
    
    'Checkboxes
    chkSelectedPolys.Value = 0
    chkLocalEffects.Value = 0
    
    'Erosion
    cboRainGrid.Clear
    optUseGRID.Value = True
    chkCalcErosion.Value = 0
    txtRainValue = ""
    
    'txtOutputFile.Text = ""
    txtThemeName = ""
    
    'clear the GRIDS
    grdLU.Clear
    grdLU.Rows = 2
    grdLCChanges.Clear
    grdLCChanges.Rows = 2
    grdCoeffs.Clear
    


  Exit Sub
ErrorHandler:
  HandleError False, "ClearForm " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub


Private Sub mnuOpen_Click()
  'On Error GoTo ErrorHandler
   
    LoadXMLFile
    
  'Exit Sub
'ErrorHandler:
  'HandleError True, "mnuOpen_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub mnuSave_Click()
  On Error GoTo ErrorHandler

    SaveXMLFile
    
  Exit Sub
ErrorHandler:
  HandleError True, "mnuSave_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub mnuSaveAs_Click()
  On Error GoTo ErrorHandler
    
    m_booExists = False
    SaveXMLFile
    
  Exit Sub
ErrorHandler:
  HandleError True, "mnuSaveAs_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub optUseGRID_Click()
  On Error GoTo ErrorHandler

    cboRainGrid.Enabled = optUseGRID.Value
    txtRainValue.Enabled = optUseValue.Value
    
  Exit Sub
ErrorHandler:
  HandleError True, "optUseGRID_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub optUseValue_Click()
  On Error GoTo ErrorHandler

    txtRainValue.Enabled = optUseValue.Value
    cboRainGrid.Enabled = optUseGRID.Value

  Exit Sub
ErrorHandler:
  HandleError True, "optUseValue_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

'*********************************************************************************************************
'FORM SPECIFIC FUNCTIONS/SUBS
'*********************************************************************************************************
Private Sub PopPollutants()
  On Error GoTo ErrorHandler


    Dim strSQLWQStd As String
    Dim rsWQStdCboClick As New ADODB.Recordset
    
    Dim strSQLWQStdPoll As String
    Dim rsWQStdPoll As New ADODB.Recordset
    
    Dim i As Integer
    
    'Selection based on combo box
    strSQLWQStd = "SELECT * FROM WQCRITERIA WHERE NAME LIKE '" & cboWQStd.Text & "'"
    rsWQStdCboClick.CursorLocation = adUseClient
    rsWQStdCboClick.Open strSQLWQStd, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
    If rsWQStdCboClick.RecordCount > 0 Then
        
        strSQLWQStdPoll = "SELECT POLLUTANT.NAME, POLL_WQCRITERIA.THRESHOLD " & _
        "FROM POLL_WQCRITERIA INNER JOIN POLLUTANT " & _
        "ON POLL_WQCRITERIA.POLLID = POLLUTANT.POLLID Where POLL_WQCRITERIA.WQCRITID = " & rsWQStdCboClick!WQCRITID
             
        rsWQStdPoll.CursorLocation = adUseClient
        rsWQStdPoll.Open strSQLWQStdPoll, modUtil.g_ADOConn, adOpenDynamic, adLockOptimistic
    
        grdCoeffs.Rows = rsWQStdPoll.RecordCount + 1
        
        For i = 1 To rsWQStdPoll.RecordCount
            
            grdCoeffs.TextMatrix(i, 2) = rsWQStdPoll!Name
            grdCoeffs.TextMatrix(i, 5) = rsWQStdPoll!Threshold
            rsWQStdPoll.MoveNext
            
        Next i
        
        m_intCount = rsWQStdPoll.RecordCount
               
        'Clean it
        rsWQStdPoll.Close
        Set rsWQStdPoll = Nothing
        
        rsWQStdCboClick.Close
        Set rsWQStdCboClick = Nothing
        
    Else
        
        MsgBox "Warning: There are no water quality standards remaining.  Please add a new one.", vbCritical, "Recordset Empty"
        Set rsWQStdCboClick = Nothing
    End If




  Exit Sub
ErrorHandler:
  HandleError False, "PopPollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub FillCboCoeffSet()
  On Error GoTo ErrorHandler

    Dim intCol As Integer
    Dim strPollName As String
    Dim strSelectCoeff As String
    Dim rsCoeffSet As New ADODB.Recordset
    Dim i As Integer
    
    cboCoeffSet.Clear
    
    intCol = m_intPollCol - 1
    strPollName = grdCoeffs.TextMatrix(m_intPollRow, intCol)
    
    strSelectCoeff = "SELECT POLLUTANT.POLLID, POLLUTANT.NAME, COEFFICIENTSET.NAME AS NAME2 FROM POLLUTANT INNER JOIN COEFFICIENTSET " & _
        "ON POLLUTANT.POLLID = COEFFICIENTSET.POLLID Where POLLUTANT.NAME LIKE '" & strPollName & "'"
    
    rsCoeffSet.Open strSelectCoeff, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    For i = 0 To rsCoeffSet.RecordCount - 1
        cboCoeffSet.AddItem rsCoeffSet!Name2, i
        rsCoeffSet.MoveNext
    Next i
    
    'cleanup
    rsCoeffSet.Close
    Set rsCoeffSet = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError False, "FillCboCoeffSet " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub FillCboLCCLass()
  On Error GoTo ErrorHandler


    Dim strLCChanges As String
    Dim rsLCChanges As New ADODB.Recordset
    Dim i As Integer
    
    strLCChanges = "SELECT LCCLASS.Name as Name2, LCTYPE.LCTYPEID FROM LCTYPE INNER JOIN LCCLASS " & _
                "ON LCTYPE.LCTYPEID = LCCLASS.LCTYPEID WHERE LCTYPE.NAME LIKE '" & cboLCType.Text & "'"

    rsLCChanges.Open strLCChanges, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    rsLCChanges.MoveFirst
    
    cboClass.Clear
    
    For i = 0 To rsLCChanges.RecordCount - 1
        cboClass.AddItem rsLCChanges!Name2, i
        rsLCChanges.MoveNext
    Next i
    
    rsLCChanges.Close
    Set rsLCChanges = Nothing

  Exit Sub
ErrorHandler:
  HandleError False, "FillCboLCCLass " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub CreateCheckBoxes(booAll As Boolean, booUser As Boolean, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler

'booAll:  making all new ones
'booUser:  flag set to determine if the boxes are being made during loading of a project file
'option intRecNo: creation of 1 box

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strChkName As String
    j = 1
    i = 1
    
    SSTab1.Tab = 0
    
    If booAll Then
        For i = 1 To grdCoeffs.Rows - 1
            
            grdCoeffs.row = i
            grdCoeffs.col = 1
            
            'Set the alignment to center
            grdCoeffs.CellAlignment = flexAlignCenterCenter
            
            k = i
            
            Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
            
            Load chkIgnore(k)
            
            Set chkIgnore(k).Container = frmPrj
            With chkIgnore(k)
                .Visible = True
                .Top = SSTab1.Top + grdCoeffs.Top + grdCoeffs.CellTop
                .Left = (grdCoeffs.Left) + (grdCoeffs.CellLeft) + (SSTab1.Left) + (grdCoeffs.CellWidth * 0.4)
                .Height = 195
                .Width = 195
                
                If booUser Then 'if during load event of project file...
                    'If Threshold > 0 or The User has choosen to ignore then...
                    If (CInt(grdCoeffs.TextMatrix(i, 5)) <> 0) And grdCoeffs.TextMatrix(i, 1) = 1 Then
                        .Value = 1
                        '.Enabled = False
                    Else
                        .Value = 0
                        '.Enabled = False
                    End If
                Else
                    If (CInt(grdCoeffs.TextMatrix(i, 5) > 0)) Then
                        .Value = 0
                        .Enabled = True
                    Else
                        .Value = 0
                        .Enabled = False
                    End If
                End If
            End With
           
            grdCoeffs.TextMatrix(i, 1) = CStr(chkIgnore(i).Value)
            Call Controls.Remove("chk" & CStr(k))
        Next i
    End If



  Exit Sub
ErrorHandler:
  HandleError False, "CreateCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub ClearCheckBoxes(booAll As Boolean, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler

    Dim k As Integer
    
    SSTab1.Tab = 0
    
    If booAll Then
    
        For k = 1 To chkIgnore().UBound
            Unload chkIgnore(k)
        Next k
    Else
        Unload chkIgnore(intRecNo)
    End If
    


  Exit Sub
ErrorHandler:
  HandleError False, "ClearCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub CreateMgmtCheckBoxes(booAll As Boolean, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler


    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strChkName As String
    j = 1
    i = 1
    
    SSTab1.Tab = 3
    
    If booAll Then
        For i = 1 To grdLCChanges.Rows - 1
            
            grdLCChanges.row = i
            grdLCChanges.col = 1
            
            'Set the alignment to center
            grdLCChanges.CellAlignment = flexAlignCenterCenter
            
            k = i
            
            Call Controls.Add("VB.CheckBox", "chkmgmt" & CStr(k), Me)
            
            Load chkIgnoreMgmt(k)
            Set chkIgnoreMgmt(k).Container = frmPrj
            
            With chkIgnoreMgmt(k)
                .Visible = True
                .Top = SSTab1.Top + grdLCChanges.Top + grdLCChanges.CellTop
                .Left = grdLCChanges.Left + (grdLCChanges.CellLeft) + (SSTab1.Left) + (grdLCChanges.CellWidth * 0.4)
                .Height = 195
                .Width = 195
                
                If grdLCChanges.TextMatrix(k, 1) <> "" Then
                    .Value = grdLCChanges.TextMatrix(k, 1)
                Else
                    .Value = 0
                End If
                grdLCChanges.TextMatrix(k, 1) = chkIgnoreMgmt(k).Value
            End With
           Call Controls.Remove("chkmgmt" & CStr(k))
        Next i
        Else
            With grdLCChanges
                .row = intRecNo
                .col = 1
                .CellAlignment = flexAlignCenterCenter
            End With
            
            k = intRecNo
            
            Call Controls.Add("VB.CheckBox", "chkmgmt" & CStr(k), Me)
            
            Load chkIgnoreMgmt(k)
            Set chkIgnoreMgmt(k).Container = frmPrj
            
            With chkIgnoreMgmt(k)
                .Visible = True
                .Top = SSTab1.Top + grdLCChanges.Top + grdLCChanges.CellTop
                .Left = SSTab1.Left + grdLCChanges.Left + grdLCChanges.CellLeft + (grdLCChanges.CellWidth * 0.4)
                .Height = 195
                .Width = 195
                .Value = CInt(grdLCChanges.TextMatrix(k, 1))
            End With
           Call Controls.Remove("chkmgmt" & CStr(k))
    End If
    
    grdLCChanges.TextMatrix(0, 0) = ""



  Exit Sub
ErrorHandler:
  HandleError False, "CreateMgmtCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub ClearMgmtCheckBoxes(booAll As Boolean, intCount As Integer, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler


    Dim k As Integer
    
    SSTab1.Tab = 3
    
    If booAll Then
        
        For k = 1 To chkIgnoreMgmt().UBound
            Unload chkIgnoreMgmt(k)
        Next k
    Else
        Unload chkIgnoreMgmt(intRecNo)
    End If
    


  Exit Sub
ErrorHandler:
  HandleError False, "ClearMgmtCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub CreateLUCheckBoxes(booAll As Boolean, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler


    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim strChkName As String
    j = 1
    i = 1
    
    SSTab1.Tab = 2
    
    If booAll Then
        For i = 1 To grdLU.Rows - 1
            
            grdLU.row = i
            grdLU.col = 1
            
            'Set the alignment to center
            grdLU.CellAlignment = flexAlignCenterCenter
            
            k = i
            
            Call Controls.Add("VB.CheckBox", "chklu" & CStr(k), Me)
            
            Load chkIgnoreLU(k)
            
            Set chkIgnoreLU(k).Container = frmPrj
            With chkIgnoreLU(k)
                .Visible = True
                .Top = SSTab1.Top + grdLU.Top + grdLU.CellTop
                .Left = grdLU.Left + (grdLU.CellLeft) + (SSTab1.Left) + (grdLU.CellWidth * 0.4)
                .Height = 195
                .Width = 195
                If grdLU.TextMatrix(i, 1) <> "" Then
                    .Value = grdLU.TextMatrix(i, 1)
                Else
                    .Value = 0
                End If
            End With
            grdLU.TextMatrix(k, 1) = chkIgnoreLU(k).Value
            Call Controls.Remove("chklu" & CStr(k))
        Next i
    Else
        With grdLU
                .row = intRecNo
                .col = 1
                .CellAlignment = flexAlignCenterCenter
            End With
            
            k = intRecNo
            
            Call Controls.Add("VB.CheckBox", "chk" & CStr(k), Me)
            Load chkIgnoreLU(k)
            Set chkIgnoreLU(k).Container = frmPrj
            With chkIgnoreLU(k)
                .Visible = True
                .Top = SSTab1.Top + grdLU.Top + grdLU.CellTop
                .Left = grdLU.Left + (grdLU.CellLeft) + (SSTab1.Left) + (grdLU.CellWidth * 0.4)
                .Height = 195
                .Width = 195
                .Value = 0
        End With
        grdLU.TextMatrix(k, 1) = chkIgnoreLU(k).Value
        Call Controls.Remove("chk" & CStr(k))
    End If
    



  Exit Sub
ErrorHandler:
  HandleError False, "CreateLUCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub
Private Sub ClearLUCheckBoxes(booAll As Boolean, Optional intRecNo As Integer)
  On Error GoTo ErrorHandler


    Dim k As Integer
    
    SSTab1.Tab = 2
    
    If booAll Then
        For k = 1 To chkIgnoreLU().UBound
            Unload chkIgnoreLU(k)
        Next k
    Else
        Unload chkIgnoreLU(intRecNo)
    End If
    


  Exit Sub
ErrorHandler:
  HandleError False, "ClearLUCheckBoxes " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
  On Error GoTo ErrorHandler

    
    Dim i As Integer
    
    Select Case SSTab1.Tab
    
        Case 0
            For i = 0 To chkIgnore.UBound
                chkIgnore(i).Visible = True
            Next i
            
            For i = 0 To chkIgnoreMgmt.UBound
                chkIgnoreMgmt(i).Visible = False
            Next i
            
            For i = 0 To chkIgnoreLU.UBound
                chkIgnoreLU(i).Visible = False
            Next i
            
            cboAreaLayer.Visible = False
            cboClass.Visible = False
            
        Case 1
            For i = 0 To chkIgnore.UBound
                chkIgnore(i).Visible = False
            Next i
            
            For i = 0 To chkIgnoreMgmt.UBound
                chkIgnoreMgmt(i).Visible = False
            Next i
                        
            For i = 0 To chkIgnoreLU.UBound
                chkIgnoreLU(i).Visible = False
            Next i
            
            cboCoeff.Visible = False
            cboCoeffSet.Visible = False
            cboAreaLayer.Visible = False
            cboClass.Visible = False
        
        Case 2
            For i = 0 To chkIgnore.UBound
                chkIgnore(i).Visible = False
            Next i
            
            For i = 0 To chkIgnoreMgmt.UBound
                chkIgnoreMgmt(i).Visible = False
            Next i
            
            For i = 0 To chkIgnoreLU.UBound
                chkIgnoreLU(i).Visible = True
            Next i
            
            cboCoeff.Visible = False
            cboCoeffSet.Visible = False
            cboAreaLayer.Visible = False
            cboClass.Visible = False
            
        Case 3
            For i = 0 To chkIgnore.UBound
                chkIgnore(i).Visible = False
            Next i
            
            For i = 0 To chkIgnoreMgmt.UBound
                chkIgnoreMgmt(i).Visible = True
            Next i
            
            For i = 0 To chkIgnoreLU.UBound
                chkIgnoreLU(i).Visible = False
            Next i
            
            cboCoeff.Visible = False
            cboCoeffSet.Visible = False
            cboAreaLayer.Visible = False
            cboClass.Visible = False
            
    End Select
    


  Exit Sub
ErrorHandler:
  HandleError True, "SSTab1_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Sub Timer1_Timer()
  On Error GoTo ErrorHandler

'Use of the timer is necessary to create the checkboxes for all Grids that require them

    Timer1.Enabled = False
    
    If m_bolFirstLoad Then
        
        CreateCheckBoxes True, False
        CreateMgmtCheckBoxes True
        CreateLUCheckBoxes True
        
        m_bolFirstLoad = False
    
    ElseIf m_booNew Then
       
        ClearCheckBoxes True
        CreateCheckBoxes True, False
        
        ClearMgmtCheckBoxes True, m_intMgmtCount
        CreateMgmtCheckBoxes True
        ClearLUCheckBoxes True, m_intLUCount
        CreateLUCheckBoxes True
                
               
    End If
    
    SSTab1.Tab = 0
    m_booNew = False



  Exit Sub
ErrorHandler:
  HandleError True, "Timer1_Timer " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Function ValidateData() As Boolean
  On Error GoTo ErrorHandler

'Time to rifle through the form ensuring kosher data across the board.
    
    Dim i As Integer
    Dim j As Integer
    Dim strUpdatePrecip As String
    Dim strMgmnt As String
    Dim pLayer As ILayer
    Dim fso As New FileSystemObject
    
    Dim clsParamsPrj As clsXMLPrjFile 'Just a holder for the xml
    Set clsParamsPrj = New clsXMLPrjFile
    
    'First check Selected Watersheds
    If chkSelectedPolys.Enabled = True And chkSelectedPolys.Value = 1 Then
        If Len(cboSelectPoly.Text) > 0 Then
            Set pLayer = m_pMap.Layer(modUtil.GetLayerIndex(cboSelectPoly.Text, m_App))
            If modUtil.GetSelectedFeatureCount(pLayer, m_pMap) > 0 Then
                ValidateData = True
            End If
        Else
            MsgBox "You have chosen 'Selected watersheds only'.  Please select watersheds.", vbCritical, "No Selected Features Found"
            ValidateData = False
            Exit Function
        End If
    End If
    
    'Project Name
    If Len(txtProjectName.Text) > 0 Then
        clsParamsPrj.strProjectName = Trim(txtProjectName.Text)
    Else
        MsgBox "Please enter a name for this project.", vbInformation, "Enter Name"
        txtProjectName.SetFocus
        ValidateData = False
        Exit Function
    End If
    
    'Working Directory
    If (Len(txtOutputWS.Text) > 0) And fso.FolderExists(txtOutputWS.Text) Then
        clsParamsPrj.strProjectWorkspace = Trim(txtOutputWS.Text)
    Else
        MsgBox "Please choose a valid output working directory.", vbInformation, "Choose Workspace"
        txtOutputWS.SetFocus
        ValidateData = False
        Exit Function
    End If
        
    'LandCover
    If cboLCLayer.Text = "" Then
        MsgBox "Please select a Land Cover layer before continuing.", vbInformation, "Select Land Cover Layer"
        cboLCLayer.SetFocus
        ValidateData = False
        Exit Function
    Else
        If modUtil.LayerInMap(cboLCLayer.Text, m_pMap) Then
            clsParamsPrj.strLCGridName = cboLCLayer.Text
            clsParamsPrj.strLCGridFileName = modUtil.GetRasterFileName(cboLCLayer.Text, m_App)
            clsParamsPrj.strLCGridUnits = cboLCUnits.ListIndex
        Else
            MsgBox "The Land Cover layer you have choosen is not in the current map frame.", vbInformation, "Layer Not Found"
            ValidateData = False
            Exit Function
        End If
    End If
    
    'LC Type
    If cboLCType.Text = "" Then
        MsgBox "Please select a Land Class Type before continuing.", vbInformation, "Select Land Class Type"
        cboLCType.SetFocus
        ValidateData = False
        Exit Function
    Else
        clsParamsPrj.strLCGridType = cboLCType.Text
    End If
    
    'REMOVED 11/30/2007 based on the whole issue of landcover records for clipped images.
    'Now Check LandCover, its table and whether or not the # of records matches in the databaset
'    If Not (CheckLandCoverFields(clsParamsPrj.strLCGridType, modUtil.ReturnRaster(clsParamsPrj.strLCGridFileName))) Then
'        MsgBox "The number of land cover classes in your " & clsParamsPrj.strLCGridName & _
'               " GRID do not match the number entered " & vbNewLine & _
'                " in the " & clsParamsPrj.strLCGridType & " land cover type.  " & _
'                "Please refer to 'Land Cover Types' in the Advanced Settings before proceeding.", _
'                vbInformation, "Records Not Compatible"
'        ValidateData = False
'        Exit Function
'    End If
    
    'Soils - use definition to find datasets, if there use, if not tell the user
    If cboSoilsLayer.Text = "" Then
        MsgBox "Please select a Soils definition before continuing.", vbInformation, "Select Soils Layer"
        cboSoilsLayer.SetFocus
        ValidateData = False
        Exit Function
    Else
        If modUtil.RasterExists(lblSoilsHyd.Caption) Then
            clsParamsPrj.strSoilsDefName = cboSoilsLayer.Text
            clsParamsPrj.strSoilsHydFileName = lblSoilsHyd.Caption
        Else
            MsgBox "The hydrologic soils layer " & lblSoilsHyd.Caption & " you have selected is missing.  Please check you soils definition.", vbInformation, "Soils Layer Not Found"
            ValidateData = False
            Exit Function
        End If
    End If
    
    'PrecipScenario
    'If the layer is in the map, get out, all is well- m_strPrecipFile is established on the
    'PrecipCbo Click event
    If modUtil.LayerInMap(modUtil.SplitFileName(m_strPrecipFile), m_pMap) Then
        clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
    Else
        'Check if you can add it, if so, all is well
        If modUtil.RasterExists(m_strPrecipFile) Then
            clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
        Else
            'Can't find it...well, then send user to Browse
            MsgBox "Unable to find precip dataset: " & m_strPrecipFile & ".  Please Correct", vbInformation, "Cannot Find Dataset"
            m_strPrecipFile = modUtil.BrowseForFileName("Raster", frmPrj, "Browse for Precipitation Dataset...")
            'If new one found, then we must update DataBase
            If Len(m_strPrecipFile) > 0 Then
                strUpdatePrecip = "UPDATE PrecipScenario SET precipScenario.PrecipFileName = '" & m_strPrecipFile & "'" & _
                    "WHERE NAME = '" & cboPrecipScen.Text & "'"
                g_ADOConn.Execute strUpdatePrecip
                'Now we can set the xmlParams
                clsParamsPrj.strPrecipScenario = cboPrecipScen.Text
                'modUtil.AddRasterLayerToMapFromFileName m_strPrecipFile, m_pMap
            Else
                MsgBox "Invalid File.", vbInformation, "Invalid File"
                cboPrecipScen.SetFocus
                ValidateData = False
            End If
        End If
    End If
    
    'Go out to a separate function for this one...WaterShed
    If ValidateWaterShed Then
        clsParamsPrj.strWaterShedDelin = cboWSDelin.Text
    Else
        MsgBox "There is a problems with the selected Watershed Delineation.", vbInformation, "Watershed Delineation"
        ValidateData = False
        Exit Function
    End If
    
    'Water Quality
    If Len(cboWQStd.Text) > 0 Then
        clsParamsPrj.strWaterQuality = cboWQStd.Text
    Else
        MsgBox "Please select a water quality standard.", vbInformation, "Water Quality Standard Missing"
        ValidateData = False
        Exit Function
    End If
    
    'Checkboxes, straight up values
    clsParamsPrj.intLocalEffects = chkLocalEffects.Value
    
    'Theoreretically, user could open file that had selected sheds.
    If chkSelectedPolys.Enabled = True Then
        clsParamsPrj.intSelectedPolys = chkSelectedPolys.Value
        clsParamsPrj.strSelectedPolyFileName = modUtil.GetFeatureFileName(cboSelectPoly.Text, m_App)
        clsParamsPrj.strSelectedPolyLyrName = cboSelectPoly.Text
    Else
        clsParamsPrj.intSelectedPolys = 0
    End If
        
    'Erosion Tab
    'Calc Erosion checkbox
    clsParamsPrj.intCalcErosion = chkCalcErosion.Value
    
    If chkCalcErosion Then
        If modUtil.RasterExists(lblKFactor.Caption) Then
            clsParamsPrj.strSoilsKFileName = lblKFactor.Caption
        Else
            MsgBox "The K Factor soils dataset " & lblSoilsHyd.Caption & " you have selected is missing.  Please check your soils definition.", vbInformation, "Soils K Factor Not Found"
            ValidateData = False
            Exit Function
        End If
        
        'Check the Rainfall Factor grid objects.
        If frameRainFall.Visible = True Then
            
                If optUseGRID.Value Then
                
                    If Len(cboRainGrid.Text) > 0 And (InStr(1, cboRainGrid.Text, cboLCLayer.Text, 1) = 0) Then
                        clsParamsPrj.intRainGridBool = 1
                        clsParamsPrj.intRainConstBool = 0
                        clsParamsPrj.strRainGridName = cboRainGrid.Text
                        clsParamsPrj.strRainGridFileName = modUtil.GetRasterFileName(cboRainGrid.Text, m_App)
                    Else
                        MsgBox "Please choose a rainfall Grid.", vbInformation, "Select Rainfall GRID"
                        SSTab1.Tab = 1
                        ValidateData = False
                        Exit Function
                        
                    End If
                    
                ElseIf optUseValue.Value Then
                    
                    If Not IsNumeric(txtRainValue.Text) Then
                        MsgBox "Numbers Only for Rain Values.", vbInformation, "Numbers Only Please"
                        txtRainValue.SetFocus
                    Else
                        If CDbl(txtRainValue.Text) < 0 Then
                            MsgBox "Positive values only please for rainfall values.", vbInformation, "Postive Values Only"
                            txtRainValue.SetFocus
                        Else
                            clsParamsPrj.intRainConstBool = 1
                            clsParamsPrj.dblRainConstValue = CDbl(txtRainValue.Text)
                            clsParamsPrj.strRainGridFileName = ""
                        End If
                    End If
                
                Else
                    MsgBox "You must choose a rainfall factor.", vbInformation, "Rainfall Factor Missing"
                    ValidateData = False
                    Exit Function
                End If
        End If
        
        'Soil Delivery Ratio
        'Added 12/03/07 to account for soil delivery ratio GRID, user can now provide.
        If chkSDR.Value = 1 Then
            If Len(txtSDRGRID.Text) > 0 Then
                If modUtil.RasterExists(txtSDRGRID.Text) Then
                    clsParamsPrj.intUseOwnSDR = 1
                    clsParamsPrj.strSDRGridFileName = txtSDRGRID.Text
                Else
                    MsgBox "SDR GRID " & txtSDRGRID.Text & " not found.", vbInformation, "SDR GRID Not Found"
                    ValidateData = False
                    Exit Function
                End If
            Else
                MsgBox "Please select an SDR GRID.", vbInformation, "SDR GRID Not Selected"
                ValidateData = False
                Exit Function
            End If
        Else
            clsParamsPrj.intUseOwnSDR = 0
            clsParamsPrj.strSDRGridFileName = txtSDRGRID.Text
        End If
            
    End If
    
    'Managment Scenarios
    If ValidateMgmtScenario Then
        For i = 1 To grdLCChanges.Rows - 1
        If Len(grdLCChanges.TextMatrix(i, 2)) > 0 Then  'if they have entered one, then go ahead and add
            Set clsParamsPrj.clsMgmtScenItem = New clsXMLMgmtScenItem
                clsParamsPrj.clsMgmtScenItem.intID = i
                clsParamsPrj.clsMgmtScenItem.intApply = CInt(grdLCChanges.TextMatrix(i, 1))
                clsParamsPrj.clsMgmtScenItem.strAreaName = grdLCChanges.TextMatrix(i, 2)
                clsParamsPrj.clsMgmtScenItem.strAreaFileName = modUtil.GetFeatureFileName(grdLCChanges.TextMatrix(i, 2), m_App)
                clsParamsPrj.clsMgmtScenItem.strChangeToClass = grdLCChanges.TextMatrix(i, 3)
            clsParamsPrj.clsMgmtScenHolder.Add clsParamsPrj.clsMgmtScenItem
        End If
        Next i
    Else
        ValidateData = False
        grdLCChanges.SetFocus
        Exit Function
    End If
    
    'Pollutants
    If ValidatePollutants Then
        For i = 1 To grdCoeffs.Rows - 1
            'Adding a New Pollutantant Item to the Project file
            Set clsParamsPrj.clsPollItem = New clsXMLPollutantItem
                    clsParamsPrj.clsPollItem.intID = i
                    clsParamsPrj.clsPollItem.intApply = CInt(grdCoeffs.TextMatrix(i, 1))
                    clsParamsPrj.clsPollItem.strPollName = grdCoeffs.TextMatrix(i, 2)
                    clsParamsPrj.clsPollItem.strCoeffSet = grdCoeffs.TextMatrix(i, 3)
                    clsParamsPrj.clsPollItem.strCoeff = grdCoeffs.TextMatrix(i, 4)
                    clsParamsPrj.clsPollItem.intThreshold = CInt(grdCoeffs.TextMatrix(i, 5))
                    If grdCoeffs.TextMatrix(i, 6) <> "" Then
                        clsParamsPrj.clsPollItem.strTypeDefXMLFile = grdCoeffs.TextMatrix(i, 6)
                    End If
            clsParamsPrj.clsPollItems.Add clsParamsPrj.clsPollItem
        Next i
    Else
        ValidateData = False
        grdCoeffs.SetFocus
        Exit Function
    End If
        
    'Land Uses
    For i = 1 To grdLU.Rows - 1
        If Len(grdLU.TextMatrix(i, 2)) > 0 Then
        Set clsParamsPrj.clsLUItem = New clsXMLLandUseItem
            clsParamsPrj.clsLUItem.intID = i
            clsParamsPrj.clsLUItem.intApply = CInt(grdLU.TextMatrix(i, 1))
            clsParamsPrj.clsLUItem.strLUScenName = grdLU.TextMatrix(i, 2)
            clsParamsPrj.clsLUItem.strLUScenXMLFile = grdLU.TextMatrix(i, 3)
        clsParamsPrj.clsLUItems.Add clsParamsPrj.clsLUItem
        End If
    Next i
    
    'If it gets to here, all is well
    ValidateData = True
    
    m_XMLPrjParams.XML = clsParamsPrj.XML
    
    'Cleanup
    Set pLayer = Nothing
    
  Exit Function
ErrorHandler:
  HandleError False, "ValidateData " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Private Sub txtProjectName_Change()
  On Error GoTo ErrorHandler

'Make title of form = to what the user types in
    frmPrj.Caption = txtProjectName.Text

  Exit Sub
ErrorHandler:
  HandleError True, "txtProjectName_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Sub

Private Function ValidatePollutants() As Boolean
  On Error GoTo ErrorHandler

'Function to validate pollutants
    Dim i As Integer
    
    For i = 1 To grdCoeffs.Rows - 1
        'if the user isn't ignoring the pollutant, then check values
        If grdCoeffs.TextMatrix(i, 1) = 1 Then
            If Len(grdCoeffs.TextMatrix(i, 3)) = 0 Then
                MsgBox "Please select a coefficient set for pollutant: " & grdCoeffs.TextMatrix(i, 2), vbCritical, "Coefficient Set Missing"
                ValidatePollutants = False
                Exit Function
            Else
                If Len(grdCoeffs.TextMatrix(i, 4)) = 0 Then
                    MsgBox "Please select a coefficient for pollutant: " & grdCoeffs.TextMatrix(i, 2), vbCritical, "Coefficient Missing"
                    ValidatePollutants = False
                    Exit Function
                Else
                    ValidatePollutants = True
                End If
            End If
        Else
            ValidatePollutants = True
        End If
    Next i
    
  Exit Function
ErrorHandler:
  HandleError False, "ValidatePollutants " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Private Function ValidateWaterShed() As Boolean
  On Error GoTo ErrorHandler

'Validate the Watershed
    Dim rsWShed As New ADODB.Recordset
    Dim strWShed As String
    Dim booUpdate As Boolean
        
    Dim strDEM As String
    Dim strFlowDirFileName As String
    Dim strFlowAccumFileName As String
    Dim strFilledDEMFileName As String
    
    booUpdate = False
    
    'Select record from current cbo Selection
    strWShed = "SELECT * FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
    rsWShed.Open strWShed, g_ADOConn, adOpenDynamic, adLockOptimistic
    
    'Check to make sure all datasets exist, if not
    'DEM
    If Not modUtil.RasterExists(rsWShed!DEMFileName) Then
        MsgBox "Unable to locate DEM dataset: " & rsWShed!DEMFileName & ".", vbCritical, "Missing Dataset"
        strDEM = modUtil.BrowseForFileName("Raster", frmPrj, "Browse for DEM...")
        If Len(strDEM) > 0 Then
            rsWShed!DEMFileName = strDEM
            booUpdate = True
        Else
            ValidateWaterShed = False
            Exit Function
        End If
    'WaterShed Delineation
    ElseIf Not modUtil.FeatureExists(rsWShed!wsfilename) Then
        MsgBox "Unable to locate Watershed dataset: " & rsWShed!wsfilename & ".", vbCritical, "Missing Dataset"
        strWShed = modUtil.BrowseForFileName("Feature", frmPrj, "Browse for Watershed Dataset...")
        If Len(strWShed) > 0 Then
            rsWShed!wsfilename = strWShed
            booUpdate = True
        Else
            ValidateWaterShed = False
            Exit Function
        End If
    'Flow Direction
    ElseIf Not modUtil.RasterExists(rsWShed!FlowDirFileName) Then
        MsgBox "Unable to locate Flow Direction GRID: " & rsWShed!FlowDirFileName & ".", vbCritical, "Missing Dataset"
        strFlowDirFileName = modUtil.BrowseForFileName("Raster", frmPrj, "Browse for Flow Direction GRID...")
        If Len(strFlowDirFileName) > 0 Then
            rsWShed!FlowDirFileName = strFlowDirFileName
            booUpdate = True
        Else
            ValidateWaterShed = False
            Exit Function
        End If
    'Flow Accumulation
    ElseIf Not modUtil.RasterExists(rsWShed!FlowAccumFileName) Then
        MsgBox "Unable to locate Flow Accumulation GRID: " & rsWShed!FlowAccumFileName & ".", vbCritical, "Missing Dataset"
        strFlowAccumFileName = modUtil.BrowseForFileName("Raster", frmPrj, "Browse for Flow Accumulation GRID...")
        If Len(strFlowAccumFileName) > 0 Then
            rsWShed!FlowAccumFileName = strFlowAccumFileName
            booUpdate = True
        Else
            ValidateWaterShed = False
            Exit Function
        End If
    'Check for non-hydro correct GRIDS
    ElseIf rsWShed!HydroCorrected = 0 Then
        If Not modUtil.RasterExists(rsWShed!FilledDEMFileName) Then
            MsgBox "Unable to locate the Filled DEM: " & rsWShed!FilledDEMFileName & ".", vbCritical, "Missing Dataset"
            strFilledDEMFileName = modUtil.BrowseForFileName("Raster", frmPrj, "Browse for Filled DEM...")
            If Len(strFilledDEMFileName) > 0 Then
                rsWShed!FilledDEMFileName = strFilledDEMFileName
                booUpdate = True
            Else
                ValidateWaterShed = False
                Exit Function
            End If
        End If
    End If
    
    If booUpdate Then
        rsWShed.Update
    End If
    
    ValidateWaterShed = True
    
    rsWShed.Close
    Set rsWShed = Nothing
    


  Exit Function
ErrorHandler:
  HandleError False, "ValidateWaterShed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Private Function ValidateMgmtScenario() As Boolean
  On Error GoTo ErrorHandler

    
    Dim i As Integer
    Dim j As Integer
    
    For i = 1 To grdLCChanges.Rows - 1
    If grdLCChanges.TextMatrix(i, 1) <> "0" Then
        For j = 1 To grdLCChanges.Cols - 1
        
            Select Case j
                Case 2
                    If Len(grdLCChanges.TextMatrix(i, j)) > 0 Then
                        If Not (modUtil.LayerInMap(grdLCChanges.TextMatrix(i, j), m_pMap)) Then
                            ValidateMgmtScenario = False
                            Exit Function
                        End If
                    End If
                Case 3
                    If Len(grdLCChanges.TextMatrix(i, j)) > 0 Then
                    If grdLCChanges.TextMatrix(i, j) = "" Then
                        ValidateMgmtScenario = False
                        MsgBox "Please select a land class in cell " & i & " ," & j, vbCritical, "Missing Value"
                        grdLCChanges.row = i
                        grdLCChanges.col = j
                        Exit Function
                    End If
                    End If
            End Select
        Next j
    End If
    Next i
    
    ValidateMgmtScenario = True
    


  Exit Function
ErrorHandler:
  HandleError False, "ValidateMgmtScenario " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
End Function

Private Function EnableChkWaterShed() As Boolean
    
    Dim rsWShed As New ADODB.Recordset
    Dim strWShed As String

On Error GoTo ErrHandler:

    strWShed = "SELECT WSFILENAME FROM WSDELINEATION WHERE NAME LIKE '" & cboWSDelin.Text & "'"
    rsWShed.Open strWShed, g_ADOConn, adOpenStatic, adLockReadOnly
    
    m_strWShed = modUtil.SplitFileName(rsWShed!wsfilename)
    
    If m_pMap.SelectionCount > 0 Then
        EnableChkWaterShed = True
    Else
        EnableChkWaterShed = False
    End If
    
    rsWShed.Close
    Set rsWShed = Nothing

Exit Function

ErrHandler:
    EnableChkWaterShed = False
    
End Function


