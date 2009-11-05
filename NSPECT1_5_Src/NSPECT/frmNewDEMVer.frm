VERSION 5.00
Begin VB.Form frmNewWSDelin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Watershed Delineation"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7560
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4931.848
   ScaleMode       =   0  'User
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Create a new watershed delineation  "
      Height          =   3585
      Left            =   165
      TabIndex        =   10
      Top             =   120
      Width           =   7185
      Begin VB.CommandButton cmdBrowseDEMFile 
         Height          =   315
         Left            =   6315
         Picture         =   "frmNewDEMVer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   870
         Width           =   375
      End
      Begin VB.CheckBox chkStreamAgree 
         Caption         =   "Force Stream Agreement"
         Enabled         =   0   'False
         Height          =   273
         Left            =   2535
         TabIndex        =   4
         Top             =   2115
         Width           =   3600
      End
      Begin VB.ComboBox cboStreamLayer 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3795
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2565
         Width           =   2314
      End
      Begin VB.TextBox txtWSDelinName 
         Height          =   285
         Left            =   1950
         TabIndex        =   0
         Top             =   465
         Width           =   2000
      End
      Begin VB.ComboBox cboDEMUnits 
         Height          =   315
         ItemData        =   "frmNewDEMVer.frx":0876
         Left            =   1950
         List            =   "frmNewDEMVer.frx":0880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1290
         Width           =   2280
      End
      Begin VB.TextBox txtDEMFile 
         Height          =   285
         Left            =   1935
         TabIndex        =   7
         Top             =   885
         Width           =   4365
      End
      Begin VB.CheckBox chkHydroCorr 
         Caption         =   "Hydrologically Correct DEM"
         Height          =   273
         Left            =   1950
         TabIndex        =   3
         Top             =   1755
         Width           =   2587
      End
      Begin VB.ComboBox cboSubWSSize 
         Height          =   315
         ItemData        =   "frmNewDEMVer.frx":0892
         Left            =   1905
         List            =   "frmNewDEMVer.frx":089F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3015
         Width           =   2310
      End
      Begin VB.Label lblStream 
         AutoSize        =   -1  'True
         Caption         =   "Stream Layer"
         Enabled         =   0   'False
         Height          =   210
         Index           =   0
         Left            =   2565
         TabIndex        =   15
         Top             =   2565
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Delineation Name"
         Height          =   165
         Index           =   3
         Left            =   105
         TabIndex        =   14
         Top             =   480
         Width           =   1650
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DEM Units"
         Height          =   165
         Index           =   2
         Left            =   390
         TabIndex        =   13
         Top             =   1290
         Width           =   1365
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DEM Grid"
         Height          =   165
         Index           =   12
         Left            =   945
         TabIndex        =   12
         Top             =   885
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subwatershed Size"
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   11
         Top             =   3000
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5295
      TabIndex        =   8
      Top             =   3855
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6375
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frmNewWSDelin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmNewWSDelin
' *************************************************************************************************
' *  Description: Form for browsing and maintaining watershed delineation
' *  Scenarios within NSPECT
' *  Called By:  frmWSDelin menu item New...
' *************************************************************************************************
Option Explicit

Dim boolChange(3) As Boolean                'Array set to track changes in conrtols: On Change, cmdCreate is enabled
Public m_pInputDEMDS As IRasterDataset      'DEM Dataset
Private m_strDemArray() As String
Private m_strDemName As String
Private m_strFullDemName As String
Private m_strAccumFileName As String        'Accumulation GRID name
Private m_strDirFileName As String          'Direction GRID name
Private m_App As IApplication               'Application
Private m_strWorkspace As String
Dim WithEvents frmProg As frmProgress
Attribute frmProg.VB_VarHelpID = -1


'Hook ArcMap
Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub


Private Sub cboDEMUnits_Click()
    
    boolChange(2) = True
    CheckEnabled

End Sub

Private Sub cboSubWSSize_Click()
    
    boolChange(3) = True
    CheckEnabled

End Sub

Private Sub chkHydroCorr_Click()
    
    Select Case chkHydroCorr.Value
        Case 1
            chkStreamAgree.Enabled = True
        Case 0
            chkStreamAgree.Enabled = False
    End Select

End Sub

Private Sub chkStreamAgree_Click()
    
    Select Case chkStreamAgree.Value
        Case 1
            cboStreamLayer.Enabled = True
            lblStream.Item(0).Enabled = True
        Case 0
            cboStreamLayer.Enabled = False
        End Select

End Sub

Private Sub cmdBrowseDEMFile_Click()  'Load button for DEM File

    Dim pDEMRasterDataset As IRasterDataset
    Dim pDistUnit As ILinearUnit
    Dim intUnit As Integer
    Dim pProjCoord As IProjectedCoordinateSystem
    
    Set m_pInputDEMDS = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID", frmNewWSDelin, 0)
    Set pDEMRasterDataset = m_pInputDEMDS
    
    'Get the spatial reference
    If CheckSpatialReference(pDEMRasterDataset) Is Nothing Then
        MsgBox "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", vbExclamation, "No Project Information Detected"
        Exit Sub
    
    Else
        
        Set pProjCoord = CheckSpatialReference(pDEMRasterDataset)
        Set pDistUnit = pProjCoord.CoordinateUnit
        intUnit = pDistUnit.MetersPerUnit
        
        If intUnit = 1 Then
            cboDEMUnits.ListIndex = 0
        Else
            cboDEMUnits.ListIndex = 1
        End If
        
        cboDEMUnits.Refresh
    End If
    
    'Get the name
    m_strDemArray = Split(pDEMRasterDataset.CompleteName, "\")
    m_strDemName = m_strDemArray(UBound(m_strDemArray))
    
    
    Set pDEMRasterDataset = Nothing
    Set pDistUnit = Nothing
    Set pProjCoord = Nothing
    

End Sub

Private Sub cmdCreate_Click()

On Err GoTo ErrHandler:
        
    
    
    Screen.MousePointer = vbHourglass
    Dim intPos As Integer
    Dim pRasterDataset As IRasterDataset
    Dim FSO As FileSystemObject
    Dim folder As folder
    
    Dim pWorkspace As IWorkspace
    Dim pRasterWorkSpaceFactory As IWorkspaceFactory
    Dim pPropertySet As IPropertySet
    
    Set pRasterWorkSpaceFactory = New RasterWorkspaceFactory
    Set pPropertySet = New PropertySet

    
    If m_strDemName = "" Then
        'Get the names
        m_strDemArray = Split(txtDEMFile.Text, , "\")           'Array
        'm_strFullDemName = txtDEMFile.Text                     'Full Name
        m_strDemName = m_strDemArray(UBound(m_strDemArray))     'Name of Raster
    End If
    
    intPos = InStrRev(txtDEMFile.Text, "\", -1)
    
    m_strWorkspace = Left(txtDEMFile.Text, intPos)
        
    If modUtil.OpenRasterDataset(m_strWorkspace, m_strDemName) Is Nothing Then
        MsgBox "Error:  Could not open DEM Raster.", vbCritical, "Could Not Open Dataset"
             Exit Sub
    Else
        Set pRasterDataset = modUtil.OpenRasterDataset(m_strWorkspace, m_strDemName)
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set folder = FSO.CreateFolder(Environ("NSPECTDAT") & "\" & txtWSDelinName)
        
        Set pWorkspace = pRasterWorkSpaceFactory.OpenFromFile(Environ("NSPECTDAT") & "\" & txtWSDelinName, 0)
        DelineateWatershed pRasterDataset, pWorkspace
        
    End If
    
    'SQL Insert
    Dim strCmdInsert As String
    
    ' DataBase Update
    'Compose the INSERT statement.
    strCmdInsert = "INSERT INTO WSDelineation " & _
        "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName,  HydroCorrected, StreamFileName, SubWSSize) " & _
        " VALUES (" & _
        "'" & CStr(txtWSDelinName.Text) & "', " & _
        "'" & CStr(txtDEMFile.Text) & "', " & _
        "'" & cboDEMUnits.ListIndex & "', " & _
        "'" & m_strDirFileName & "', " & _
        "'" & m_strAccumFileName & "', " & _
        "'" & chkHydroCorr.Value & "', " & _
        "'" & cboStreamLayer.ListIndex & "', " & _
        "'" & cboSubWSSize.ListIndex & "'" & _
        ")"

    'Execute the statement.
    modUtil.g_ADOConn.Execute strCmdInsert, adCmdText
    
    Screen.MousePointer = vbNormal
    'Confirm
    MsgBox txtWSDelinName.Text & " successfully added.", vbOKOnly, "Record Added"
    
    Unload frmNewWSDelin
    Unload frmWSDelin
         
    Exit Sub
    
    
ErrHandler:
    
    MsgBox Err.Number & Err.Description & "Cmd_CreatE"
    
    

End Sub



Private Sub cmdQuit_Click()

   Unload Me

End Sub

Private Sub Form_Load()
        
    Set frmProg = frmProgress
        
    Dim i As Integer
    
    For i = 0 To UBound(boolChange)
        boolChange(i) = False
    Next i

End Sub



Private Sub txtDEMFile_Change()
    
    boolChange(1) = True
    CheckEnabled

End Sub

Private Sub txtWSDelinName_Change()

    boolChange(0) = True
    CheckEnabled
    

End Sub


Private Sub CheckEnabled()

    If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
        cmdCreate.Enabled = True
    Else
        cmdCreate.Enabled = False
    End If
    

End Sub

'Checks the form to make sure all is filled in correctly.
Private Function CheckForm() As Boolean


End Function

'***************************************************************************************************
Private Sub DelineateWatershed(pSurfaceDatasetIn As IRasterDataset, pWorkspace As IRasterWorkspace)
    
On Err GoTo ErrHandler
    
    'Map Stuff
    Dim pMap As IMap                            '[1]
    Dim pMxDoc As IMxDocument                   '[2]
    
    'Hydro Operations
    Dim pFlowDirHydrologyOp As IHydrologyOp     '[3]
    Dim pAccumHydrologyOp As IHydrologyOp       '[4]
    
    'Declare the geodataset objects
    Dim pFlowDirGDS As IGeoDataset              'Flow Direction GDS
    Dim pAccumGDS As IGeoDataset                'Flow Accumulation GDS
    Dim pSurface As IGeoDataset
    Dim pEnv As IRasterAnalysisEnvironment      'Analysis Environment

    'Flow direction bands and new dataset
    Dim pDirBands As IRasterBandCollection
    Dim pDirBand As IRasterBand
    Dim pDirNewGeo As IGeoDataset
    
    'Flow accumulation bands and new dataset
    Dim pAccumBands As IRasterBandCollection
    Dim pAccumBand As IRasterBand
    Dim pAccumNewGeo As IGeoDataset
    
    'Rasterdataset objects
    Dim pFlowDirRDS As IRasterDataset           'Flow Direction RDS
    Dim pAccumRDS As IRasterDataset             'Flow Accumulation RDS
    
    'Raster Layers
    Dim pFlowDirRasterLayer As IRasterLayer
    Dim pAccumRasterLayer As IRasterLayer
    
    
'------------------------------------------------------------------------------------------------
    'Set the map stuff up
    Set pMxDoc = m_App.Document
    Set pMap = pMxDoc.FocusMap
    
    ' First, check for Spatial Analyst License
     
    modUtil.CheckSpatialAnalystLicense
     
    'Have to hide the existing forms, to be able to show the progress
    Me.Hide
    frmWSDelin.Hide
   
    'Progress Form setup
    With frmProgress
        .prgBar.Max = 3
        .prgBar.Value = 1
        .lblProgress.Caption = "Hyrdrology Operations underway..."
        .Show
        .Refresh
    End With
    
   'Initialize the Hydro Ops
    Set pFlowDirHydrologyOp = New RasterHydrologyOp
    Set pAccumHydrologyOp = New RasterHydrologyOp
    
    'Set the Environment
    Set pEnv = pFlowDirHydrologyOp
    
    'Set pSurface to incoming RasterDataset -- QI
    Set pSurface = pSurfaceDatasetIn
    
    With pEnv
        Set .OutWorkspace = pWorkspace                          'Set workspace to incoming surface
        Set .OutSpatialReference = pSurface.SpatialReference    'Ditto with spatial reference
    End With
    
   
           
    'Set the output of the Hydro Operations
    frmProgress.prgBar.Value = 2
    frmProgress.lblProgress.Caption = "Creating Flow Direction..."
    
    Set pFlowDirGDS = pFlowDirHydrologyOp.FlowDirection(pSurface, True, True)  'Flow Direction
    
    frmProgress.lblProgress.Caption = "Creating Flow Accumulation..."
    frmProgress.lblProgress.Refresh
    
    Set pAccumGDS = pAccumHydrologyOp.FlowAccumulation(pFlowDirGDS)         'Flow Accumulation
    
    'Need to acquire the Rasterdataset from the prior two to save out
    Set pDirBands = pFlowDirGDS
    Set pDirBand = pDirBands.Item(0)
    Set pFlowDirRDS = pDirBand.RasterDataset      'RasterDataset
    
    'Have to get hold of the Temp dataset to save it permanently
    Dim pFlowTempDS As ITemporaryDataset
    Set pFlowTempDS = pFlowDirRDS
    
    Set pDirBands = Nothing
    Set pFlowDirRDS = Nothing
    Set pFlowDirGDS = Nothing
    Set pDirBand = Nothing
    
    'Save the flow direction as a new GRID
    
    frmProgress.prgBar.Value = 3
    frmProgress.lblProgress.Caption = "Saving GRIDS..."
    frmProgress.lblProgress.Refresh
    
    Set pDirNewGeo = pFlowTempDS.MakePermanentAs("Flowdir", pWorkspace, "GRID")
        
    'Again, get the RasterDataset, this time from the Accumulation GRID
    Set pAccumBands = pAccumGDS
    Set pAccumBand = pAccumBands.Item(0)
    Set pAccumRDS = pAccumBand.RasterDataset
    
    Dim pAccumTempDS As ITemporaryDataset
    Set pAccumTempDS = pAccumRDS
    
    Set pAccumBands = Nothing
    Set pAccumRDS = Nothing
    Set pAccumGDS = Nothing
    Set pAccumBand = Nothing
      
    
    'Save the flow accumulation as a new GRID
    Set pAccumNewGeo = pAccumTempDS.MakePermanentAs("flowacc", pWorkspace, "GRID")
    
'    'QI to get Dataset to make RasterLayer
    Set pFlowDirRDS = pDirNewGeo
    Set pAccumRDS = pAccumNewGeo
        
    'Set up RasterLayers
    Set pFlowDirRasterLayer = New RasterLayer
    Set pAccumRasterLayer = New RasterLayer
    
    'Create the RasterLayers
    pFlowDirRasterLayer.CreateFromDataset pFlowDirRDS
    pAccumRasterLayer.CreateFromDataset pAccumRDS
    
    'Add to Map
    pMap.AddLayer pFlowDirRasterLayer
    pMap.AddLayer pAccumRasterLayer
    
    Unload frmProgress
    
    Me.Show
    
    
    m_strAccumFileName = pAccumRasterLayer.FilePath
    m_strDirFileName = pFlowDirRasterLayer.FilePath
    
    'Clean up
    Set pMap = Nothing
    Set pMxDoc = Nothing
    Set pFlowDirHydrologyOp = Nothing
    Set pAccumHydrologyOp = Nothing
    Set pFlowDirGDS = Nothing
    Set pAccumGDS = Nothing
    Set pSurface = Nothing
    Set pEnv = Nothing
    Set pDirBands = Nothing
    Set pDirBand = Nothing
    Set pDirNewGeo = Nothing
    Set pAccumBands = Nothing
    Set pAccumBand = Nothing
    Set pAccumNewGeo = Nothing
    Set pFlowDirRDS = Nothing
    Set pAccumRDS = Nothing
    Set pFlowDirRasterLayer = Nothing
    Set pAccumRasterLayer = Nothing
    
   

Exit Sub

ErrHandler:
    MsgBox "Delineate Watershed Error " & Err.Source
    

End Sub


Private Sub txtWSDelinName_Validate(Cancel As Boolean)

    Dim FSO As FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
      
    Debug.Print Environ("NSPECTDAT") & "\" & txtWSDelinName.Text
      
    If FSO.FolderExists(Environ("NSPECTDAT") & "\" & txtWSDelinName.Text) Then
        MsgBox "A scenario named " & txtWSDelinName.Text & " already exists.  Please select another name.", vbCritical, "Duplicate Name"
        txtWSDelinName.SetFocus
    End If

End Sub
