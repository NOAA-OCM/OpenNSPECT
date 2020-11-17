VERSION 5.00
Begin VB.Form frmNewWSDelin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Watershed Delineation"
   ClientHeight    =   3135
   ClientLeft      =   3195
   ClientTop       =   2940
   ClientWidth     =   7035
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMain 
      Caption         =   "Create a new watershed delineation  "
      Height          =   2400
      Left            =   195
      TabIndex        =   7
      Top             =   135
      Width           =   6555
      Begin VB.CheckBox chkHydroCorr 
         Caption         =   "DEM is hyrdologically correct (filled)"
         Height          =   225
         Left            =   1785
         TabIndex        =   17
         Top             =   1140
         Width           =   3615
      End
      Begin VB.CommandButton cmdBrowseDEMFile 
         Height          =   315
         Left            =   5865
         Picture         =   "frmNewWSDelin.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   690
         Width           =   375
      End
      Begin VB.TextBox txtWSDelinName 
         Height          =   285
         Left            =   1770
         TabIndex        =   0
         Top             =   300
         Width           =   2000
      End
      Begin VB.ComboBox cboDEMUnits 
         Height          =   315
         ItemData        =   "frmNewWSDelin.frx":0876
         Left            =   1785
         List            =   "frmNewWSDelin.frx":0880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1485
         Width           =   2280
      End
      Begin VB.TextBox txtDEMFile 
         Height          =   285
         Left            =   1770
         TabIndex        =   4
         Top             =   720
         Width           =   4065
      End
      Begin VB.ComboBox cboSubWSSize 
         Height          =   315
         ItemData        =   "frmNewWSDelin.frx":0892
         Left            =   1785
         List            =   "frmNewWSDelin.frx":089F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1935
         Width           =   2265
      End
      Begin VB.Frame frmAdvanced 
         Caption         =   "Advanced Parameters (optional) "
         Height          =   1650
         Left            =   285
         TabIndex        =   12
         Top             =   2310
         Visible         =   0   'False
         Width           =   6090
         Begin VB.CommandButton cmdOptions 
            Caption         =   "Options..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   4215
            TabIndex        =   16
            Top             =   1065
            Width           =   885
         End
         Begin VB.ComboBox cboStreamLayer 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmNewWSDelin.frx":08B9
            Left            =   1830
            List            =   "frmNewWSDelin.frx":08C0
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1095
            Width           =   2314
         End
         Begin VB.CheckBox chkStreamAgree 
            Caption         =   "Force Stream Agreement"
            Enabled         =   0   'False
            Height          =   273
            Left            =   720
            TabIndex        =   13
            Top             =   735
            Width           =   2235
         End
         Begin VB.Label lblStream 
            AutoSize        =   -1  'True
            Caption         =   "Stream Layer:"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   735
            TabIndex        =   15
            Top             =   1140
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Delineation Name:"
         Height          =   195
         Index           =   3
         Left            =   255
         TabIndex        =   11
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DEM Units:"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   10
         Top             =   1530
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DEM Grid:"
         Height          =   195
         Index           =   12
         Left            =   255
         TabIndex        =   9
         Top             =   770
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Subwatershed Size:"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   1935
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4230
      TabIndex        =   5
      Top             =   2670
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5355
      TabIndex        =   6
      Top             =   2670
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

Private boolChange(3) As Boolean            'Array set to track changes in controls: On Change, cmdCreate is enabled
Public m_pInputDEMDS As IRasterDataset      'DEM Dataset
Public m_booProject As Boolean              'Boolean to determine if this was called from frmPrj

Private m_strDemArray() As String
Private m_strDemName As String              'DEM Name
Private m_strFullDemName As String          'DEM GRID name
Private m_strAccumFileName As String        'Accumulation GRID name
Private m_strDirFileName As String          'Direction GRID name
Private m_strFilledDEMFileName As String    'Filled DEM name
Private m_strWShedFileName As String        'WShed Name
Private m_strStreamLayer As String          'Stream Layer
Private m_strLSFileName As String           'Lenght Slope GRID File name
Private m_strNibbleName As String           'Nibble Grid File Name
Private m_strDEM2BName As String            '2 Cell DEM buffer grid path
Private m_strWorkspace As String            'Workspace

'Objects from the Watershed Delineation needed in the LS calc
Private m_pFilledDEMGDS As IGeoDataset
Private m_pFlowDirRS As IRasterDataset
Private m_intGridUnits As Integer           'Grid Units: 0 = meters, 1 = feet
Private m_intCellSize As Integer            'Cell Size of DEM Grid, used in Length Slope Calculation

Private m_App As IApplication               'Application
Private m_pMap As IMap                      'Map

Private Const dblSmall As Double = 0.03 '0.001    '
Private Const dblMedium As Double = 0.06 '0.01    '-Subwatershed sizes (3%, 6%, 10%)
Private Const dblLarge As Double = 0.1      '

Private m_intSize As Integer                'Index for Size Combo

Const c_sModuleFileName As String = "frmNewWSDelin.frm"
Private m_ParentHWND As Long          ' Set this to get correct parenting of Error handler forms

'Hook ArcMap
Public Sub init(ByVal pApp As IApplication)
  On Error GoTo ErrorHandler

    Set m_App = pApp
    Dim pDoc As IMxDocument
    Set pDoc = pApp.Document
    
    Set m_pMap = pDoc.FocusMap

  Exit Sub
ErrorHandler:
  HandleError True, "init " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Private Sub cboDEMUnits_Click()
  On Error GoTo ErrorHandler

    boolChange(2) = True
    CheckEnabled

  Exit Sub
ErrorHandler:
  HandleError True, "cboDEMUnits_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cboStreamLayer_Click()

On Error GoTo ErrorHandler
    
    If cboStreamLayer.Text <> "" Then
        cmdOptions.Enabled = True
    Else
        cmdOptions.Enabled = False
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "cboStreamLayer_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cboSubWSSize_Click()

On Error GoTo ErrorHandler

    boolChange(3) = True
    CheckEnabled
    m_intSize = cboSubWSSize.ListIndex

 Exit Sub
ErrorHandler:
  HandleError True, "cboSubWSSize_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub



Private Sub chkHydroCorr_Click()
  On Error GoTo ErrorHandler

    Select Case chkHydroCorr.Value
        Case 1
            chkStreamAgree.Enabled = True
            cboStreamLayer.Enabled = True
            cmdOptions.Enabled = True
            g_boolHydCorr = True
            
        Case 0
        
            chkStreamAgree.Enabled = False
            cboStreamLayer.Enabled = False
            cmdOptions.Enabled = False
            g_boolHydCorr = False
    
    End Select

    CheckEnabled
    

  Exit Sub
ErrorHandler:
  HandleError True, "chkHydroCorr_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub chkStreamAgree_Click()
  On Error GoTo ErrorHandler

    Select Case chkStreamAgree.Value
        Case 1
            cboStreamLayer.Enabled = True
            lblStream.Item(0).Enabled = True
            g_boolAgree = True
        Case 0
            cboStreamLayer.Enabled = False
            cmdOptions.Enabled = False
    End Select
    
    CheckEnabled

  Exit Sub
ErrorHandler:
  HandleError True, "chkStreamAgree_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdBrowseDEMFile_Click()  'Load button for DEM File

    Dim pDEMRasterDataset As IRasterDataset
    Dim pDistUnit As ILinearUnit
    Dim intUnit As Integer
    Dim pProjCoord As IProjectedCoordinateSystem
    
On Error GoTo ErrHandler:
    
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
    
ErrHandler:
    Exit Sub
    

End Sub

Private Sub cmdCreate_Click()
  
On Err GoTo ErrHandler:
 
    Screen.MousePointer = vbHourglass
    Dim intPos As Integer
    Dim pRasterDataset As IRasterDataset
    Dim fso As FileSystemObject
    Dim Folder As Folder
    
    Dim pWorkspace As IWorkspace
    Dim pRasterWorkspaceFactory As IWorkspaceFactory
    Dim pPropertySet As IPropertySet
    
    Set pRasterWorkspaceFactory = New RasterWorkspaceFactory
    Set pPropertySet = New PropertySet
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(txtDEMFile.Text) Or fso.FileExists(txtDEMFile.Text) Then
    
        If m_strDemName = "" Then
            'Get the names
            m_strDemArray = Split(txtDEMFile.Text, "\")            'Array
            m_strDemName = m_strDemArray(UBound(m_strDemArray))     'Name of Raster
        End If
        
    Else
        MsgBox "The File you have choosen does not exist", vbCritical, "File Not Found"
        txtDEMFile.SetFocus
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    intPos = InStrRev(txtDEMFile.Text, "\", -1)
    
    m_strWorkspace = Left(txtDEMFile.Text, intPos)
          
    Set pRasterDataset = modUtil.OpenRasterDataset(m_strWorkspace, m_strDemName)
    
    If pRasterDataset Is Nothing Then
        MsgBox "Error:  Could not open DEM Raster.", vbCritical, "Could Not Open Dataset"
        Exit Sub
    
    Else

        If Not fso.FolderExists(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text) Then
            Set Folder = fso.CreateFolder(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text)
        Else
            MsgBox "Name in use.  Please select another.", vbCritical, "Choose New Name"
            txtWSDelinName.SetFocus
            Exit Sub
        End If
                
        Set pWorkspace = pRasterWorkspaceFactory.OpenFromFile(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text, 0)
        
        'Give the call; if successful insert new record
        If DelineateWatershed(pRasterDataset, pWorkspace) Then
            
            'SQL Insert
            Dim strCmdInsert As String
            
            'DataBase Update
            'Compose the INSERT statement.
            strCmdInsert = "INSERT INTO WSDelineation " & _
                "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName," & _
                "FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName) " & _
                " VALUES (" & _
                "'" & CStr(txtWSDelinName.Text) & "', " & _
                "'" & CStr(txtDEMFile.Text) & "', " & _
                "'" & cboDEMUnits.ListIndex & "', " & _
                "'" & m_strDirFileName & "', " & _
                "'" & m_strAccumFileName & "', " & _
                "'" & m_strFilledDEMFileName & "', " & _
                "'" & chkHydroCorr.Value & "', " & _
                "'" & m_strStreamLayer & "', " & _
                "'" & cboSubWSSize.ListIndex & "', " & _
                "'" & m_strWShedFileName & "', " & _
                "'" & m_strLSFileName & "', " & _
                "'" & m_strNibbleName & "', " & _
                "'" & m_strDEM2BName & "')"
        
            'Execute the statement.
            modUtil.g_ADOConn.Execute strCmdInsert, adCmdText
            Screen.MousePointer = vbNormal
            
            'Confirm
            MsgBox txtWSDelinName.Text & " successfully added.", vbOKOnly, "Record Added"
            
            If g_boolNewWShed Then
                'frmPrj.Show
                frmPrj.Frame.Visible = True
                frmPrj.cboWSDelin.Clear
                modUtil.InitComboBox frmPrj.cboWSDelin, "WSDelineation"
                frmPrj.cboWSDelin.ListIndex = modUtil.GetCboIndex(txtWSDelinName.Text, frmPrj.cboWSDelin)
                Unload Me
            Else
                Unload frmNewWSDelin
                Unload frmWSDelin
            End If
        Else
            Screen.MousePointer = vbNormal
            Exit Sub
        End If
    End If

    'Reset project boolean
    m_booProject = False
    
    'Cleanup
    Set pRasterDataset = Nothing
    Set fso = Nothing
    Set Folder = Nothing
    Set pWorkspace = Nothing
    Set pRasterWorkspaceFactory = Nothing
    Set pPropertySet = Nothing
    
Exit Sub
    
ErrHandler:

    Screen.MousePointer = vbNormal
    MsgBox Err.Number & Err.Description & " :New Watershed Delineation"

End Sub

Private Sub cmdQuit_Click()
  On Error GoTo ErrorHandler

    Dim intvbYesNo As Integer
    
    intvbYesNo = MsgBox("Are you sure you want to exit?  Your changes will not be saved.", vbYesNo, "Exit")
    
    If intvbYesNo = vbYes Then
        If m_booProject Then
            frmPrj.cboWSDelin.ListIndex = 0
            Unload frmNewWSDelin
        Else
            Unload frmNewWSDelin
        End If
    Else
        Exit Sub
    End If
    
    m_booProject = False

  Exit Sub
ErrorHandler:
  HandleError True, "cmdQuit_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

    'Init bool variables
    g_boolAgree = False
    g_boolHydCorr = False
    g_boolParams = False
        
    Dim i As Integer
    
    For i = 0 To UBound(boolChange)
        boolChange(i) = False
    Next i

    AddFeatureLayerToComboBox cboStreamLayer, m_pMap, "line"
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error GoTo ErrorHandler
    
    m_booProject = False
    Set m_pMap = Nothing
    Set m_App = Nothing
    
  Exit Sub
ErrorHandler:
  HandleError True, "Form_Unload " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtDEMFile_Change()
  On Error GoTo ErrorHandler

    
    boolChange(1) = True
    CheckEnabled



  Exit Sub
ErrorHandler:
  HandleError True, "txtDEMFile_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub txtWSDelinName_Change()
  On Error GoTo ErrorHandler

    boolChange(0) = True
    CheckEnabled

  Exit Sub
ErrorHandler:
  HandleError True, "txtWSDelinName_Change " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


Public Sub CheckEnabled()
  On Error GoTo ErrorHandler

    If chkHydroCorr And chkStreamAgree Then
        If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) And g_boolParams Then
            cmdCreate.Enabled = True
        Else
            cmdCreate.Enabled = False
        End If
    ElseIf chkHydroCorr And Not chkStreamAgree Then
        If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
            cmdCreate.Enabled = True
        Else
            cmdCreate.Enabled = False
        End If
    ElseIf Not chkHydroCorr Then
        If boolChange(0) And boolChange(1) And boolChange(2) And boolChange(3) Then
            cmdCreate.Enabled = True
        Else
            cmdCreate.Enabled = False
        End If
    End If
        
  Exit Sub
ErrorHandler:
  HandleError True, "CheckEnabled " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub


'***************************************************************************************************
Private Function DelineateWatershed(pSurfaceDatasetIn As IRasterDataset, pWorkspace As IRasterWorkspace) As Boolean

On Error GoTo ErrorHandler

    'Map Stuff
    Dim pMap As IMap
    Dim pMxDoc As IMxDocument
    
    'Hydro Operations
    Dim pFillHydrologyOp As IHydrologyOp        ' Fill
    Dim pFlowDirHydrologyOp As IHydrologyOp     ' Flow Direction
    Dim pAccumHydrologyOp As IHydrologyOp       ' Flow Accumulation
    Dim pWaterShedOp As IHydrologyOp            ' Watershed
    Dim pStreamLinkOp As IHydrologyOp           ' Streamlink
    Dim pBasinOp As IHydrologyOp                'Creates polys out of watershed grid
    
    'Declare the geodataset objects
    Dim pFlowDirRaster As IRaster               'Flow Direction
    Dim pAccumRaster As IRaster                 'Flow Accumulation
    Dim pFillRaster As IRaster                  'Fill GDS
    Dim pWShedRaster As IRaster                 'WaterShed GDS
    Dim pBasinRaster As IRaster                 'Basin GDS
    Dim pSurface As IGeoDataset                 'Incoming surface
    
    Dim pEnv As IRasterAnalysisEnvironment      'Analysis Environment
    Dim pDEMRasterp As IRaster
    Dim pDEMRasterProps As IRasterProps
    
    Dim pAccumStats As IRasterStatistics        'Flow Accum Stats-used to get max
    Dim dblMax As Double
    Dim pExtr As IExtractionOp
    Dim pRasDes As IRasterDescriptor
    Dim pQueryFilter As IQueryFilter
    Dim pStream1 As IGeoDataset
    Dim pStream2 As IGeoDataset
    
    'Added 12/18
    Dim pEnvelope As IEnvelope
    'Set pEnvelope = New Envelope
        
    'Featureclass objects
    Dim pBasinFeatClass As IFeatureClass        'Basin Featureclass
    
    'Raster to Shape
    Dim pBasinRastConvert As IRasterConvertHelper
    Dim pShedRastConvert As IRasterConvertHelper
           
    Dim strProgTitle As String
    strProgTitle = "Watershed Delineation Processing..."
    
    'Set the map stuff up
    Set pMxDoc = m_App.Document
    Set pMap = pMxDoc.FocusMap
    
    ' First, check for Spatial Analyst License
    If Not modUtil.CheckSpatialAnalystLicense Then
        MsgBox "There is no Spatial Analyst license available.", vbCritical, "Spatial Analyst License Necessary"
        DelineateWatershed = False
        Exit Function
    End If
     
    'Have to hide the existing forms, to be able to show the progress
    If g_boolNewWShed Then
       Me.Hide
       frmPrj.Hide
    ElseIf Not g_boolNewWShed Then
       Me.Hide
       frmWSDelin.Hide
    End If
    
    'Initialize the Hydro Ops
    Set pFlowDirHydrologyOp = New RasterHydrologyOp
    Set pAccumHydrologyOp = New RasterHydrologyOp
    Set pFillHydrologyOp = New RasterHydrologyOp
    Set pWaterShedOp = New RasterHydrologyOp
    Set pBasinOp = New RasterHydrologyOp
    Set pExtr = New RasterExtractionOp
    Set pStreamLinkOp = New RasterHydrologyOp
    
    'Get the rasterprops through pDEMRasterp, a default raster of pSurfaceDatasetIn
    Set pDEMRasterp = pSurfaceDatasetIn.CreateDefaultRaster
    Set pDEMRasterProps = pDEMRasterp
    Set pSurface = pSurfaceDatasetIn

    'Get cell size from DEM; needed later in the Length Slope Calculation
    m_intCellSize = pDEMRasterProps.MeanCellSize.X
    
    'Expand the envelope 2 cell sizes to account for later analysis
    'pEnvelope.PutCoords pDEMRasterProps.Extent.XMin - (2 * m_intCellSize), _
            pDEMRasterProps.Extent.YMin - (2 * m_intCellSize), _
            pDEMRasterProps.Extent.XMax + (2 * m_intCellSize), _
            pDEMRasterProps.Extent.YMax + (2 * m_intCellSize)
                        
    Set pEnvelope = pDEMRasterProps.Extent
    pEnvelope.Expand m_intCellSize * 2, m_intCellSize * 2, False
                        
               
        'Set the Environment
    Set pEnv = pFillHydrologyOp
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pSurface.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope   'pDEMRasterProps.Extent
       .SetAsNewDefaultEnvironment
    End With
    
    'STEP 1:  Fill the Surface
    'if hydrocorrect, then skip the Fill, just use the incoming DEM
    If chkHydroCorr.Value = 1 Then
        Set pFillRaster = pSurfaceDatasetIn.CreateDefaultRaster
    Else
       
       'Call to ProgDialog to use throughout process: keep user informed.
       modProgDialog.ProgDialog "Filling DEM...", strProgTitle, _
       0, 10, 1, m_App.hwnd
       
       If modProgDialog.g_boolCancel Then
           Set pFillRaster = pFillHydrologyOp.Fill(pSurface)
       Else
           GoTo ProgCancel
       End If
       
    End If  'End if Fill
    
    'STEP 2: Flow Direction
    Set pEnv = pFlowDirHydrologyOp
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pSurface.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope 'pDEMRasterProps.Extent
    End With
    
    modProgDialog.ProgDialog "Computing Flow Direction...", strProgTitle, _
       0, 10, 2, m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
       Set pFlowDirRaster = pFlowDirHydrologyOp.FlowDirection(pFillRaster, False, False)
    Else
       GoTo ProgCancel
    End If
    
    'STEP 3: Flow Accumulation
    Set pEnv = pAccumHydrologyOp
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pSurface.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope 'pDEMRasterProps.Extent
    End With
    
    modProgDialog.ProgDialog "Computing Flow Accumulation...", strProgTitle, _
       0, 10, 3, m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
       Set pAccumRaster = pAccumHydrologyOp.FlowAccumulation(pFlowDirRaster)
    Else
       GoTo ProgCancel
    End If
    
    'STEP 4: Stream Layer
    'How it's done:
       '1. Get the Stat.Max of the Accum GRID
       '2. Multiply the Max * percentage as deemed by user: small = 0.01% med = 1% large = 10%
       '3. Query Accumulation GRID for all values large than result from #2
       '4. Use #3 in Watershed method with the Flow Direction Grid
    
    Dim pAccumBands As IRasterBandCollection
    Dim pAccumBand As IRasterBand
    
    Set pAccumBands = pAccumRaster
    Set pAccumBand = pAccumBands.Item(0)
    
    Set pAccumStats = pAccumBand.Statistics
    dblMax = pAccumStats.Maximum
                    
    'Initialize ExtractionOp
    Set pRasDes = New RasterDescriptor
    Set pQueryFilter = New QueryFilter
    
    Dim dblSubShedSize As Double
    
    Select Case m_intSize
       Case 0  'small
           dblSubShedSize = dblMax * dblSmall
       Case 1  'medium
           dblSubShedSize = dblMax * dblMedium
       Case 2  'large
           dblSubShedSize = dblMax * dblLarge
    End Select
    
    pQueryFilter.WhereClause = " value > " & dblSubShedSize
    pRasDes.Create pAccumRaster, pQueryFilter, "value"

    Set pEnv = pExtr
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pSurface.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope 'pDEMRasterProps.Extent
    End With
       
    Set pStream1 = pExtr.Attribute(pRasDes)
    
    'Step 5: Using Hydrology Op to create stream network
    Set pEnv = pStreamLinkOp
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pSurface.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope ' pDEMRasterProps.Extent
    End With
    
    modProgDialog.ProgDialog "Creating Stream Network...", strProgTitle, _
       0, 10, 4, m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
       Set pStream2 = pStreamLinkOp.StreamLink(pStream1, pFlowDirRaster)
    Else
       GoTo ProgCancel
    End If
    
    'Step 6: Do WaterShed Op
    Set pEnv = pWaterShedOp
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pSurface.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope 'pDEMRasterProps.Extent
    End With

    modProgDialog.ProgDialog "Creating Watershed GRID...", strProgTitle, _
       0, 10, 5, m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
       Set pWShedRaster = pWaterShedOp.Watershed(pFlowDirRaster, pStream2)
    Else
       GoTo ProgCancel
    End If
    
    Set pBasinRastConvert = New RasterConvertHelper
       
    Set pBasinFeatClass = pBasinRastConvert.ToShapefile(pWShedRaster, esriGeometryPolygon, pEnv)
    SetFeatureClassName pBasinFeatClass, "basinpolytemp"

    'STEP 7: Basin
    'Ed Removed Basin operation from the code 11/19/2007
    'Now using polygons output from the watershed.
'    Set pEnv = pBasinOp
'    With pEnv
'       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
'       Set .OutWorkspace = pWorkspace
'       Set .OutSpatialReference = pSurface.SpatialReference
'       .SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
'    End With
'
'    modProgDialog.ProgDialog "Creating Watershed Shapefile...", strProgTitle, _
'       0, 10, 6, m_App.hWnd
'
'    If modProgDialog.g_boolCancel Then
'        Set pBasinRaster = pBasinOp.Basin(pFlowDirRaster)
'
'        With pEnv
'            .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
'            Set .OutWorkspace = pWorkspace
'            Set .OutSpatialReference = pSurface.SpatialReference
'            .SetExtent esriRasterEnvValue, pDEMRasterProps.Extent
'        End With
'
'        Set pBasinRastConvert = New RasterConvertHelper
'
'        Set pBasinFeatClass = pBasinRastConvert.ToShapefile(pBasinRaster, esriGeometryPolygon, pEnv)
'        SetFeatureClassName pBasinFeatClass, "basinpolytemp"
'    Else
'       GoTo ProgCancel
'    End If

    'STEP 8: Get rid of small polys in the basin shapefile along the coast
    Dim pFinalBasinClass As IFeatureClass
    
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pSurface.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope 'pDEMRasterProps.Extent
    End With
    
    Set pFinalBasinClass = RemoveSmallPolys(pBasinFeatClass, pFillRaster, pEnv)
    'END STEP 7
    
    'Save the flow direction as a new GRID
    modProgDialog.ProgDialog "Saving Fill GRID...", strProgTitle, _
       0, 10, 7, m_App.hwnd

    If modProgDialog.g_boolCancel Then

       If Not Len(modUtil.MakePerminentGrid(pFillRaster, pEnv.OutWorkspace.PathName, "demfill")) > 0 Then
            MsgBox "Could Not Save DEM Fill GRID", vbCritical, "Error Saving File"
            DelineateWatershed = False
            Exit Function
       End If

       modProgDialog.ProgDialog "Saving Flow Direction GRID...", strProgTitle, _
       0, 10, 8, m_App.hwnd

       If Not Len(modUtil.MakePerminentGrid(pFlowDirRaster, pEnv.OutWorkspace.PathName, "flowdir")) > 0 Then
            MsgBox "Could Not Save Flow Direction GRID", vbCritical, "Error Saving File"
            DelineateWatershed = False
            Exit Function
       End If

       modProgDialog.ProgDialog "Saving Flow Accumulation...", strProgTitle, _
       0, 10, 9, m_App.hwnd

       If Not Len(modUtil.MakePerminentGrid(pAccumRaster, pEnv.OutWorkspace.PathName, "flowacc")) > 0 Then
            MsgBox "Could Not Save Flow Accumulation GRID", vbCritical, "Error Saving File"
            DelineateWatershed = False
            Exit Function
       End If

       modProgDialog.ProgDialog "Saving Watersheds", strProgTitle, _
       0, 10, 10, m_App.hwnd
       If Not Len(modUtil.MakePerminentGrid(pWShedRaster, pEnv.OutWorkspace.PathName, "wshed")) > 0 Then
            MsgBox "Could Not Save Watershed GRID", vbCritical, "Error Saving File"
            DelineateWatershed = False
            Exit Function
       End If

    Else
       GoTo ProgCancel
    End If
    
    modProgDialog.KillDialog
    
    'With all of that done, now go get the name of the LS Grid while actually computing said LS Grid
    m_strLSFileName = CalcLengthSlope(pFillRaster, pFlowDirRaster, pAccumRaster, pEnv, "0", pWorkspace)
            
    'Now get file paths to throw back in Database
    m_strAccumFileName = pEnv.OutWorkspace.PathName & "flowacc"
    m_strDirFileName = pEnv.OutWorkspace.PathName & "flowdir"
    
    If chkHydroCorr.Value = 1 Then
        m_strFilledDEMFileName = txtDEMFile.Text
    Else
        m_strFilledDEMFileName = pEnv.OutWorkspace.PathName & "demfill"
    End If
    
    m_strWShedFileName = pEnv.OutWorkspace.PathName & "basinpoly"
    
    DelineateWatershed = True
    
    'Clean up
    Set pMap = Nothing
    Set pMxDoc = Nothing
    Set pSurface = Nothing
    Set pFlowDirRaster = Nothing
    Set pAccumRaster = Nothing
    Set pFillRaster = Nothing
    Set pWShedRaster = Nothing
    Set pBasinRaster = Nothing
    Set pEnv = Nothing
    Set pAccumBands = Nothing
    Set pAccumBand = Nothing
    Set pBasinRastConvert = Nothing
    'Hydro Operations
    Set pFillHydrologyOp = Nothing
    Set pFlowDirHydrologyOp = Nothing
    Set pAccumHydrologyOp = Nothing
    Set pWaterShedOp = Nothing
    Set pStreamLinkOp = Nothing
    Set pBasinOp = Nothing
    
Exit Function
ProgCancel:
    modProgDialog.KillDialog
    MsgBox "The watershed delineation process has been stopped by the user.  Changes have been discarded.", vbCritical, "Process Stopped"
    DelineateWatershed = False
    
ErrorHandler:
  If Err.Number = -2147217297 Then 'User cancelled operation
    modProgDialog.g_boolCancel = False
    DelineateWatershed = False
  Else
    HandleError False, "DelineateWatershed " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
    DelineateWatershed = False
    modProgDialog.KillDialog
  End If
  
End Function

Private Function RemoveSmallPolys(pFeatureClass As IFeatureClass, pDEMRaster As IRaster, pEnv As IRasterAnalysisEnvironment) As IFeatureClass
  On Error GoTo ErrorHandler


    Dim pAlgebraOp As IMapAlgebraOp
    Dim pDEMOutline As IRaster
    Dim pDEMGeo As IGeoDataset
    Dim pDEMFeatClass As IFeatureClass
    Dim pRasterConvert As IConversionOp
    
    'Basin featureclass editing stuff
    Dim pWorkspace As IWorkspace
    Dim pFeatureCursor As IFeatureCursor
    Dim pFeature As IFeature
    Dim pArea As IArea
    Dim pExtension As IExtension
    Dim pExtMgr As IExtensionManager
    Dim pEditor As IEditor
    Dim pUID As New UID
    
    'Union goodies
    Dim pDEMTable As ITable
    Dim pBasinTable As ITable
    Dim pFeatClassName As IFeatureClassName
    Dim pNewWSName As IWorkspaceName
    Dim pDatasetName As IDatasetName
    Dim pOutputFeatClass As IFeatureClass
    Dim pBGP As IBasicGeoprocessor
    Dim pEnvNew As IRasterAnalysisEnvironment
    Dim pTempDEM As IDataset
    Dim pTempBasin As IDataset
        
    Dim strExpression As String
    Dim dblArea As Double
    Dim strWorkspace As String
    Set pEnvNew = pEnv
    
    strWorkspace = pEnvNew.OutWorkspace.PathName
        
    Set pAlgebraOp = New RasterMapAlgebraOp
    Set pEnvNew = pAlgebraOp

'#1 First step, get a 1, 0 representation of the DEM, so we can get the outline
    pAlgebraOp.BindRaster pDEMRaster, "DEMo"
    strExpression = "Con([DEMo] >= 0, 1, 0)"

    Set pDEMOutline = pAlgebraOp.Execute(strExpression)

    pAlgebraOp.UnbindRaster "DEMo"
        
'#2 init the rasterconverter and create the poly
    Set pRasterConvert = New RasterConversionOp
    'Convert the DEM outline to a polygon and save as 'demoutline'
    Set pWorkspace = SetFeatureShapeWorkspace(strWorkspace)
    Set pDEMFeatClass = pRasterConvert.ToFeatureData(pDEMOutline, esriGeometryPolygon, pWorkspace, "demout")
    
    modUtil.AddFeatureLayer m_App, pDEMFeatClass, "demout"
    modUtil.AddFeatureLayer m_App, pFeatureClass, "basinpolytemp"
    
'#3 determine size of 'small' watersheds, this is the area
    'Number of cells in the DEM area that are not null * a number dave came up with * CellSize Squared
    'dblArea = ((modUtil.ReturnCellCount(pDEMRaster)) * 0.004) * m_intCellSize * m_intCellSize
    
'#4 Now with the Area of small sheds determined we can remove polygons that are too small.  To do this
'   simply loop through the features and test the area.
    'Set pFeatureCursor = pFeatureClass.Update(Nothing, False)
    'Set pFeature = pFeatureCursor.NextFeature
    
    'Set up the Editor
    'pUID.Value = "{F8842F20-BB23-11D0-802B-0000F8037368}"

    'Set pExtMgr = m_App
    'Set pExtension = pExtMgr.FindExtension(pUID)
    'Set pEditor = pExtension
    
    'pEditor.StartEditing pWorkspace
    
    'Do While Not pFeature Is Nothing
    '    Set pArea = pFeature.Shape
        
    '    If pArea.Area <= dblArea Then
    '        pFeatureCursor.DeleteFeature
    '    End If
        
    '    Set pFeature = pFeatureCursor.NextFeature
    'Loop
    
    'Stop and Save
    'pFeatureCursor.Flush
    'pEditor.StopEditing True
    
    'Set pFeatureCursor = Nothing
        
    'Have to add the damn thing in to do the union in the next step
    'modUtil.AddFeatureLayer m_App, pFeatureClass, "Basin Polygons"
    
'5#  Now, time to union the outline of the of the DEM with the newly paired down basin poly
    Set pDEMTable = m_pMap.Layer(1)
    Set pBasinTable = m_pMap.Layer(0)
    
    Set pFeatClassName = New FeatureClassName
        
    ' Set output location and feature class name
    Set pNewWSName = New WorkspaceName
    pNewWSName.WorkspaceFactoryProgID = "esriDataSourcesFile.ShapeFileWorkspaceFactory.1"
    pNewWSName.PathName = pWorkspace.PathName
    
    Set pDatasetName = pFeatClassName
    pDatasetName.Name = "basinpoly"

    Set pDatasetName.WorkspaceName = pNewWSName
    
    ' Set the tolerance.  Passing 0.0 causes the default tolerance to be used.
    ' The default tolerance is 1/10,000 of the extent of the data frame's spatial domain
    Dim tol As Double
    tol = 0
    
    ' Perform the union
    Set pBGP = New BasicGeoprocessor
    Set pOutputFeatClass = pBGP.Union(pDEMTable, False, pBasinTable, False, tol, pFeatClassName)
    
    Set RemoveSmallPolys = pOutputFeatClass
        
    'Cleanup
    'Remove the layers first
    m_pMap.DeleteLayer m_pMap.Layer(0)
    m_pMap.DeleteLayer m_pMap.Layer(0)
    
    Set pTempDEM = pDEMFeatClass
    Set pTempBasin = pFeatureClass
    
    If pTempDEM.CanDelete Then
        pTempDEM.Delete
    End If
    
    If pTempBasin.CanDelete Then
        pTempBasin.Delete
    End If
    
    'Cleanup
    Set pAlgebraOp = Nothing
    Set pDEMOutline = Nothing
    Set pDEMFeatClass = Nothing
    Set pRasterConvert = Nothing
    Set pWorkspace = Nothing
    Set pFeature = Nothing
    Set pArea = Nothing
    Set pEditor = Nothing
    Set pDEMTable = Nothing
    Set pBasinTable = Nothing
    Set pFeatClassName = Nothing
    Set pNewWSName = Nothing
    Set pDatasetName = Nothing
    Set pBGP = Nothing
    Set pExtMgr = Nothing
    Set pTempDEM = Nothing
    Set pTempBasin = Nothing
    Set pEnvNew = Nothing
    Set pEditor = Nothing
    Set pExtension = Nothing

  Exit Function
ErrorHandler:
      HandleError False, "RemoveSmallPolys " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND
    
End Function



Private Sub txtWSDelinName_Validate(Cancel As Boolean)
  On Error GoTo ErrorHandler
    
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
      
    If fso.FileExists(modUtil.g_nspectPath & "\" & txtWSDelinName.Text) Then
        MsgBox "A scenario named " & txtWSDelinName.Text & " already exists.  Please select another name.", vbCritical, "Duplicate Name"
        txtWSDelinName.SetFocus
    End If

  Exit Sub
ErrorHandler:
  HandleError True, "txtWSDelinName_Validate " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Function ReturnRasterDataset(pInComingGeoDataSet As IGeoDataset, strName As String, pWorkspace As IWorkspace) As IGeoDataset
  On Error GoTo ErrorHandler

    Dim pBands As IRasterBandCollection
    Dim pBand As IRasterBand
    Dim pRasterDataset As IRasterDataset
    Dim pNewGeo As IGeoDataset

    Set pBands = pInComingGeoDataSet
    Set pBand = pBands.Item(0)
    Set pRasterDataset = pBand.RasterDataset
    
    Dim pTempDS As ITemporaryDataset
    Set pTempDS = pRasterDataset
    
MsgBox "Rename Workspace: " & pWorkspace.PathName
MsgBox "rename type: " & pWorkspace.Type

    Set ReturnRasterDataset = pTempDS.MakePermanentAs(strName, pWorkspace, "GRID")
    
    Set pRasterDataset = Nothing
    Set pBands = Nothing
    Set pBand = Nothing
    Set pInComingGeoDataSet = Nothing
    



  Exit Function
ErrorHandler:
  HandleError False, "ReturnRasterDataset " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function


Private Function CalcLengthSlope(pDEMRaster As IRaster, pFlowDirRaster As IRaster, pAccumRaster As IRaster, pEnv As IRasterAnalysisEnvironment, _
                                 strUnits As String, pWS As IWorkspace) As String
    
'From the Delineate watershed, we are garnering the DEM, Flow Direction, flowaccum, environment and units
'to be used in this function
'Returns: Name of the LS File for use in the database table
On Error GoTo ErrHandler:

    'The rasters                                'Associated Steps
    Dim pDEMOneCell As IRaster              'STEP 1: 1 Cell Buffer DEM
    Dim pDEMTwoCell As IRaster              'STEP 2: 2 Cell Buffer DEM
    Dim pFlowDir As IRaster                 'STEP 3: Flow Direction
    Dim pFlowDirBV As IRaster               'STEP 4: Pre Nibble Null
    Dim pMask As IRaster                    'STEP 4: Mask
    Dim pNibble As IRaster                  'STEP 4: Nibble
    Dim pDownSlope As IRaster               'STEP 5: Down Slope
    Dim pDownAngle As IRaster               'STEP 5a: Tweak down slope
    Dim pRelativeSlope As IRaster           'STEP 6: Relative Slope
    Dim pRelSlopeThreshold As IRaster       'STEP 7: Relative slope threshold
    Dim pSlopeBreak As IRaster              'STEP 7a: Slope Break
    Dim pFlowDirBreak As IRaster            'STEP 8: Flow Direction Break
    Dim pWeight As IRaster                  'STEP 9: Weight GRID
    Dim pFlowLength As IRaster              'STEP 10: Flow Length
    Dim pFlowLengthFt As IRaster            'STEP 11: Flow Length to Feet
    Dim pSlopeExp As IRaster                'STEP 12: Slope Exponent
    Dim pRusleLFactor As IRaster            'STEP 13: Rusle L Factor
    Dim pRusleSFactor As IRaster            'STEP 14: Rusle S Factor
    Dim pLSFactor As IRaster                'STEP 15: LS Factor
    Dim pFinalLS As IRaster                 'STEP 15a: clippage
    Dim strProgTitle As String

    'Analysis Environment
    Dim pDEMGeoDS As IGeoDataset            'GeoDataset to get spat. ref
    Dim pSpatRef As ISpatialReference       'Spatial Reference
    
    'Create Map Algebra Operator
    Dim pMapAlgebraOp As IMapAlgebraOp      'Workhorse
    'String to hold calculations
    Dim strExpression As String
    
    Set pDEMGeoDS = pDEMRaster
    Set pSpatRef = pDEMGeoDS.SpatialReference
    
    If Not modUtil.CheckSpatialAnalystLicense Then
        MsgBox "No Spatial Analyst License Available.", vbCritical, "Pay Your Licensing Fee"
        CalcLengthSlope = "ERROR"
        Exit Function
    End If
    
    If pDEMRaster Is Nothing Or pFlowDirRaster Is Nothing Then
        MsgBox "CaclLengthSlope Error."
        CalcLengthSlope = "ERROR"
    End If
    
    'Initialize the Map AlgebraOp, same thing as the Map Calculator
    'All of the following steps use the same methodology.  They take a Raster, bind it
    'to a symbol, in this case a string that represents them in some map calculation.
    'You then simply use the MapAlgebraOp to execute the expression.
    
    'Get a hold of the Spatial Reference of the DEM, and init MapAlgebra
    Set pMapAlgebraOp = New RasterMapAlgebraOp
      
    'Set the environment
    Set pEnv = pMapAlgebraOp
    
    strProgTitle = "Processing the LS GRID..."
    
    'STEP 1: ----------------------------------------------------------------------
    'Buffer the DEM by one cell
    modProgDialog.ProgDialog "Creating one cell buffer...", strProgTitle, _
    0, 15, 1, m_App.hwnd
    
    Set pDEMOneCell = Nothing

    If modProgDialog.g_boolCancel Then
      
        pMapAlgebraOp.BindRaster pDEMRaster, "aml_fdem"
        strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"
    
        Set pDEMOneCell = pMapAlgebraOp.Execute(strExpression)
    
        pMapAlgebraOp.UnbindRaster "aml_fdem"
    
    Else
        GoTo ProgCancel
    End If
    
    'END STEP 1: ------------------------------------------------------------------
    
    'STEP 2: ----------------------------------------------------------------------
    'Buffer the DEM buffer by one more cell, that's 2
    modProgDialog.ProgDialog "Creating two cell buffer...", strProgTitle, _
    0, 15, 2, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        pMapAlgebraOp.BindRaster pDEMOneCell, "dem_b"
        strExpression = "Con(isnull([dem_b]), focalmin([dem_b]), [dem_b])"
    
        Set pDEMTwoCell = pMapAlgebraOp.Execute(strExpression)
        
        m_strDEM2BName = modUtil.MakePerminentGrid(pDEMTwoCell, pWS.PathName, "dem2b")
    
        pMapAlgebraOp.UnbindRaster "dem_b"
    
    Else
        GoTo ProgCancel
    End If
    'END STEP 2: ------------------------------------------------------------------
    
    'STEP 3: ----------------------------------------------------------------------
    'Flow Direction
        
    Set pFlowDir = pFlowDirRaster
    
    'END STEP 3: ------------------------------------------------------------------

    
    'STEP 3a: ---------------------------------------------------------------------
    modProgDialog.ProgDialog "Creating mask...", strProgTitle, _
    0, 15, 3, m_App.hwnd

    If modProgDialog.g_boolCancel Then

        With pMapAlgebraOp
            .BindRaster pDEMOneCell, "mask"
        End With

        'strExpression = "con(isnull([mask]),0,1)"
        strExpression = "con([mask] >= 0, 1, 0)"


        Set pMask = pMapAlgebraOp.Execute(strExpression)

        With pEnv
            Set .Mask = pMask
            Set .OutWorkspace = pWS
        End With

        Set pEnv = pMapAlgebraOp
    Else
        GoTo ProgCancel
    End If
   
    'STEP 4: ----------------------------------------------------------------------
    'Buffering Flow Direction and do the nibble to fill it in
    'Needed in case there is outflow from the DEM grid.
    'The following algorithms need to access the downslope DEM grid cell.
    'We find this by nibbling the original flow direction grid instead of recalculating
    'flow direction from the nibbled DEM because we want the elevation that is assumed
    'to be downstream of the edge cell.  Using flow direction on the buffered DEM
    'may not give that same result.
    modProgDialog.ProgDialog "Buffering slope direction...", strProgTitle, _
    0, 15, 4, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        With pMapAlgebraOp
            .BindRaster pFlowDir, "fdr_b"
        End With
        strExpression = "con(isnull([fdr_b]),0,[fdr_b])"
    
        Set pFlowDirBV = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "fdr_b"
    
        'Nibble
        pMapAlgebraOp.BindRaster pFlowDirBV, "fdr_bv"
        pMapAlgebraOp.BindRaster pMask, "waia_reg"
        strExpression = "nibble([fdr_bv],[waia_reg], dataonly)"
    
        Set pNibble = pMapAlgebraOp.Execute(strExpression)
        
        'Get nibble's path for use in the database
        m_strNibbleName = modUtil.MakePerminentGrid(pNibble, pWS.PathName, "nibble")
        
        With pMapAlgebraOp
            .UnbindRaster "fdr_bv"
            .UnbindRaster "waia_reg"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 4: ------------------------------------------------------------------

    'STEP 5: ----------------------------------------------------------------------
    'Calculate Slope
    'Actually this is calculating the SLOPE, in degrees, not the slope change.
    'Note that ESRI's SLOPE command should not be used here.  That command fits
    'a plane to the 3x3 grid surrounding the central point and assigns the central
    'point the slope of that plane.  The algorithm used here calculates only the slope
    'between the central point and it's immediate downstream neighbor.
    'That is what is needed by RUSLE.

    modProgDialog.ProgDialog "Calculating Slope change...", strProgTitle, _
    0, 15, 5, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        With pMapAlgebraOp
            .BindRaster pNibble, "fdrnib"
            .BindRaster pDEMTwoCell, "dem_2b"
        End With
        strExpression = "Con(([fdrnib] ge 0.5 and [fdrnib] lt 1.5), deg * atan(([dem_2b] - [dem_2b](1,0)) / (" & m_intCellSize & "))," & _
            "Con(([fdrnib] ge 1.5 and [fdrnib] lt 3.0), deg * atan(([dem_2b] - [dem_2b](1,1)) / (" & m_intCellSize & " * 1.4142))," & _
            "Con(([fdrnib] ge 3.0 and [fdrnib] lt 6.0), deg * atan(([dem_2b] - [dem_2b](0,1)) / (" & m_intCellSize & "))," & _
            "Con(([fdrnib] ge 6.0 and [fdrnib] lt 12.0), deg * atan(([dem_2b] - [dem_2b](-1,1)) / (" & m_intCellSize & " * 1.4142))," & _
            "Con(([fdrnib] ge 12.0 and [fdrnib] lt 24.0), deg * atan(([dem_2b] - [dem_2b](-1,0)) / (" & m_intCellSize & "))," & _
            "Con(([fdrnib] ge 24.0 and [fdrnib] lt 48.0), deg * atan(([dem_2b] - [dem_2b](-1,-1)) / (" & m_intCellSize & " * 1.4142))," & _
            "Con(([fdrnib] ge 48.0 and [fdrnib] lt 96.0), deg * atan(([dem_2b] - [dem_2b](0,-1)) / (" & m_intCellSize & "))," & _
            "Con(([fdrnib] ge 96.0 and [fdrnib] lt 192.0), deg * atan(([dem_2b] - [dem_2b](1,-1)) / (" & m_intCellSize & " * 1.4142))," & _
            "Con(([fdrnib] ge 192.0 and [fdrnib] le 255.0), deg * atan(([dem_2b] - [dem_2b](1,0)) / (" & m_intCellSize & "))," & _
            "0.1 )))))))))"
            
        Set pDownSlope = pMapAlgebraOp.Execute(strExpression)
        'Cleanup
        With pMapAlgebraOp
            .UnbindRaster "fdrnib"
            .UnbindRaster "dem_2b"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 5 ---------------------------------------------------------------------
    
    'STEP 5a: -----------------------------------------------------------------------
    'Tweak slope where it equals 0 to 0.1
    pMapAlgebraOp.BindRaster pDownSlope, "dwnsltmp"
    strExpression = "Con([dwnsltmp] le 0, 0.1,[dwnsltmp])"

    Set pDownAngle = pMapAlgebraOp.Execute(strExpression)
    
    pMapAlgebraOp.UnbindRaster "dwnsltmp"
    
    'END STEP 5a: -------------------------------------------------------------------

    'STEP 6: ------------------------------------------------------------------------
    'Relative Slope Change
    modProgDialog.ProgDialog "Calculating Relative Slope Change...", strProgTitle, _
    0, 15, 6, m_App.hwnd

    If modProgDialog.g_boolCancel Then
        With pMapAlgebraOp
            .BindRaster pDownAngle, "dwnslangle"
            .BindRaster pNibble, "fdrnib"
        End With
        strExpression = "Con([fdrnib] == 1, ([dwnslangle] - [dwnslangle](1,0)) / [dwnslangle]," & _
            "Con([fdrnib] == 2, ([dwnslangle] - [dwnslangle](1,1)) / [dwnslangle]," & _
            "Con([fdrnib] == 4, ([dwnslangle] - [dwnslangle](0,1)) / [dwnslangle]," & _
            "Con([fdrnib] == 8, ([dwnslangle] - [dwnslangle](-1,1)) / [dwnslangle]," & _
            "Con([fdrnib] == 16, ([dwnslangle] - [dwnslangle](-1,0)) / [dwnslangle]," & _
            "Con([fdrnib] == 32, ([dwnslangle] - [dwnslangle](-1,-1)) / [dwnslangle]," & _
            "Con([fdrnib] == 64, ([dwnslangle] - [dwnslangle](0,-1)) / [dwnslangle]," & _
            "Con([fdrnib] == 128, ([dwnslangle] - [dwnslangle](1,-1)) / [dwnslangle]," & _
            "Con([fdrnib] == 255, ([dwnslangle] - [dwnslangle](1,0)) / [dwnslangle]," & _
            "0.1 )))))))))"
    
        Set pRelativeSlope = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "dwnslangle"
            .UnbindRaster "fdrnib"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 6 -----------------------------------------------------------------------

    'STEP 7: --------------------------------------------------------------------------
    'Identify breakpoints: relative difference where slope angle exceeds threshold values
    modProgDialog.ProgDialog "Identifying breakpoints...", strProgTitle, _
    0, 15, 7, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        pMapAlgebraOp.BindRaster pDownAngle, "dwnslangle"
        strExpression = "Con(([dwnslangle] gt 2.86240), 0.5, Con(([dwnslangle] le 2.86240), 0.7, 0.0 ))"
    
        Set pRelSlopeThreshold = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "dwnslangle"
    'END STEP 7 -----------------------------------------------------------------------
    
    'STEP 7a: -------------------------------------------------------------------------
    'Slope break
        With pMapAlgebraOp
            .BindRaster pRelSlopeThreshold, "threshold"
            .BindRaster pRelativeSlope, "delslprel"
        End With
        strExpression = "Con(([delslprel] gt [threshold]), 1, 0 )"
    
        Set pSlopeBreak = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "threshold"
            .UnbindRaster "delslprel"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 7a -------------------------------------------------------------------------

    'STEP 8 -------------------------------------------------------------------------------
    'Create Modified Flow Direction GRID
    modProgDialog.ProgDialog "Creating modified flow direction GRID...", strProgTitle, _
    0, 15, 8, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        With pMapAlgebraOp
            .BindRaster pSlopeBreak, "slopebreak"
            .BindRaster pFlowDir, "fdr"
        End With
        strExpression = "Con([slopebreak] eq 0, [fdr], 0)"
    
        Set pFlowDirBreak = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "slopebreak"
            .UnbindRaster "fdr"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 8: ----------------------------------------------------------------------

    'STEP 9: --------------------------------------------------------------------------
    'Create weight grid
    modProgDialog.ProgDialog "Creating weight GRID...", strProgTitle, _
    0, 15, 9, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        
        'Dave's comments:
        'This is an error.  In the original AML code, it is needed to correctly account for
        'diagonal flow in the flow length calculation of step 10.  However, ArcMap already
        'makes this correction.  There is another weighting function needed, however, to be
        'consistent with the procedure used in the original AML code. That is what should replace this con.
        
        'Removed 12/19/07
        'pMapAlgebraOp.BindRaster pFlowDir, "fdr"
        'strExpression = "Con([fdr] eq 2, 1.41421," & _
        '    "Con([fdr] eq 8, 1.41421," & _
        '    "Con([fdr] eq 32, 1.41421," & _
        '    "Con([fdr] eq 128, 1.41421," & _
        '    "1.0))))"
        'End Remove
            
        'Added 12/19/07
        'pAccumRaster is passed over from delin watershed
        pMapAlgebraOp.BindRaster pAccumRaster, "flowacc"
        strExpression = "Con([flowacc] eq 0, 0.5,1.0)"
    
        Set pWeight = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "flowacc"
        'End Added
    Else
        GoTo ProgCancel
    End If
    'END STEP 9 -------------------------------------------------------------------------

    'STEP 10: ---------------------------------------------------------------------------
    'Flow Length GRID
    modProgDialog.ProgDialog "Creating flow length GRID...", strProgTitle, _
    0, 15, 10, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        With pMapAlgebraOp
            .BindRaster pFlowDirBreak, "fdrbrk"
            .BindRaster pWeight, "weight"
        End With
        strExpression = "FlowLength([fdrbrk], [weight], UPSTREAM)"
    
        Set pFlowLength = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "fdrbrk"
            .UnbindRaster "weight"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 10: -----------------------------------------------------------------------

    'STEP 11: ---------------------------------------------------------------------------
    'Convert Meters To Feet
    'TODO: Check measure units, won't have to do if already in Feet
    modProgDialog.ProgDialog "Checking measurement units...", strProgTitle, _
    0, 15, 11, m_App.hwnd

    If modProgDialog.g_boolCancel Then
        pMapAlgebraOp.BindRaster pFlowLength, "flowlen"
        strExpression = "[flowlen] / 0.3048"
    
        Set pFlowLengthFt = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "flowlen"
    Else
        GoTo ProgCancel
    End If
    'END STEP 11: -----------------------------------------------------------------------

    'STEP 12: ---------------------------------------------------------------------------
    'Calculate the slope length exponent value 'M'
    modProgDialog.ProgDialog "Calculating slope length exponent value 'M'...", strProgTitle, _
    0, 15, 12, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        pMapAlgebraOp.BindRaster pDownAngle, "dwnslangle"
        strExpression = "Con(([dwnslangle] le 0.1), 0.01," & _
            "Con(([dwnslangle] gt 0.1 and [dwnslangle] lt 0.2), 0.02," & _
            "Con(([dwnslangle] ge 0.2 and [dwnslangle] lt 0.4), 0.04," & _
            "Con(([dwnslangle] ge 0.4 and [dwnslangle] lt 0.85), 0.08," & _
            "Con(([dwnslangle] ge 0.85 and [dwnslangle] lt 1.4), 0.14," & _
            "Con(([dwnslangle] ge 1.4 and [dwnslangle] lt 2.0), 0.18," & _
            "Con(([dwnslangle] ge 2.0 and [dwnslangle] lt 2.6), 0.22," & _
            "Con(([dwnslangle] ge 2.6 and [dwnslangle] lt 3.1), 0.25," & _
            "Con(([dwnslangle] ge 3.1 and [dwnslangle] lt 3.7), 0.28," & _
            "Con(([dwnslangle] ge 3.7 and [dwnslangle] lt 5.2), 0.32," & _
            "Con(([dwnslangle] ge 5.2 and [dwnslangle] lt 6.3), 0.35," & _
            "Con(([dwnslangle] ge 6.3 and [dwnslangle] lt 7.4), 0.37," & _
            "Con(([dwnslangle] ge 7.4 and [dwnslangle] lt 8.6), 0.40," & _
            "Con(([dwnslangle] ge 8.6 and [dwnslangle] lt 10.3), 0.41," & _
            "Con(([dwnslangle] ge 10.3 and [dwnslangle] lt 12.9), 0.44," & _
            "Con(([dwnslangle] ge 12.9 and [dwnslangle] lt 15.7), 0.47," & _
            "Con(([dwnslangle] ge 15.7 and [dwnslangle] lt 20.0), 0.49," & _
            "Con(([dwnslangle] ge 20.0 and [dwnslangle] lt 25.8), 0.52," & _
            "Con(([dwnslangle] ge 25.8 and [dwnslangle] lt 31.5), 0.54," & _
            "Con(([dwnslangle] ge 31.5 and [dwnslangle] lt 37.2), 0.55," & _
            "0.56))))))))))))))))))))"
    
        Set pSlopeExp = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "dwnslangle"
    Else
        GoTo ProgCancel
    End If
    'END STEP 12: ------------------------------------------------------------------------

    'STEP 13: ----------------------------------------------------------------------------
    'Calculate the L-Factor
    modProgDialog.ProgDialog "Calculating the L factor...", strProgTitle, _
    0, 15, 13, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        With pMapAlgebraOp
            .BindRaster pFlowLengthFt, "flowlenft"
            .BindRaster pSlopeExp, "new_slpexp"
        End With
        'This non-dimensionalizes flowlenft and, after raising the result to the power in new_slpexp, we have the L factor.
        strExpression = "Pow(([flowlenft] / 72.6), [new_slpexp])"
    
        Set pRusleLFactor = pMapAlgebraOp.Execute(strExpression)
        'AddRasterLayer Application, pRusleLFactor, "RussleL"
        
        With pMapAlgebraOp
            .UnbindRaster "flowlenft"
            .UnbindRaster "new_slpexp"
        End With
    Else
        GoTo ProgCancel
    End If
    'END STEP 13: ------------------------------------------------------------------------

    'STEP 14: ----------------------------------------------------------------------------
    'Calculate the S-Factor
    modProgDialog.ProgDialog "Creating flow length GRID...", strProgTitle, _
    0, 15, 14, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        'Here is the calculation for S, which is not, actually the slope,
        'but IS a function of the slope between a cell and its immediate downstream neighbor.
        pMapAlgebraOp.BindRaster pDownAngle, "dwnslangle"
        strExpression = "Con([dwnslangle] ge 5.1428, 16.8 * (sin(([dwnslangle] - 0.5) div deg))," & _
        "10.8 * (sin(([dwnslangle] + 0.03) div deg)))"
    
        Set pRusleSFactor = pMapAlgebraOp.Execute(strExpression)
        
        pMapAlgebraOp.UnbindRaster "dwnslangle"
    Else
        GoTo ProgCancel
    End If
    'END STEP 14: -------------------------------------------------------------------------
    
    'STEP 15: ----------------------------------------------------------------------------
    'Calculate the LS Factor
    modProgDialog.ProgDialog "Calculating the LS Factor...", strProgTitle, _
    0, 15, 15, m_App.hwnd

    If modProgDialog.g_boolCancel Then
    
        'quick math to clip this bugger
        With pMapAlgebraOp
            .BindRaster pRusleSFactor, "Sfactor"
            .BindRaster pRusleLFactor, "Lfactor"
        End With
        
        strExpression = "[Sfactor] * [Lfactor]"
        
        Set pLSFactor = pMapAlgebraOp.Execute(strExpression)
        
        With pMapAlgebraOp
            .UnbindRaster "Sfactor"
            .UnbindRaster "Lfactor"
        End With
    'END STEP 15: -------------------------------------------------------------------------
    
    'STEP 15a: ----------------------------------------------------------------------------
        With pMapAlgebraOp
            .BindRaster pLSFactor, "LSFactor"
            .BindRaster pMask, "Mask"
        End With
        
        strExpression = "[LSFactor] * [Mask]"
    
        Set pFinalLS = pMapAlgebraOp.Execute(strExpression)
    
        With pMapAlgebraOp
            .UnbindRaster "LSFactor"
            .UnbindRaster "Mask"
        End With
    Else
        GoTo ProgCancel
    End If
    
    CalcLengthSlope = modUtil.MakePerminentGrid(pLSFactor, pWS.PathName, "LSGrid")
    
    modProgDialog.KillDialog
    
 'Cleanup
    Set pDEMOneCell = Nothing              'STEP 1: 1 Cell Buffer DEM
    Set pDEMTwoCell = Nothing              'STEP 2: 2 Cell Buffer DEM
    Set pFlowDir = Nothing                 'STEP 3: Flow Direction
    Set pFlowDirBV = Nothing               'STEP 4: Pre Nibble Null
    Set pMask = Nothing                    'STEP 4: Mask
    Set pNibble = Nothing                  'STEP 4: Nibble
    Set pDownSlope = Nothing               'STEP 5: Down Slope
    Set pDownAngle = Nothing               'STEP 5a: Tweak down slope
    Set pRelativeSlope = Nothing           'STEP 6: Relative Slope
    Set pRelSlopeThreshold = Nothing       'STEP 7: Relative slope threshold
    Set pSlopeBreak = Nothing              'STEP 7a: Slope Break
    Set pFlowDirBreak = Nothing            'STEP 8: Flow Direction Break
    Set pWeight = Nothing                  'STEP 9: Weight GRID
    Set pFlowLength = Nothing              'STEP 10: Flow Length
    Set pFlowLengthFt = Nothing            'STEP 11: Flow Length to Feet
    Set pSlopeExp = Nothing                'STEP 12: Slope Exponent
    Set pRusleLFactor = Nothing            'STEP 13: Rusle L Factor
    Set pRusleSFactor = Nothing            'STEP 14: Rusle S Factor
    Set pFinalLS = Nothing                 'STEP 15: Final LS Factor
    Set pLSFactor = Nothing                'STEP 15a: Masked LS Factor
    
Exit Function
ProgCancel:
    modProgDialog.KillDialog
    MsgBox "The LS GRID calculation has been stopped by the user.  Changes have been discarded.", vbCritical, "Process Stopped"
    CalcLengthSlope = "ERROR"

Exit Function
ErrHandler:
    MsgBox Err.Number & ": " & Err.Description

End Function






