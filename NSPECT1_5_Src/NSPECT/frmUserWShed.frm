VERSION 5.00
Begin VB.Form frmUserWShed 
   Caption         =   "New Watershed Delineation"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6045
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "OK"
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdBrowseLS 
      Height          =   315
      Left            =   6480
      Picture         =   "frmUserWShed.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "Define Watershed Delineation"
      Height          =   3210
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6900
      Begin VB.ComboBox cboDEMUnits 
         Height          =   315
         ItemData        =   "frmUserWShed.frx":0876
         Left            =   2550
         List            =   "frmUserWShed.frx":0880
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1150
         Width           =   2295
      End
      Begin VB.TextBox txtFlowDir 
         Height          =   285
         Left            =   2550
         TabIndex        =   19
         Top             =   1530
         Width           =   3675
      End
      Begin VB.CommandButton cmdBrowseFlowDir 
         Height          =   315
         Left            =   6240
         Picture         =   "frmUserWShed.frx":0892
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1530
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseWS 
         Height          =   315
         Left            =   6240
         Picture         =   "frmUserWShed.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtWaterSheds 
         Height          =   285
         Left            =   2550
         TabIndex        =   15
         Top             =   2760
         Width           =   3675
      End
      Begin VB.CommandButton cmdBrowseFlowAcc 
         Height          =   315
         Left            =   6240
         Picture         =   "frmUserWShed.frx":197E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1950
         Width           =   375
      End
      Begin VB.TextBox txtFlowAcc 
         Height          =   285
         Left            =   2550
         TabIndex        =   10
         Top             =   1950
         Width           =   3675
      End
      Begin VB.CommandButton cmdBrowseDEMFile 
         Height          =   315
         Left            =   6240
         Picture         =   "frmUserWShed.frx":21F4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   735
         Width           =   375
      End
      Begin VB.TextBox txtWSDelinName 
         Height          =   285
         Left            =   2550
         TabIndex        =   0
         Top             =   360
         Width           =   3675
      End
      Begin VB.TextBox txtDEMFile 
         Height          =   285
         Left            =   2550
         TabIndex        =   1
         Top             =   768
         Width           =   3675
      End
      Begin VB.TextBox txtLS 
         Height          =   285
         Left            =   2550
         TabIndex        =   4
         Top             =   2355
         Width           =   3675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DEM Units:"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   21
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Watersheds:"
         Height          =   195
         Index           =   1
         Left            =   1410
         TabIndex        =   16
         Top             =   2760
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Length-slope Grid:"
         Height          =   195
         Index           =   5
         Left            =   1020
         TabIndex        =   9
         Top             =   2385
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flow Accumulation Grid:"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   8
         Top             =   1980
         Width           =   1710
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Watershed Delineation Name:"
         Height          =   195
         Index           =   3
         Left            =   255
         TabIndex        =   7
         Top             =   405
         Width           =   2130
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flow Direction Grid:"
         Height          =   195
         Index           =   0
         Left            =   930
         TabIndex        =   6
         Top             =   1605
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DEM Grid:"
         Height          =   255
         Index           =   12
         Left            =   1560
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUserWShed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_App As IApplication
Private m_strDEM2BFileName As String
Private m_strNibbleName As String

Private Sub cmdBrowseDEMFile_Click()
    
    'ReturnGRIDPath txtDEMFile, "Select DEM GRID"
    
    Dim pDEMRasterDataset As IRasterDataset
    Dim pDistUnit As ILinearUnit
    Dim intUnit As Integer
    Dim pProjCoord As IProjectedCoordinateSystem
    Dim strInputDEM As String
    
On Error GoTo ErrHandler:
    
    Set pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM GRID", frmUserWShed, 0)
            
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
    
    Set pDEMRasterDataset = Nothing
    Set pDistUnit = Nothing
    Set pProjCoord = Nothing
Exit Sub

ErrHandler:
    Exit Sub
    

End Sub

Private Sub cmdBrowseFlowAcc_Click()
    
    ReturnGRIDPath txtFlowAcc, "Select Flow Accumulation GRID"
    
End Sub

Private Sub cmdBrowseFlowDir_Click()
    
    ReturnGRIDPath txtFlowDir, "Select Flow Direction GRID"
    
End Sub

Private Sub cmdBrowseLS_Click()

    ReturnGRIDPath txtLS, "Select Length-Slope GRID"

End Sub

Private Sub cmdBrowseWS_Click()

    txtWaterSheds.Text = BrowseForFileName("Feature", frmUserWShed, "Select Watersheds Shapefile")

End Sub


Private Function ValidateDataFormInput() As Boolean
    
    Dim fso As New FileSystemObject
    Dim Folder As Folder
    
    'check name
    If Len(Trim(txtWSDelinName.Text)) = 0 Then
        MsgBox "Please enter a name for your watershed delineation.", vbInformation, "Name Missing"
        txtWSDelinName.SetFocus
        ValidateDataFormInput = False
        Exit Function
    End If
    
    
    If Not fso.FolderExists(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text) Then
        Set Folder = fso.CreateFolder(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text)
    Else
        MsgBox "Name in use.  Please select another.", vbCritical, "Choose New Name"
        txtWSDelinName.SetFocus
        ValidateDataFormInput = False
        Exit Function
    End If
        
    'check dem
    If Len(Trim(txtDEMFile.Text)) = 0 Then
        MsgBox "Please select a DEM for your watershed delineation.", vbInformation, "DEM Missing"
        txtDEMFile.SetFocus
        ValidateDataFormInput = False
        Exit Function
    Else
        If Not (modUtil.RasterExists(txtDEMFile.Text)) Then
            MsgBox "The DEM selected does not appear to be valid.", vbInformation, "Invalid Dataset"
            txtDEMFile.SetFocus
            ValidateDataFormInput = False
            Exit Function
        End If
    End If
    
    'check flowacc
    If Len(Trim(txtFlowAcc.Text)) = 0 Then
        MsgBox "Please select a Flow Accumulation Grid for your watershed delineation.", vbInformation, "Flow Accumulation Grid Missing"
        txtFlowAcc.SetFocus
        ValidateDataFormInput = False
        Exit Function
    Else
        If Not (modUtil.RasterExists(txtFlowAcc.Text)) Then
            MsgBox "The Flow Accumulation file selected does not appear to be valid.", vbInformation, "Invalid Dataset"
            txtFlowAcc.SetFocus
            ValidateDataFormInput = False
            Exit Function
        End If
    End If
    
    'Check flowdir
    If Len(Trim(txtFlowDir.Text)) = 0 Then
        MsgBox "Please select a Flow Direction Grid for your watershed delineation.", vbInformation, "Flow Direction Grid Missing"
        txtFlowDir.SetFocus
        ValidateDataFormInput = False
        Exit Function
    Else
        If Not (modUtil.RasterExists(txtFlowDir.Text)) Then
            MsgBox "The Flow Direction file selected does not appear to be valid.", vbInformation, "Invalid Dataset"
            txtFlowDir.SetFocus
            ValidateDataFormInput = False
            Exit Function
        End If
    End If
    
    'Check LS
    If Len(Trim(txtLS.Text)) = 0 Then
        MsgBox "Please select a Length-slope Grid for your watershed delineation.", vbInformation, "Length Slope Grid Missing"
        txtLS.SetFocus
        ValidateDataFormInput = False
        Exit Function
    Else
        If Not (modUtil.RasterExists(txtLS.Text)) Then
            MsgBox "The Length-slope file selected does not appear to be valid.", vbInformation, "Invalid Dataset"
            txtLS.SetFocus
            ValidateDataFormInput = False
            Exit Function
        End If
    End If
    
    'Check watersheds
    If Len(Trim(txtWaterSheds.Text)) = 0 Then
        MsgBox "Please select a watershed shapefile for your watershed delineation.", vbInformation, "Watershed Shapefile Missing"
        txtWaterSheds.SetFocus
        ValidateDataFormInput = False
        Exit Function
    Else
        If Not (modUtil.FeatureExists(txtWaterSheds.Text)) Then
            MsgBox "The watersheds file selected does not appear to be valid.", vbInformation, "Invalid Dataset"
            txtWaterSheds.SetFocus
            ValidateDataFormInput = False
            Exit Function
        End If
    End If
    
    'if we got through all that, return true.
    
    ValidateDataFormInput = True
    
    
    
End Function
Private Sub cmdCreate_Click()

    Dim strCmdInsert As String
    Dim strDEM2BFileName As String
    Dim strNibbleFileName As String
           
    If Not ValidateDataFormInput Then
        Exit Sub
    End If
    
    modProgDialog.ProgDialog "Validating input...", "Adding New Delineation...", _
       0, 3, 1, m_App.hwnd
    
On Error GoTo ErrHandler:

    modProgDialog.ProgDialog "Creating 2 Cell Buffer and Nibble GRIDs...", "Adding New Delineation...", _
       0, 3, 2, m_App.hwnd
       
    Return2BDEM txtDEMFile.Text, txtFlowDir.Text
    
    modProgDialog.ProgDialog "Updating Database...", "Adding New Delineation...", _
       0, 3, 2, m_App.hwnd
           
    strCmdInsert = "INSERT INTO WSDelineation " & _
                "(Name, DEMFileName, DEMGridUnits, FlowDirFileName, FlowAccumFileName," & _
                "FilledDEMFileName, HydroCorrected, StreamFileName, SubWSSize, WSFileName, LSFileName, NibbleFileName, DEM2bFileName) " & _
                " VALUES (" & _
                "'" & CStr(txtWSDelinName.Text) & "', " & _
                "'" & CStr(txtDEMFile.Text) & "', " & _
                "'" & cboDEMUnits.ListIndex & "', " & _
                "'" & txtFlowDir.Text & "', " & _
                "'" & txtFlowAcc.Text & "', " & _
                "'" & txtDEMFile.Text & "', " & _
                "'" & "0" & "', " & _
                "'" & "" & "', " & _
                "'" & "0" & "', " & _
                "'" & txtWaterSheds.Text & "', " & _
                "'" & txtLS.Text & "', " & _
                "'" & m_strNibbleName & "', " & _
                "'" & m_strDEM2BFileName & "')"
        
    'Execute the statement.
    modUtil.g_ADOConn.Execute strCmdInsert, adCmdText
    Screen.MousePointer = vbNormal
            
    modProgDialog.KillDialog
            
    'Confirm
    MsgBox txtWSDelinName.Text & " successfully added.", vbOKOnly, "Record Added"
    
    If g_boolNewWShed Then
        'frmPrj.Show
        frmPrj.Frame.Visible = True
        frmPrj.cboWSDelin.Clear
        modUtil.InitComboBox frmPrj.cboWSDelin, "WSDelineation"
        frmPrj.cboWSDelin.ListIndex = modUtil.GetCboIndex(txtWSDelinName.Text, frmPrj.cboWSDelin)
        Unload Me
        Unload frmNewWSDelin
    Else
        Unload Me
        Unload frmNewWSDelin
        Unload frmWSDelin
    End If
            
    
Exit Sub
            
ErrHandler:
    MsgBox "An error occurred while processing your Watershed Delineation.", vbCritical, "Error"
    modProgDialog.KillDialog
    
End Sub

Private Sub ReturnGRIDPath(txtBox As TextBox, strTitle As String)

    Dim pDEMRasterDataset As IRasterDataset
    Dim pProjCoord As IProjectedCoordinateSystem
    
On Error GoTo ErrHandler:
    
    Set pDEMRasterDataset = AddInputFromGxBrowserText(txtBox, strTitle, frmUserWShed, 0)
    
        
    'Get the spatial reference
    If CheckSpatialReference(pDEMRasterDataset) Is Nothing Then
        
        MsgBox "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", vbExclamation, "No Project Information Detected"
        Exit Sub
    
    End If
    
    'Get the name
    Set pDEMRasterDataset = Nothing
    
ErrHandler:
    Exit Sub
End Sub

Private Sub Return2BDEM(strDEMFileName As String, strFlowDirFileName As String)

    Dim pMapAlgebraOp As IMapAlgebraOp
    Dim pEnv As IRasterAnalysisEnvironment
    Dim pDEMOneCell As IRaster
    Dim pDEMTwoCell As IRaster
    Dim pDEMRaster As IRaster
    Dim pDEMRasterProps As IRasterProps
    Dim pFlowDir As IRaster
    Dim pFlowDirBV As IRaster
    Dim pNibble As IRaster
    Dim pMask As IRaster
    Dim pRasterWorkspaceFactory As IWorkspaceFactory
    Dim pWorkspace As IWorkspace
    Dim intCellSize As Integer
    Dim pEnvelope As IEnvelope
    
    Set pRasterWorkspaceFactory = New RasterWorkspaceFactory
    Set pWorkspace = pRasterWorkspaceFactory.OpenFromFile(modUtil.g_nspectPath & "\wsdelin\" & txtWSDelinName.Text, 0)
    Set pDEMRaster = modUtil.ReturnRaster(strDEMFileName)
    Set pFlowDir = modUtil.ReturnRaster(strFlowDirFileName)
    Set pDEMRasterProps = pDEMRaster
    
    intCellSize = pDEMRasterProps.MeanCellSize.X
    
    Set pEnvelope = pDEMRasterProps.Extent
    pEnvelope.Expand intCellSize * 2, intCellSize * 2, False
    
    Set pMapAlgebraOp = New RasterMapAlgebraOp
    Set pEnv = pMapAlgebraOp
    
    With pEnv
       .SetCellSize esriRasterEnvValue, pDEMRasterProps.MeanCellSize.X
       Set .OutWorkspace = pWorkspace
       Set .OutSpatialReference = pDEMRasterProps.SpatialReference '.SpatialReference
       .SetExtent esriRasterEnvValue, pEnvelope 'pDEMRasterProps.Extent
    End With
    
    'STEP 1: ----------------------------------------------------------------------
    'Buffer the DEM by one cell
    pMapAlgebraOp.BindRaster pDEMRaster, "aml_fdem"
    strExpression = "Con(isnull([aml_fdem]), focalmin([aml_fdem]), [aml_fdem])"
    
    Set pDEMOneCell = pMapAlgebraOp.Execute(strExpression)
    pMapAlgebraOp.UnbindRaster "aml_fdem"
    'END STEP 1: ------------------------------------------------------------------
    
    'STEP 2: ----------------------------------------------------------------------
    'Buffer the DEM buffer by one more cell
    pMapAlgebraOp.BindRaster pDEMOneCell, "dem_b"
    strExpression = "Con(isnull([dem_b]), focalmin([dem_b]), [dem_b])"
    
    Set pDEMTwoCell = pMapAlgebraOp.Execute(strExpression)
    m_strDEM2BFileName = modUtil.MakePerminentGrid(pDEMTwoCell, pWorkspace.PathName, "dem2b")
    pMapAlgebraOp.UnbindRaster "dem_b"
    
    'STEP 3: ----------------------------------------------------------------------
    pMapAlgebraOp.BindRaster pDEMOneCell, "mask"
    strExpression = "con([mask] >= 0, 1, 0)"
    Set pMask = pMapAlgebraOp.Execute(strExpression)

    With pEnv
        Set .Mask = pMask
        Set .OutWorkspace = pWorkspace
    End With

    Set pEnv = pMapAlgebraOp
        
    'STEP 4: ----------------------------------------------------------------------
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
    m_strNibbleName = modUtil.MakePerminentGrid(pNibble, pWorkspace.PathName, "nibble")
        
    With pMapAlgebraOp
       .UnbindRaster "fdr_bv"
       .UnbindRaster "waia_reg"
    End With
    
    'Cleanup
    Set pMapAlgebraOp = Nothing
    Set pEnv = Nothing
    Set pDEMOneCell = Nothing
    Set pDEMTwoCell = Nothing
    Set pDEMRaster = Nothing
    Set pDEMRasterProps = Nothing
    Set pRasterWorkspaceFactory = Nothing
    Set pWorkspace = Nothing
    
End Sub

Public Sub init(ByVal pApp As IApplication)
  

    Set m_App = pApp
   

End Sub


Private Sub cmdQuit_Click()
    Unload frmUserWShed
End Sub
