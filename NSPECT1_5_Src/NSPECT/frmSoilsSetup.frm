VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSoilsSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Soils Setup"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Advanced MUSLE Specific Coefficients"
      Height          =   2355
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   4935
      Begin VB.TextBox txtMUSLEVal 
         Height          =   330
         Left            =   840
         TabIndex        =   17
         Text            =   "95"
         Top             =   840
         Width           =   645
      End
      Begin VB.TextBox txtMUSLEExp 
         Height          =   330
         Left            =   2040
         TabIndex        =   16
         Text            =   "0.56"
         Top             =   840
         Width           =   525
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSoilsSetup.frx":0000
         Height          =   855
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "b="
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "a="
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "    b"
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "MUSLE Equation for sediment yield:"
         Height          =   240
         Left            =   270
         TabIndex        =   20
         Top             =   285
         Width           =   3780
      End
      Begin VB.Label Label9 
         Caption         =   "a * (Q * qp)   * K * C * P * LS"
         Height          =   300
         Left            =   480
         TabIndex        =   19
         Top             =   570
         Width           =   3180
      End
      Begin VB.Label Label8 
         Caption         =   "Locally calibrated MUSLE coefficients can be entered above."
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   4545
      End
   End
   Begin VB.TextBox txtSoilsName 
      Height          =   315
      Left            =   1335
      TabIndex        =   0
      ToolTipText     =   "Choose the filled DEM you are using.  This will provide Spatial Analyst with the proper analysis environment."
      Top             =   90
      Width           =   3180
   End
   Begin VB.CommandButton cmdDEMBrowse 
      Height          =   315
      Left            =   4605
      Picture         =   "frmSoilsSetup.frx":0092
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   585
      Width           =   375
   End
   Begin VB.TextBox txtDEMFile 
      Height          =   315
      Left            =   1335
      TabIndex        =   1
      ToolTipText     =   "Choose the filled DEM you are using.  This will provide Spatial Analyst with the proper analysis environment."
      Top             =   570
      Width           =   3180
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4110
      TabIndex        =   8
      Top             =   5370
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "OK"
      Height          =   375
      Left            =   3060
      TabIndex        =   7
      Top             =   5370
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   " Soils  "
      Height          =   1800
      Left            =   240
      TabIndex        =   9
      Top             =   1035
      Width           =   4920
      Begin VB.ComboBox cboSoilFieldsK 
         Height          =   315
         Left            =   2595
         TabIndex        =   6
         Top             =   1245
         Width           =   2145
      End
      Begin VB.ComboBox cboSoilFields 
         Height          =   315
         Left            =   2595
         TabIndex        =   5
         Top             =   735
         Width           =   2145
      End
      Begin VB.CommandButton cmdBrowseFile 
         Height          =   315
         Left            =   4365
         Picture         =   "frmSoilsSetup.frx":0908
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox txtSoilsDS 
         Height          =   300
         Left            =   1395
         TabIndex        =   3
         Top             =   285
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "K Factor Attribute:"
         Height          =   285
         Left            =   195
         TabIndex        =   13
         Top             =   1275
         Width           =   1830
      End
      Begin VB.Label Label3 
         Caption         =   "Hydrologic Soil Group Attribute:"
         Height          =   450
         Left            =   165
         TabIndex        =   11
         Top             =   780
         Width           =   2385
      End
      Begin VB.Label Label1 
         Caption         =   "Soils Data Set:"
         Height          =   270
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Width           =   1230
      End
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   75
      Top             =   5340
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Name:"
      Height          =   270
      Left            =   270
      TabIndex        =   14
      Top             =   135
      Width           =   885
   End
   Begin VB.Label Label5 
      Caption         =   "DEM GRID:"
      Height          =   285
      Left            =   255
      TabIndex        =   12
      Top             =   615
      Width           =   1155
   End
End
Attribute VB_Name = "frmSoilsSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmSoilsSetup
' *************************************************************************************
' *  Description: Form for allowing the user to create soils and k-factor soils grids
' *
' *
' *  Called By:  frmSoils
' *************************************************************************************

Option Explicit
Private m_App As IApplication                   'Application handle
Private m_pRasterProps As IRasterProps          'Raster props garnered from DEM

Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub

Private Sub cmdBrowseFile_Click()

On Error GoTo ErrHandler:

   'browse...get output filename
   dlgOpen.FileName = Empty
   With dlgOpen
     .Filter = MSG6
     .DialogTitle = "Open Soils Dataset"
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .CancelError = True
     .ShowOpen
   End With
   
   If Len(dlgOpen.FileName) > 0 Then
      txtSoilsDS.Text = Trim(dlgOpen.FileName)
      PopulateCbo
   End If
   
ErrHandler:
    Exit Sub

End Sub

Private Sub PopulateCbo()

'Populate cboSoilFields & cboSoilFieldsK with the fields in the selected Soils layer
    Dim i As Integer
    Dim pFields As IFields
    Dim pFeatureClass As IFeatureClass
    
    cboSoilFields.Clear
    cboSoilFieldsK.Clear
        
    Set pFeatureClass = modUtil.ReturnFeatureClass(txtSoilsDS.Text)
    Set pFields = pFeatureClass.Fields
    
    'Pop both cbos with field names
    For i = 1 To pFields.FieldCount - 1
        cboSoilFields.AddItem pFields.Field(i).Name
        cboSoilFieldsK.AddItem pFields.Field(i).Name
    Next i
    
    'Cleanup
    Set pFields = Nothing
    Set pFeatureClass = Nothing
    
End Sub

Private Sub cmdQuit_Click()
    
    Dim intvbYesNo As Integer
    
    intvbYesNo = MsgBox("Do you want to save changes you made to soils setup?", vbYesNoCancel + vbExclamation, "N-SPECT")
    
    If intvbYesNo = vbYes Then
        Call SaveSoils
    ElseIf intvbYesNo = vbNo Then
        Unload frmSoilsSetup
    ElseIf intvbYesNo = vbCancel Then
        Exit Sub
    End If
    
    
End Sub

Private Sub cmdSave_Click()
    
    Call SaveSoils
      
End Sub

Private Sub SaveSoils()
    
    'Check data, if OK create soils grids
    If ValidateData Then
        If CreateSoilsGrid(txtSoilsDS.Text, cboSoilFields.Text, cboSoilFieldsK.Text) Then
            If frmSoils.Visible Then
                frmSoils.cboSoils.Clear
                modUtil.InitComboBox frmSoils.cboSoils, "Soils"
                frmSoils.cboSoils.ListIndex = modUtil.GetCboIndex(txtSoilsName.Text, frmSoils.cboSoils)
                Unload frmSoilsSetup
            End If
        Else
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    

End Sub


Private Function ValidateData() As Boolean

    If Len(txtSoilsName.Text) > 0 Then
        If modUtil.UniqueName("Soils", txtSoilsName.Text) Then
            ValidateData = True
        Else
            MsgBox "The name you have chosen is already in use.  Please select another.", vbCritical, "Select Unique Name"
            ValidateData = False
            txtSoilsName.SetFocus
            Exit Function
        End If
    Else
        MsgBox "Please enter a name.", vbCritical, "Soils Name Missing"
        ValidateData = False
        txtSoilsName.SetFocus
        Exit Function
        
    End If

    If Len(txtSoilsDS.Text) = 0 Then
        MsgBox "Please select a soils dataset.", vbCritical, "Soils Dataset Missing"
        txtSoilsDS.SetFocus
        ValidateData = False
        Exit Function
    Else
        ValidateData = True
    End If
    
    If Len(cboSoilFields.Text) = 0 Then
        MsgBox "Please select a soils attribute.", vbCritical, "Choose Soils Attribute"
        cboSoilFields.SetFocus
        ValidateData = False
        Exit Function
    Else
        ValidateData = True
    End If
    
    If Len(cboSoilFieldsK.Text) = 0 Then
        MsgBox "Please select a k-factor soils attribute.", vbCritical, "Choose K-Factor Attribute"
        cboSoilFieldsK.SetFocus
        ValidateData = False
        Exit Function
    Else
        ValidateData = True
    End If
    
    If Len(txtMUSLEVal.Text) > 0 Then
        If IsNumeric(CDbl(txtMUSLEVal.Text)) Then
            ValidateData = True
        Else
            MsgBox "Please enter a numeric value for the MUSLE equation.", vbCritical, "Numeric Values Only"
            ValidateData = False
        End If
    Else
        MsgBox "Please enter a value for the MUSLE equation.", vbCritical, "Missing Value"
        txtMUSLEVal.SetFocus
        ValidateData = False
    End If
    
    If Len(txtMUSLEExp.Text) > 0 Then
        If IsNumeric(CDbl(txtMUSLEExp.Text)) Then
            ValidateData = True
        Else
            MsgBox "Please enter a numeric value for the MUSLE equation.", vbCritical, "Numeric Values Only"
            ValidateData = False
        End If
    Else
        MsgBox "Please enter a value for the MUSLE equation.", vbCritical, "Missing Value"
        txtMUSLEExp.SetFocus
        ValidateData = False
    End If
    
    
End Function

Private Function CreateSoilsGrid(strSoilsFileName As String, strHydFieldName As String, Optional strKFactor As String) As Boolean
'Incoming:
    'strSoilsFileName: string of soils file name path
    'strHydFieldName: string of hydrologic group attribute
    'strKFactor: string of K factor attribute
    
On Error GoTo ErrHandler:
    
    Dim pSoilsFeatClass As IFeatureClass                        'Soils Featureclass
    Dim pSoilsFeatCursor As IFeatureCursor                      'Cursor to loop through soils
    Dim pSoilsFeature As IFeature                               'Feature for use
    Dim lngHydFieldIndex As Long                                'HydGroup Field Index
    Dim lngNewHydFieldIndex As Long                             'New Group Field index
    Dim pFieldEdit As IFieldEdit                                'FieldEdit, case we have to add one
    Dim pField As IField                                        'Field
    Dim strHydValue As String                                   'Hyd Value to check
    Dim lngValue As Long                                        'Count
    Dim strSoilsRas As String                                   'Soils Raster Name
    Dim strSoilsKRas As String                                  'Soils K raster Name
    Dim strCmd As String                                        'String to insert new stuff in dbase
    Dim strOutSoils As String                                   'OutSoils name
    Dim strOutKSoils As String
    
    Dim pConversionOp As IConversionOp                          'Feat to Grid convert
    Dim pConversionOpK As IConversionOp                         'K Feat to Grid convert
    Dim pSoilsRaster As IRasterDataset                          'Soils raster ds
    Dim pSoilsKRaster As IRasterDataset                         'Soils k factor ds
    Dim pSoilsDS As IGeoDataset
    Dim pSoilsKDS As IGeoDataset
    Dim pSoilsFeatureDescriptor As IFeatureClassDescriptor      'Descript for soils
    Dim pSoilsKFeatureDescriptor As IFeatureClassDescriptor     'Descript for soils K
    
    Dim pEnv As IRasterAnalysisEnvironment
    Dim pWSFact As IWorkspaceFactory
    Dim pWS As IRasterWorkspace

    'Get the soils featurclass
    Set pSoilsFeatClass = modUtil.ReturnFeatureClass(strSoilsFileName)
    
    'Check for fields
    lngHydFieldIndex = pSoilsFeatClass.FindField(strHydFieldName)
    lngNewHydFieldIndex = pSoilsFeatClass.FindField("GROUP")
    
    'If the GROUP field is missing, we have to add it
    If lngNewHydFieldIndex = -1 Then
        Set pFieldEdit = New esriGeoDatabase.Field
        With pFieldEdit
            .Name = "GROUP"
            .Type = esriFieldTypeInteger
            .length = 2
        End With
        Set pField = pFieldEdit

        pSoilsFeatClass.AddField pField

        lngNewHydFieldIndex = pSoilsFeatClass.FindField("GROUP")
    End If

    lngValue = 1
    
    'Get all features in a cursor
    Set pSoilsFeatCursor = pSoilsFeatClass.Update(Nothing, False)
    
    'Get all the features into the cursor
    Set pSoilsFeature = pSoilsFeatCursor.NextFeature

    'Now calc the Values
    Do While Not pSoilsFeature Is Nothing
        modProgDialog.ProgDialog "Calculating soils values...", "Processing Soils", 0, pSoilsFeatClass.FeatureCount(Nothing), lngValue, m_App.hwnd
        'Find the current value
        If modProgDialog.g_boolCancel Then
            strHydValue = pSoilsFeature.Value(lngHydFieldIndex)
            'Based on current value, change GROUP to appropriate setting
            Select Case strHydValue
                Case "A"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 1
                Case "B"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 2
                Case "C"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 3
                Case "D"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 4
                Case "A/B"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 2
                Case "B/C"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 3
                Case "C/D"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 4
                Case "B/D"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 4
                Case "A/D"
                    pSoilsFeature.Value(lngNewHydFieldIndex) = 4
                Case ""
                    MsgBox "Your soils dataset contains missing values for Hydrologic Soils Attribute.  Please correct.", vbCritical, "Missing Values Detected"
                    CreateSoilsGrid = False
                    modProgDialog.KillDialog
                    Exit Function
            End Select
            'Update row and move to next
            pSoilsFeatCursor.UpdateFeature pSoilsFeature
            Set pSoilsFeature = pSoilsFeatCursor.NextFeature
            lngValue = lngValue + 1
        Else
            'If they cancel, kill the dialog
            modProgDialog.KillDialog
            Exit Function
        End If
    Loop
    
    'Close dialog
    modProgDialog.KillDialog
    
    'STEP 2:
    'Now do the conversion: Convert soils layer to GRID using new
    'Group field as the value
    'First set the descriptor to use the 'Group' field
    Set pSoilsFeatureDescriptor = New FeatureClassDescriptor
    pSoilsFeatureDescriptor.Create pSoilsFeatClass, Nothing, "GROUP"
    
    Set pSoilsDS = pSoilsFeatureDescriptor
    
    'Conversion Operation
    Set pConversionOp = New RasterConversionOp
    
    'Set the spat Environment
    Set pEnv = pConversionOp
    
    'Get the goodies from the rasterprops
    With pEnv
        .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
        .SetExtent esriRasterEnvValue, m_pRasterProps.Extent, True
        Set .OutSpatialReference = m_pRasterProps.SpatialReference
    End With
    
    'Set the workspace
    Set pWSFact = New RasterWorkspaceFactory
    Set pWS = pWSFact.OpenFromFile(modUtil.SplitWorkspaceName(strSoilsFileName), m_App.hwnd)
    
    modProgDialog.ProgDialog "Converting Soils Dataset...", "Processing Soils", 0, 2, 1, m_App.hwnd
    
    If modProgDialog.g_boolCancel Then
        strOutSoils = modUtil.GetUniqueName("soils", modUtil.SplitWorkspaceName(strSoilsFileName))
        Set pSoilsRaster = New RasterDataset
        Set pSoilsRaster = pConversionOp.ToRasterDataset(pSoilsDS, "GRID", pWS, strOutSoils)
        
        strSoilsRas = pSoilsRaster.CompleteName
        
    'STEP 3:
    'Now do the conversion: Convert soils layer to GRID using
    'k factor field as the value
        'If they are doing a K factor then repeat the process, this time using 'k' field
        If Len(strKFactor) > 0 Then
            
            Set pSoilsKFeatureDescriptor = New FeatureClassDescriptor
            pSoilsKFeatureDescriptor.Create pSoilsFeatClass, Nothing, strKFactor
            
            Set pSoilsKDS = pSoilsKFeatureDescriptor
            Set pConversionOpK = New RasterConversionOp
            Set pEnv = pConversionOpK
            
            With pEnv
                .SetCellSize esriRasterEnvValue, m_pRasterProps.MeanCellSize.X
                .SetExtent esriRasterEnvValue, m_pRasterProps.Extent, True
                Set .OutSpatialReference = m_pRasterProps.SpatialReference
            End With
            
            modProgDialog.ProgDialog "Converting Soils K Dataset...", "Processing Soils", 0, 2, 2, m_App.hwnd
        
            strOutKSoils = modUtil.GetUniqueName("soilsk", modUtil.SplitWorkspaceName(strSoilsFileName))
            Set pSoilsKRaster = New RasterDataset
            Set pSoilsKRaster = pConversionOpK.ToRasterDataset(pSoilsKDS, "GRID", pWS, strOutKSoils)
        
            strSoilsKRas = pSoilsKRaster.CompleteName
    
    Else
        modProgDialog.KillDialog
    End If
    
    'STEP 4:
    'Now enter all into database
    strCmd = "INSERT INTO SOILS (NAME,SOILSFILENAME,SOILSKFILENAME,MUSLEVal,MUSLEExp) VALUES ('" & _
            Replace(txtSoilsName.Text, "'", "''") & "', '" & _
            Replace(strSoilsRas, "'", "''") & "', '" & _
            Replace(strSoilsKRas, "'", "''") & "', " & _
            CDbl(txtMUSLEVal.Text) & ", " & _
            CDbl(txtMUSLEExp.Text) & ")"
    
    g_ADOConn.Execute strCmd, adCmdText
       
    modProgDialog.KillDialog
    
    Else
        modProgDialog.KillDialog
    End If
    
    CreateSoilsGrid = True
    
    'Cleanup
    Set pSoilsFeatClass = Nothing
    Set pSoilsFeatCursor = Nothing
    Set pSoilsFeature = Nothing
    Set pFieldEdit = Nothing
    Set pField = Nothing
    Set pConversionOp = Nothing
    Set pConversionOpK = Nothing
    Set pSoilsRaster = Nothing
    Set pSoilsKRaster = Nothing
    Set pSoilsDS = Nothing
    Set pSoilsKDS = Nothing
    Set pSoilsFeatureDescriptor = Nothing
    Set pSoilsKFeatureDescriptor = Nothing
    Set pEnv = Nothing
    
Exit Function

ErrHandler:
    MsgBox Err.Number & ": " & Err.Description
    CreateSoilsGrid = False
End Function

Private Sub cmdDEMBrowse_Click()
    'Browse for DEM
    Dim pDEMRasterDataset As IRasterDataset
    Set pDEMRasterDataset = AddInputFromGxBrowserText(txtDEMFile, "Choose DEM Dataset", frmSoilsSetup, 0)
    
    If Not pDEMRasterDataset Is Nothing Then
        Set m_pRasterProps = pDEMRasterDataset.CreateDefaultRaster
    Else
        MsgBox "The Raster Dataset you have chosen is invalid.", vbCritical, "DEM Error"
        Exit Sub
    End If
    
    Set pDEMRasterDataset = Nothing
    
End Sub


