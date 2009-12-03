VERSION 5.00
Begin VB.Form frmNewPrecip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Precipitation Scenario"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7155
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleMode       =   0  'User
   ScaleWidth      =   7213.289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Enter new scenario information  "
      Height          =   3600
      Left            =   120
      TabIndex        =   8
      Top             =   105
      Width           =   6915
      Begin VB.TextBox txtRainingDays 
         Height          =   286
         Left            =   5055
         TabIndex        =   19
         Top             =   2670
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.ComboBox cboTimePeriod 
         Height          =   315
         ItemData        =   "frmNewPrecip.frx":0000
         Left            =   1680
         List            =   "frmNewPrecip.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2655
         Width           =   2145
      End
      Begin VB.ComboBox cboPrecipType 
         Height          =   315
         ItemData        =   "frmNewPrecip.frx":001D
         Left            =   1680
         List            =   "frmNewPrecip.frx":002D
         TabIndex        =   14
         Top             =   3135
         Width           =   2145
      End
      Begin VB.TextBox txtPrecipName 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   375
         Width           =   1635
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   810
         Width           =   4035
      End
      Begin VB.TextBox txtPrecipFile 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   1230
         Width           =   4035
      End
      Begin VB.CommandButton cmdBrowseFile 
         Height          =   315
         Left            =   5775
         Picture         =   "frmNewPrecip.frx":0055
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1215
         Width           =   375
      End
      Begin VB.ComboBox cboPrecipUnits 
         Height          =   315
         ItemData        =   "frmNewPrecip.frx":08CB
         Left            =   1680
         List            =   "frmNewPrecip.frx":08D5
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Width           =   2145
      End
      Begin VB.ComboBox cboGridUnits 
         Height          =   315
         ItemData        =   "frmNewPrecip.frx":08EE
         Left            =   1680
         List            =   "frmNewPrecip.frx":08F8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   2145
      End
      Begin VB.Label lblRainingDays 
         Caption         =   "Raining Days: "
         Height          =   315
         Left            =   3975
         TabIndex        =   18
         Top             =   2700
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Period:"
         Height          =   255
         Left            =   315
         TabIndex        =   17
         Top             =   2685
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Type:"
         Height          =   255
         Left            =   315
         TabIndex        =   15
         Top             =   3150
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Scenario Name:"
         Height          =   255
         Index           =   7
         Left            =   345
         TabIndex        =   13
         Top             =   375
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Description:"
         Height          =   285
         Index           =   1
         Left            =   75
         TabIndex        =   12
         Top             =   810
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Precipitation Grid:"
         Height          =   285
         Index           =   0
         Left            =   75
         TabIndex        =   11
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Precipitation Units:"
         Height          =   285
         Index           =   2
         Left            =   75
         TabIndex        =   10
         Top             =   2145
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Units:"
         Height          =   285
         Index           =   6
         Left            =   75
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5940
      TabIndex        =   7
      Top             =   3855
      Width           =   967
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4890
      TabIndex        =   6
      Top             =   3855
      Width           =   967
   End
End
Attribute VB_Name = "frmNewPrecip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmNewPrecip
' *************************************************************************************
' *  Description: Form for entering a new precipitation scenarios
' *  within NSPECT
' *
' *  Called By:  frmPrecip
' *************************************************************************************

Option Explicit

Public m_pInputPrecipDS As IRasterDataset



Private Sub cboTimePeriod_Click()
    
    If cboTimePeriod.ListIndex = 0 Then
        lblRainingDays.Visible = True
        txtRainingDays.Visible = True
    Else
        lblRainingDays.Visible = False
        txtRainingDays.Visible = False
    End If
    
End Sub

Private Sub cmdBrowseFile_Click()
    
    Dim pPrecipRasterDataset As IRasterDataset
    Dim pDistUnit As ILinearUnit
    Dim intUnit As Integer
    Dim pProjCoord As IProjectedCoordinateSystem
        
On Error GoTo ErrHandler:

    Set m_pInputPrecipDS = AddInputFromGxBrowserText(txtPrecipFile, "Choose Precipitation GRID", frmPrecip, 0)
        
    If Not m_pInputPrecipDS Is Nothing Then
    
        Set pPrecipRasterDataset = m_pInputPrecipDS
    
        If CheckSpatialReference(pPrecipRasterDataset) Is Nothing Then
            
            MsgBox "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", vbExclamation, "No Project Information Detected"
            Exit Sub
        
        Else
        
            Set pProjCoord = CheckSpatialReference(pPrecipRasterDataset)
            Set pDistUnit = pProjCoord.CoordinateUnit
            intUnit = pDistUnit.MetersPerUnit
            
            If intUnit = 1 Then
                cboGridUnits.ListIndex = 0
            Else
                cboGridUnits.ListIndex = 1
            End If
            
            cboGridUnits.Refresh
        End If
    End If
    
    Set pPrecipRasterDataset = Nothing
    Set pDistUnit = Nothing
    Set pProjCoord = Nothing
    
Exit Sub
ErrHandler:
Exit Sub
    
    
End Sub

Private Sub cmdCancel_Click()
              
    Dim intvbYesNo As Integer
    
    intvbYesNo = MsgBox("Are you sure you want to exit?  All changes not saved will be lost.", vbYesNo, "Exit?")
    
    If intvbYesNo = vbYes Then
        If IsLoaded("frmPrj") Then
            frmPrj.cboPrecipScen.ListIndex = 0
        End If
    Else
        Exit Sub
    End If
    
   Unload Me

End Sub

Private Sub cmdOK_Click()
    
    Dim intType As Integer
    Dim intRainingDays As Integer
    
    
    If CheckParams Then
        'Process the time period
        intType = cboTimePeriod.ListIndex
        If intType = 0 Then
            intRainingDays = CInt(txtRainingDays.Text)
        Else
            intRainingDays = 0
        End If
        
        Dim strCmdInsert As String
    
        'Compose the INSERT statement.
        strCmdInsert = "INSERT INTO PrecipScenario " & _
            "(Name, Description, PrecipFileName, PrecipGridUnits, PrecipUnits, Type, PrecipType, RainingDays) VALUES (" & _
            "'" & Replace(CStr(txtPrecipName.Text), "'", "''") & "', " & _
            "'" & Replace(CStr(txtDesc.Text), "'", "''") & "', " & _
            "'" & Replace(txtPrecipFile.Text, "'", "''") & "', " & _
            "" & cboGridUnits.ListIndex & ", " & _
            "" & cboPrecipUnits.ListIndex & ", " & _
            "" & intType & ", " & _
            "" & cboPrecipType.ListIndex & ", " & _
            "" & intRainingDays & _
            ")"

       
        Debug.Print strCmdInsert
        
        If modUtil.UniqueName("PrecipScenario", txtPrecipName.Text) Then
            'Execute the statement.
            modUtil.g_ADOConn.Execute strCmdInsert, adCmdText
            'Confirm
            MsgBox txtPrecipName.Text & " successfully added.", vbOKOnly, "Record Added"
            
            If IsLoaded("frmPrj") Then
                frmPrj.cboPrecipScen.Clear
                modUtil.InitComboBox frmPrj.cboPrecipScen, "PrecipScenario"
                frmPrj.cboPrecipScen.AddItem "New precipitation scenario...", frmPrj.cboPrecipScen.ListCount
                frmPrj.cboPrecipScen.ListIndex = modUtil.GetCboIndex(txtPrecipName.Text, frmPrj.cboPrecipScen)
                Unload Me
            End If
            
            If IsLoaded("frmPrecip") Then
                
                frmPrecip.cboScenName.Clear
                modUtil.InitComboBox frmPrecip.cboScenName, "PrecipScenario"
                frmPrecip.cboScenName.ListIndex = modUtil.GetCboIndex(txtPrecipName.Text, frmPrecip.cboScenName)
                Unload Me
        
            End If
            
        Else
                MsgBox "Name already in use.  Please choose a different one.", vbCritical, "Name In Use"
                txtPrecipName.SetFocus
                Exit Sub
        End If
            
        
    End If
    
    
End Sub




Private Sub txtDesc_Validate(Cancel As Boolean)
    
    txtDesc.Text = Replace(txtDesc.Text, "'", "")
    
End Sub

Private Sub txtPrecipName_Change()

    txtPrecipName.Text = Replace(txtPrecipName.Text, "'", "")
    
End Sub

Private Function CheckParams() As Boolean

    'Check the inputs of the form, before saving
    If Len(txtDesc.Text) = 0 Then
        MsgBox "Please enter a description for this scenario", vbCritical, "Description Missing"
        txtDesc.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If txtPrecipFile.Text = " " Or txtPrecipFile.Text = "" Then
        MsgBox "Please select a valid precipitation GRID before saving.", vbCritical, "GRID Missing"
        txtPrecipFile.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If cboGridUnits.Text = "" Then
        MsgBox "Please select GRID units.", vbCritical, "Units Missing"
        cboGridUnits.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If cboPrecipUnits.Text = "" Then
        MsgBox "Please select precipitation units.", vbCritical, "Units Missing"
        cboPrecipUnits.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If Len(cboPrecipType.Text) = 0 Then
        MsgBox "Please select a Precipitation Type.", vbCritical, "Precipitation Type Missing"
        cboPrecipType.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If Len(cboTimePeriod.Text) = 0 Then
        MsgBox "Please select a Time Period.", vbCritical, "Precipitation Time Period Missing"
        cboTimePeriod.SetFocus
        CheckParams = False
        Exit Function
    End If
    
    If cboTimePeriod.ListIndex = 0 Then
        If Not IsNumeric(txtRainingDays.Text) Or Len(txtRainingDays.Text) = 0 Then
            MsgBox "Please enter a numeric value for Raining Days.", vbCritical, "Raining Days Value Incorrect"
            txtRainingDays.SetFocus
            CheckParams = False
            Exit Function
        End If
    End If
    
    'if it got through all that, then set it to true
    CheckParams = True
    

End Function
