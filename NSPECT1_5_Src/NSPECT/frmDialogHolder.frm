VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDialogHolder 
   Caption         =   "Length Slop GRID "
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAML 
      Height          =   330
      Left            =   3690
      Picture         =   "frmDialogHolder.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   375
      Width           =   435
   End
   Begin VB.TextBox txtAML 
      Height          =   300
      Left            =   840
      TabIndex        =   19
      Top             =   390
      Width           =   2805
   End
   Begin VB.CommandButton cmdWS 
      Height          =   330
      Left            =   3705
      Picture         =   "frmDialogHolder.frx":0876
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2385
      Width           =   435
   End
   Begin VB.CommandButton cmdDEM 
      Height          =   330
      Left            =   3690
      Picture         =   "frmDialogHolder.frx":10EC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1710
      Width           =   435
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3270
      TabIndex        =   14
      Top             =   5205
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2205
      TabIndex        =   13
      Top             =   5205
      Width           =   975
   End
   Begin VB.TextBox txtSlopeGT5 
      Height          =   285
      Left            =   870
      TabIndex        =   12
      Text            =   "0.5"
      Top             =   4860
      Width           =   615
   End
   Begin VB.TextBox txtSlopeLT5 
      Height          =   285
      Left            =   870
      TabIndex        =   10
      Text            =   "0.7"
      Top             =   4215
      Width           =   615
   End
   Begin VB.OptionButton optUnits 
      Caption         =   "Feet"
      Height          =   240
      Index           =   1
      Left            =   3210
      TabIndex        =   8
      Top             =   3525
      Width           =   1095
   End
   Begin VB.OptionButton optUnits 
      Caption         =   "Meters"
      Height          =   240
      Index           =   0
      Left            =   2205
      TabIndex        =   7
      Top             =   3510
      Width           =   960
   End
   Begin VB.TextBox txtWatershed 
      Height          =   285
      Left            =   855
      TabIndex        =   4
      Top             =   2415
      Width           =   2805
   End
   Begin VB.TextBox txtDEM 
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   1725
      Width           =   2805
   End
   Begin VB.TextBox txtPrefix 
      Height          =   285
      Left            =   855
      TabIndex        =   0
      Top             =   1020
      Width           =   900
   End
   Begin MSComDlg.CommonDialog dlgAMLOpen 
      Left            =   4770
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Choose location of MUSLE AML:"
      Height          =   285
      Index           =   1
      Left            =   270
      TabIndex        =   18
      Top             =   135
      Width           =   3315
   End
   Begin VB.Label Label2 
      Caption         =   "* Please Note:  The DEM and watershed boundary grids must be located in the same workspace as the AML."
      Height          =   570
      Left            =   390
      TabIndex        =   15
      Top             =   2835
      Width           =   3885
   End
   Begin VB.Label Label7 
      Caption         =   "Enter slope cutoff factor for slopes >= 5% (suggested = .5):"
      Height          =   255
      Left            =   315
      TabIndex        =   11
      Top             =   4575
      Width           =   4200
   End
   Begin VB.Label Label6 
      Caption         =   "Enter slope cutoff factor for slopes < 5% (suggested = .7):"
      Height          =   255
      Left            =   315
      TabIndex        =   9
      Top             =   3960
      Width           =   4200
   End
   Begin VB.Label Label5 
      Caption         =   "DEM measurement units:"
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   3495
      Width           =   1980
   End
   Begin VB.Label Label4 
      Caption         =   "Choose the watershed boundary grid* :"
      Height          =   225
      Left            =   240
      TabIndex        =   5
      Top             =   2130
      Width           =   3360
   End
   Begin VB.Label Label3 
      Caption         =   "Choose the Input DEM grid* :"
      Height          =   255
      Left            =   255
      TabIndex        =   2
      Top             =   1410
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Enter a 4 character prefix for the study area:"
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   720
      Width           =   3315
   End
End
Attribute VB_Name = "frmDialogHolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_App As IApplication

Private strPrefix As String
Private strWorkspace As String
Private strDEMGrid As String
Private strWSGrid As String
Private strMeterFeet As String
Private strSlopeLT5 As String
Private strSlopeGT5 As String




' Variables used by the Error handler function

Public Sub init(ByVal pApp As IApplication)

    Set m_App = pApp

End Sub

Private Sub cmdAML_Click()
    
    With dlgAMLOpen
        .Filter = "Arc Macro Language File (*.aml)|*.aml"
        .DialogTitle = "Choose the Length/Slope AML"
        .DefaultExt = ".aml"
        .InitDir = App.path
        .ShowOpen
    End With
    
    txtAML.Text = dlgAMLOpen.FileName
    
    
End Sub

Private Sub cmdDEM_Click()
    
    Dim pDEMDS As IRasterDataset
    Dim pDistUnit As ILinearUnit
    Dim intUnit As Integer
    Dim pProjCoord As IProjectedCoordinateSystem
    
On Error GoTo ErrHandler:

    Set pDEMDS = AddInputFromGxBrowserText(txtDEM, "Choose DEM GRID", frmDialogHolder, 0)
    
    If pDEMDS Is Nothing Then
        Exit Sub
    Else
    
        If CheckSpatialReference(pDEMDS) Is Nothing Then
            
            MsgBox "The GRID you have choosen has no spatial reference information.  Please define a projection before continuing.", vbExclamation, "No Project Information Detected"
            txtDEM.Text = ""
            Exit Sub
        
        Else
        
            Set pProjCoord = CheckSpatialReference(pDEMDS)
            Set pDistUnit = pProjCoord.CoordinateUnit
            intUnit = pDistUnit.MetersPerUnit
            
                If intUnit = 1 Then
                    optUnits(0) = 1
                Else
                    optUnits(1) = 0
                End If
            
        End If
    
    End If
    
    
    Set pPrecipRasterDataset = Nothing
    Set pDistUnit = Nothing
    Set pProjCoord = Nothing
    Set pDEMDS = Nothing

Exit Sub
ErrHandler:
    Exit Sub
    
End Sub

Private Sub cmdOK_Click()

    Dim pos As Integer
    
    strPrefix = txtPrefix.Text
    strWorkspace = modUtil.SplitWorkspaceName(txtDEM.Text)
    strDEMGrid = modUtil.SplitFileName(txtDEM.Text)
    strWSGrid = modUtil.SplitFileName(txtWatershed.Text)
    
    If optUnits(0) Then
        strMeterFeet = "meters"
    Else
        strMeterFeet = "feet"
    End If
    
    strSlopeLT5 = txtSlopeLT5.Text
    strSlopeGT5 = txtSlopeGT5.Text
    
    RunAML
    
End Sub

Private Sub cmdWS_Click()

    Dim pWSDS As IRasterDataset
    Set pWSDS = AddInputFromGxBrowserText(txtWatershed, "Choose Watershed Boundary GRID", frmDialogHolder, 0)

End Sub



Private Sub RunAML()

'    Dim arcMod As ESRI.Arc
'    Dim arcResults As ESRIutil.Strings
'    Set arcMod = New ESRI.Arc
'    Set arcResults = New ESRIutil.Strings
'
'    Dim ArcStatus As Long
'    Dim cmdChangeWS As String
'    Dim cmdRunAML As String
'
'    frmDialogHolder.Hide
'    frmLSProcess.Show
'
'    cmdChangeWS = "w " & modUtil.SplitWorkspaceName(frmDialogHolder.dlgAMLOpen.FileName)
'    ArcStatus = arcMod.Command(cmdChangeWS, Nothing)
'
'    If ArcStatus = 0 Or ArcStatus = 1 Then
'
'        With arcMod
'            .PushString strPrefix
'            .PushString strWorkspace
'            .PushString strDEMGrid
'            .PushString strWSGrid
'            .PushString strMeterFeet
'            .PushString strSlopeLT5
'            .PushString strSlopeGT5
'        End With
'
'        cmdRunAML = "&run " & frmDialogHolder.dlgAMLOpen.FileName
'        ArcStatus = arcMod.Command(cmdRunAML, arcResults)
'
'    Else
'        MsgBox "Can not get to workspace.", vbCritical, "Can't Find Workspace"
'        Exit Sub
'    End If
    

    
End Sub



