VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportCoeffSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Coefficient Set"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLCType 
      Height          =   315
      ItemData        =   "frmImportCoeffSet.frx":0000
      Left            =   2430
      List            =   "frmImportCoeffSet.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   585
      Width           =   2200
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   315
      Left            =   5820
      Picture         =   "frmImportCoeffSet.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1020
      Width           =   375
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5130
      TabIndex        =   5
      Top             =   1500
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4095
      TabIndex        =   4
      Top             =   1500
      Width           =   975
   End
   Begin VB.TextBox txtImpFile 
      Height          =   285
      Left            =   2415
      TabIndex        =   2
      Top             =   1035
      Width           =   3360
   End
   Begin VB.TextBox txtCoeffSetName 
      Height          =   285
      Left            =   2430
      TabIndex        =   0
      Top             =   180
      Width           =   2145
   End
   Begin MSComDlg.CommonDialog dlgCMD1 
      Left            =   234
      Top             =   1521
      _ExtentX        =   767
      _ExtentY        =   767
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Coefficient Set Name:"
      Height          =   255
      Index           =   5
      Left            =   -720
      TabIndex        =   8
      Top             =   225
      Width           =   2835
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Import File:"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   1035
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Land Cover Type:"
      Height          =   255
      Index           =   7
      Left            =   210
      TabIndex        =   6
      Top             =   630
      Width           =   1905
   End
End
Attribute VB_Name = "frmImportCoeffSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  frmImportCoeffSet
' *************************************************************************************
' *  Description:  Allows the user to import a text file containing a coeff set
' *
' *  Called By:  frmWQStd
' *************************************************************************************

Option Explicit
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = "frmImportCoeffSet.frm"


Private Sub cmdBrowse_Click()
  On Error GoTo ErrorHandler

   'browse...get output filename
   dlgCMD1.FileName = Empty
   With dlgCMD1
     .Filter = MSG1
     .DialogTitle = MSG2
     .FilterIndex = 1
     .Flags = FileOpenConstants.cdlOFNHideReadOnly + FileOpenConstants.cdlOFNFileMustExist
     .ShowOpen
   End With
   If Len(dlgCMD1.FileName) > 0 Then
      txtImpFile.Text = Trim(dlgCMD1.FileName)
   End If



  Exit Sub
ErrorHandler:
  HandleError True, "cmdBrowse_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdCancel_Click()
  On Error GoTo ErrorHandler

   Unload Me


  Exit Sub
ErrorHandler:
  HandleError True, "cmdCancel_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

Private Sub cmdOK_Click()
  On Error GoTo ErrorHandler

    
If modUtil.UniqueName("CoefficientSet", txtCoeffSetName.Text) Then
    
    If modUtil.ValidateCoeffTextFile(txtImpFile.Text, cboLCType.Text) Then
        frmPollutants.UpdateCoeffSet g_adoRSCoeff, txtCoeffSetName.Text, txtImpFile.Text
    End If

Else
    MsgBox "The name you have chosen is in use, please enter a different name.", vbCritical, "Name Detected"
    txtCoeffSetName.SetFocus
    
End If



  Exit Sub
ErrorHandler:
  HandleError True, "cmdOK_Click " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
    End Sub

Private Sub Form_Load()
  On Error GoTo ErrorHandler

     InitComboBox cboLCType, "LCTYPE"


  Exit Sub
ErrorHandler:
  HandleError True, "Form_Load " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Sub

