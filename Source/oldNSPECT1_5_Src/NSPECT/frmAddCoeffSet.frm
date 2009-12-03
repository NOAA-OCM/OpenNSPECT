VERSION 5.00
Begin VB.Form frmAddCoeffSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Coefficient Set"
   ClientHeight    =   1725
   ClientLeft      =   7785
   ClientTop       =   3465
   ClientWidth     =   5580
   Icon            =   "frmAddCoeffSet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLCType 
      Height          =   315
      ItemData        =   "frmAddCoeffSet.frx":058A
      Left            =   2310
      List            =   "frmAddCoeffSet.frx":058C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   690
      Width           =   2595
   End
   Begin VB.TextBox txtCoeffSetName 
      Height          =   315
      Left            =   2310
      TabIndex        =   0
      Top             =   225
      Width           =   2625
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3975
      TabIndex        =   3
      Top             =   1260
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2910
      TabIndex        =   2
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Coefficient Set Name:"
      Height          =   255
      Index           =   5
      Left            =   -210
      TabIndex        =   5
      Top             =   270
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Land Cover Type:"
      Height          =   255
      Index           =   7
      Left            =   495
      TabIndex        =   4
      Top             =   735
      Width           =   1530
   End
End
Attribute VB_Name = "frmAddCoeffSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  frmAddCoeffSet
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey  -  ed.dempsey@noaa.gov
' *************************************************************************************
' *  Description: Form that handles the adding of a new coefficient set.
' *
' *
' *  Called By:  frmPollutants:Add Coefficient Set
' *************************************************************************************

Option Explicit

Private Sub cmdCancel_Click()
   
   Unload Me

End Sub

Private Sub cmdOK_Click()
    
    If modUtil.UniqueName("CoefficientSet", txtCoeffSetName.Text) Then
        'uses code in frmPollutants to do the work
        If g_boolAddCoeff Then
            frmPollutants.AddCoefficient txtCoeffSetName.Text, cboLCType.List(cboLCType.ListIndex)
        Else
            frmNewPollutants.AddCoefficient txtCoeffSetName.Text, cboLCType.List(cboLCType.ListIndex)
        End If
    Else
        MsgBox Err2, vbCritical, "Coefficient set name already in use.  Please enter new name"
        Exit Sub
    End If
    
    Unload frmAddCoeffSet
End Sub

Private Sub Form_Load()
    
    modUtil.InitComboBox cboLCType, "LCTYPE"
    
End Sub

Private Sub txtCoeffSetName_Validate(Cancel As Boolean)
    
    If Trim(Len(txtCoeffSetName.Text)) <> 0 Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
    
End Sub
