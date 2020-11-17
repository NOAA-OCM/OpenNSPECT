VERSION 5.00
Begin VB.Form frmCopyCoeffSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy Coefficient Set"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5505
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCoeffSet 
      Height          =   315
      ItemData        =   "frmCopyCoeffSet.frx":0000
      Left            =   2640
      List            =   "frmCopyCoeffSet.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   195
      Width           =   2370
   End
   Begin VB.TextBox txtCoeffSetName 
      Height          =   300
      Left            =   2640
      TabIndex        =   1
      Top             =   690
      Width           =   2340
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2985
      TabIndex        =   2
      Top             =   1215
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4035
      TabIndex        =   3
      Top             =   1215
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Copy from Coefficient Set:"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   270
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Coefficient Set Name:"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   4
      Top             =   750
      Width           =   2100
   End
End
Attribute VB_Name = "frmCopyCoeffSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************************************************
' *  frmCopyCoeffSet
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey  -  ed.dempsey@noaa.gov
' *************************************************************************************
' *  Description:  Form that allows for the copying of pollutant coefficient sets
' *
' *
' *  Called By:  frmPollutants
' *************************************************************************************
Option Explicit

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If modUtil.UniqueName("CoefficientSet", txtCoeffSetName.Text) And Trim(txtCoeffSetName.Text) <> "" Then
        If g_boolCopyCoeff Then
            frmPollutants.CopyCoefficient txtCoeffSetName.Text, cboCoeffSet.Text
        Else
            frmNewPollutants.CopyCoefficient txtCoeffSetName.Text, cboCoeffSet.Text
        End If
    Else
        MsgBox "The name you have choosen for coefficient set is already in use.  Please pick another.", vbCritical, _
        "Name In Use"
        With txtCoeffSetName
            .SelStart = 0
            .SelLength = Len(txtCoeffSetName.Text)
            .SetFocus
        End With
        
    End If
    
End Sub

Public Sub init(rsCoeffSet As ADODB.Recordset)
'The form is passed a recordest containing the names of all coefficient sets, allows for
'easier populating

    Dim i As Integer
    
    rsCoeffSet.MoveFirst
    
    For i = 0 To rsCoeffSet.RecordCount - 1
        cboCoeffSet.AddItem rsCoeffSet!Name
        rsCoeffSet.MoveNext
    Next i
        
End Sub


