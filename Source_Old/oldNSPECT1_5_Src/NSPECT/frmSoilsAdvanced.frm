VERSION 5.00
Begin VB.Form frmSoilsAdvanced 
   Caption         =   "MUSLE Parameters"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "MUSLE Specific Parameters "
      Height          =   2190
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtMUSLEExp 
         Height          =   330
         Left            =   1845
         TabIndex        =   2
         Top             =   1500
         Width           =   525
      End
      Begin VB.TextBox txtMUSLEVal 
         Height          =   330
         Left            =   360
         TabIndex        =   1
         Top             =   1470
         Width           =   510
      End
      Begin VB.Label Label5 
         Caption         =   "* K * C* P * LS"
         Height          =   240
         Left            =   2490
         TabIndex        =   7
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "* (Q * qp)^"
         Height          =   240
         Left            =   990
         TabIndex        =   6
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label3 
         Caption         =   "Locally calibrated MUSLE equation for sediment yield being used by N-SPECT: "
         Height          =   480
         Left            =   270
         TabIndex        =   5
         Top             =   930
         Width           =   3675
      End
      Begin VB.Label Label2 
         Caption         =   "95 * (Q * qp)^2 * K * C * P * LS"
         Height          =   300
         Left            =   495
         TabIndex        =   4
         Top             =   570
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "MUSLE Equation for sediment yield:"
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   285
         Width           =   3780
      End
   End
End
Attribute VB_Name = "frmSoilsAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
