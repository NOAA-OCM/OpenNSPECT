VERSION 5.00
Begin VB.Form frmPrjCalc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Coefficient Type Calculation Definition"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4155
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2775
      TabIndex        =   15
      Top             =   3135
      Width           =   1075
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1605
      TabIndex        =   14
      Top             =   3135
      Width           =   1075
   End
   Begin VB.ComboBox cboAttrib 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   525
      Width           =   2250
   End
   Begin VB.ComboBox cboLayer 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   105
      Width           =   2265
   End
   Begin VB.Frame fraDef 
      Caption         =   "Define Code"
      Height          =   2002
      Index           =   1
      Left            =   255
      TabIndex        =   24
      Top             =   960
      Visible         =   0   'False
      Width           =   3480
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   3
         Left            =   1410
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   1515
         Width           =   1785
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   2
         Left            =   1410
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   1115
         Width           =   1785
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   1
         Left            =   1410
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   715
         Width           =   1785
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Index           =   0
         ItemData        =   "frmPrjCalc.frx":0000
         Left            =   1410
         List            =   "frmPrjCalc.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   315
         Width           =   1785
      End
      Begin VB.TextBox txtTypeA 
         Height          =   286
         Index           =   3
         Left            =   1425
         TabIndex        =   28
         Top             =   315
         Width           =   1261
      End
      Begin VB.TextBox txtTypeB 
         Height          =   286
         Index           =   3
         Left            =   1455
         TabIndex        =   27
         Top             =   765
         Width           =   1261
      End
      Begin VB.TextBox txtTypeC 
         Height          =   286
         Index           =   3
         Left            =   1425
         TabIndex        =   26
         Top             =   1140
         Width           =   1261
      End
      Begin VB.TextBox txtTypeD 
         Height          =   286
         Index           =   3
         Left            =   1410
         TabIndex        =   25
         Top             =   1521
         Width           =   1261
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Type 1   equal to"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   13
         Left            =   120
         TabIndex        =   32
         Top             =   315
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Type 2   equal to"
         Height          =   285
         Index           =   12
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Type 3   equal to"
         Height          =   285
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   1125
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Type 4   equal to"
         Height          =   285
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   1515
         Width           =   1200
      End
   End
   Begin VB.Frame fraDef 
      Caption         =   "Define Numeric Range"
      Height          =   2002
      Index           =   0
      Left            =   150
      TabIndex        =   18
      Top             =   975
      Width           =   3757
      Begin VB.TextBox txtTypeB 
         Height          =   286
         Index           =   0
         Left            =   936
         TabIndex        =   4
         Top             =   819
         Width           =   900
      End
      Begin VB.TextBox txtTypeB 
         Height          =   286
         Index           =   1
         Left            =   2457
         TabIndex        =   5
         Top             =   819
         Width           =   900
      End
      Begin VB.TextBox txtTypeC 
         Height          =   286
         Index           =   0
         Left            =   936
         TabIndex        =   6
         Top             =   1170
         Width           =   900
      End
      Begin VB.TextBox txtTypeC 
         Height          =   286
         Index           =   1
         Left            =   2457
         TabIndex        =   7
         Top             =   1170
         Width           =   900
      End
      Begin VB.TextBox txtTypeD 
         Height          =   286
         Index           =   0
         Left            =   936
         TabIndex        =   8
         Top             =   1521
         Width           =   900
      End
      Begin VB.TextBox txtTypeD 
         Height          =   286
         Index           =   1
         Left            =   2457
         TabIndex        =   9
         Top             =   1521
         Width           =   900
      End
      Begin VB.TextBox txtTypeA 
         Height          =   286
         Index           =   1
         Left            =   2457
         TabIndex        =   3
         Top             =   468
         Width           =   900
      End
      Begin VB.TextBox txtTypeA 
         Height          =   286
         Index           =   0
         Left            =   936
         TabIndex        =   2
         Top             =   468
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "and <="
         Height          =   169
         Index           =   14
         Left            =   1900
         TabIndex        =   36
         Top             =   871
         Width           =   400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "and <="
         Height          =   169
         Index           =   15
         Left            =   1900
         TabIndex        =   35
         Top             =   1222
         Width           =   400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "and <="
         Height          =   169
         Index           =   16
         Left            =   1900
         TabIndex        =   34
         Top             =   1573
         Width           =   400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "and <="
         Height          =   169
         Index           =   6
         Left            =   1900
         TabIndex        =   23
         Top             =   520
         Width           =   400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Type 4   >"
         Height          =   286
         Index           =   5
         Left            =   40
         TabIndex        =   22
         Top             =   1573
         Width           =   820
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Type 3   >"
         Height          =   286
         Index           =   4
         Left            =   40
         TabIndex        =   21
         Top             =   1222
         Width           =   820
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Type 2   >"
         Height          =   286
         Index           =   3
         Left            =   40
         TabIndex        =   20
         Top             =   871
         Width           =   820
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Type 1   >"
         ForeColor       =   &H80000008&
         Height          =   286
         Index           =   2
         Left            =   40
         TabIndex        =   19
         Top             =   520
         Width           =   820
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "> and <="
      Height          =   286
      Index           =   9
      Left            =   2106
      TabIndex        =   33
      Top             =   1755
      Width           =   598
   End
   Begin VB.Label Label1 
      Caption         =   "Layer:"
      Height          =   285
      Index           =   0
      Left            =   255
      TabIndex        =   17
      Top             =   150
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "Attribute:"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   555
      Width           =   810
   End
End
Attribute VB_Name = "frmPrjCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_App As IApplication                   'Ref to ArcMap
Private m_pMap As IMap
Private m_pTable As ITable
Private m_pFields As IFields
Private m_intTypeDefType As Integer             '0 = alpha, 1 = numeric
Private m_strXMLFile As String

Private clsTypeDef As clsXMLCoeffTypeDef


Private Sub cboLayer_Click()
    
    cboAttrib.Clear
    
    Dim i As Integer
    Dim pFields As IFields
    Dim pFeatClass As IFeatureClass
    
    Set pFeatClass = modUtil.GetFeatureClass(cboLayer.Text, m_App)
    Set m_pTable = pFeatClass
    Set m_pFields = m_pTable.Fields
    
    For i = 1 To m_pFields.FieldCount - 1
        cboAttrib.AddItem m_pFields.Field(i).Name
    Next i
    
End Sub

Private Sub cmdOK_Click()

    If ValidateData Then
        m_strXMLFile = CreateXMLFile
        
        clsTypeDef.SaveFile m_strXMLFile
        frmPrj.grdCoeffs.TextMatrix(g_intCoeffRow, 6) = m_strXMLFile
        
        
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    modUtil.AddFeatureLayerToComboBox cboLayer, m_pMap, "poly"
    Set clsTypeDef = New clsXMLCoeffTypeDef

    If Len(g_strCoeffCalc) > 0 Then
        clsTypeDef.XML = g_strCoeffCalc
        PopulateForm
    End If
    

End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cboAttrib_Click()
   
    Dim lngField As Long
    Dim pField As IField
    Dim lstAttributes() As String
    Dim pAttList As String
    Dim pCursor As ICursor
    Dim pRow As iRow
    Dim i As Integer
    Dim j As Integer
        
    Screen.MousePointer = vbHourglass
  
    'Clear out the Type combos
    For i = 0 To cboType.UBound
        cboType(i).Clear
    Next i
    
    lngField = m_pFields.FindField(cboAttrib.Text)
    Set pField = m_pFields.Field(lngField)
    
    'Depending on field type, alpha, or numeric, make the correct frame and controls visilbe
    If Not pField Is Nothing Then
   
        Select Case pField.Type
        
            Case esriFieldTypeSmallInteger
                fraDef(0).Visible = True
                fraDef(1).Visible = False
                m_intTypeDefType = 1
                
            Case esriFieldTypeInteger
                fraDef(0).Visible = True
                fraDef(1).Visible = False
                m_intTypeDefType = 1
                
            Case esriFieldTypeSingle
                fraDef(0).Visible = True
                fraDef(1).Visible = False
                m_intTypeDefType = 1
                
            Case esriFieldTypeDouble
                fraDef(0).Visible = True
                fraDef(1).Visible = False
                m_intTypeDefType = 1
                
            Case esriFieldTypeString
                fraDef(0).Visible = False
                fraDef(1).Visible = True
                
                Set pCursor = m_pTable.Search(Nothing, True)
                Set pRow = pCursor.NextRow
        
                While Not pRow Is Nothing
                     If Not (pAttList = "") Then
                         pAttList = pAttList & ","
                     End If
                     pAttList = pAttList & pRow.Value(lngField)
                    Set pRow = pCursor.NextRow
                Wend
                
                lstAttributes = (Split(pAttList, ","))
                For i = 0 To UBound(lstAttributes)
                    
                    For j = 0 To cboType.UBound
                    
                        cboType(j).AddItem (lstAttributes(i))
                       
                    Next j
                    
                Next
                   
                For j = 0 To cboType.UBound
                    LoadUniqueValues cboType(j)
                Next j
                m_intTypeDefType = 0
            
            Case esriFieldTypeDate
                fraDef(1).Visible = False
                fraDef(0).Visible = False
                
            Case esriFieldTypeOID
                fraDef(1).Visible = False
                fraDef(0).Visible = False
                
            Case esriFieldTypeGeometry
                fraDef(1).Visible = False
                fraDef(0).Visible = False
                
            Case esriFieldTypeBlob
                fraDef(1).Visible = False
                fraDef(0).Visible = False
        
        End Select
        
    End If

    Screen.MousePointer = vbNormal
        
End Sub
   
Public Function LoadUniqueValues(combo1 As ComboBox)
    
    Dim dicUnique As Scripting.Dictionary
    Set dicUnique = New Scripting.Dictionary
    
    Dim i As Integer
    
    For i = 1 To combo1.ListCount
        If Not dicUnique.Exists(combo1.List(i)) Then dicUnique.Add combo1.List(i), combo1.List(i)
    Next
    
    combo1.Clear
    
    For i = 0 To dicUnique.Count - 1
        combo1.AddItem dicUnique.Keys(i)
    Next
    
    Set dicUnique = Nothing
    
    
End Function
  
Public Sub init(ByVal pApp As IApplication)
    
    Dim pMxDoc As IMxDocument
    
    Set m_App = pApp
    Set pMxDoc = m_App.Document
    
    Set m_pMap = pMxDoc.FocusMap

End Sub

Private Function CreateXMLFile() As String

   Dim i As Integer
   Dim clsType As New clsXMLCoeffTypeDef
   
   With clsType
   
        .strTDLyrName = Trim(cboLayer.Text)
        .strTDLyrFileName = modUtil.GetFeatureFileName(cboLayer.Text, m_App)
        .strTDAttribute = Trim(cboAttrib.Text)
        .intTDType = m_intTypeDefType
        
        Select Case m_intTypeDefType
            Case 0  'If Alpha
                .strTDDef1 = cboType(0).Text
                .strTDDef2 = cboType(1).Text
                .strTDDef3 = cboType(2).Text
                .strTDDef4 = cboType(3).Text
            Case 1  'If numeric
                .strTDDef1 = txtTypeA(0).Text & "," & txtTypeA(1).Text
                .strTDDef2 = txtTypeB(0).Text & "," & txtTypeB(1).Text
                .strTDDef3 = txtTypeC(0).Text & "," & txtTypeC(1).Text
                .strTDDef4 = txtTypeD(0).Text & "," & txtTypeD(1).Text
        End Select
       
    End With
    
    CreateXMLFile = clsType.XML
    
    
End Function

Private Sub PopulateForm()
    
    Dim strLyrName As String
    Dim strLyrFileName As String
    Dim strAttribute As String
    Dim intYesNo As Integer
    Dim i As Integer
     
    strLyrName = clsTypeDef.strTDLyrName
    strLyrFileName = clsTypeDef.strTDLyrFileName
    strAttribute = clsTypeDef.strTDAttribute
    
    If modUtil.LayerInMap(strLyrName, m_pMap) Then
        cboLayer.ListIndex = modUtil.GetCboIndex(strLyrName, cboLayer)
        cboLayer.Refresh
        cboAttrib.ListIndex = modUtil.GetCboIndex(strAttribute, cboAttrib)
    Else
        If modUtil.AddFeatureLayerToMapFromFileName(strLyrFileName, m_pMap) Then
            With cboLayer
                .AddItem strLyrName
                .Refresh
                .ListIndex = modUtil.GetCboIndex(strLyrName, cboLayer)
            End With
        Else
            intYesNo = MsgBox("Could not find the layer: " & strLyrFileName & ".  Would you like to browse for it", _
            vbCritical, "Missing Layer")
            
            If intYesNo = vbYes Then
                clsTypeDef.strTDLyrFileName = modUtil.AddInputFromGxBrowser(cboLayer, frmPrjCalc, "Feature")
                    If clsTypeDef.strTDLyrFileName <> "" Then
                        If modUtil.AddFeatureLayerToMapFromFileName(clsTypeDef.strTDLyrFileName, m_pMap) Then
                            strLyrName = modUtil.SplitFileName(clsTypeDef.strTDLyrFileName)
                            cboLayer.AddItem strLyrName
                            cboLayer.ListIndex = modUtil.GetCboIndex(strLyrName, cboLayer)
                        End If
                    Else
                        Exit Sub
                    End If
            Else
                Exit Sub
            End If
        End If
    End If
    
    Select Case clsTypeDef.intTDType
    
        Case 0  'If Alpha
            fraDef(0).Visible = False
            fraDef(1).Visible = True
            cboType(0).ListIndex = modUtil.GetCboIndex(clsTypeDef.strTDDef1, cboType(0))
            cboType(1).ListIndex = modUtil.GetCboIndex(clsTypeDef.strTDDef2, cboType(1))
            cboType(2).ListIndex = modUtil.GetCboIndex(clsTypeDef.strTDDef3, cboType(2))
            cboType(3).ListIndex = modUtil.GetCboIndex(clsTypeDef.strTDDef4, cboType(3))
            
        Case 1  'If numeric
            fraDef(0).Visible = True
            fraDef(1).Visible = False
            txtTypeA(0) = Split(clsTypeDef.strTDDef1, ",")(0)
            txtTypeA(1) = Split(clsTypeDef.strTDDef1, ",")(1)
            txtTypeB(0) = Split(clsTypeDef.strTDDef2, ",")(0)
            txtTypeB(1) = Split(clsTypeDef.strTDDef2, ",")(1)
            txtTypeC(0) = Split(clsTypeDef.strTDDef3, ",")(0)
            txtTypeC(1) = Split(clsTypeDef.strTDDef3, ",")(1)
            txtTypeD(0) = Split(clsTypeDef.strTDDef4, ",")(0)
            txtTypeD(1) = Split(clsTypeDef.strTDDef4, ",")(1)

    End Select
    
End Sub



Private Function ValidateData() As Boolean
'Function returns true if form inputs are valid
    'For alpha values, all must be unique
    'For numeric, column 1 value must be <= column 2 value and value sets must be
    'mutually exclusive.
    
    Const strWarning = "The first value must be less than or equal to the second."
    Const strWarning2 = "Incorrect Values Found"
    Const strWarning3 = "Value pairs must be mutually exclusive"
    Const strWarning4 = "Duplicate Values Found"
    
    Dim i As Integer
    Dim j As Integer
    Dim strTypeValue As String  'String values of types
    Dim intValue As Integer     'Int values of types
    Dim varValuesA As New Collection    'Collection of first column values
    Dim varValuesB As New Collection    'Collection of second column values
    
    
    'Check Name
    If cboLayer.Text = "" Then
        MsgBox "Please select a Layer.  If the combobox is empty, you must add a layer to the map.", vbCritical, "Select Layer"
        ValidateData = False
        Exit Function
    End If
    
    'Check Attribute
    If cboAttrib.Text = "" Then
        MsgBox "Please select an attribute. ", vbCritical, "Select Attribute"
        cboAttrib.SetFocus
        ValidateData = False
        Exit Function
    End If
    
    'Based on attribute type, int or alpha, text the rest
    Select Case m_intTypeDefType
        Case 0 'Alpha
            For i = 0 To cboType.UBound
                'If blank, set focus to offending cbo
                If cboType(i).Text = "" Then
                    MsgBox "Please select a value for Type " & i + 1 & ".", vbCritical, "Missing Value"
                    ValidateData = False
                    cboType(i).SetFocus
                    Exit Function
                End If
            Next i
            
            'Now test for unique values
            For i = 0 To cboType.UBound
                strTypeValue = cboType(i).Text
                For j = 0 To cboType.UBound
                    If j <> i Then
                        If cboType(j).Text = strTypeValue Then
                            MsgBox "Type values must be unique.", vbCritical, "Duplicate Values Found"
                            ValidateData = False
                            cboType(j).SetFocus
                            Exit Function
                        End If
                    End If
                Next j
            Next i
                
        Case 1 'Numeric
        
        On Error GoTo ErrHandler    'Lazy way to handle strings/blanks
            
            If Not (CInt(txtTypeA(0).Text) <= CInt(txtTypeA(1).Text)) Then
                MsgBox strWarning, vbCritical, strWarning2
                ValidateData = False
                txtTypeA(0).SetFocus
                Exit Function
            Else
                varValuesA.Add txtTypeA(0).Text
                varValuesA.Add txtTypeA(1).Text
            End If
            
            If Not (CInt(txtTypeB(0).Text) <= CInt(txtTypeB(1).Text)) Then
                MsgBox strWarning, vbCritical, strWarning2
                ValidateData = False
                txtTypeB(0).SetFocus
                Exit Function
            Else
                varValuesA.Add txtTypeB(0).Text
                varValuesA.Add txtTypeB(1).Text
            End If
            
            If Not (CInt(txtTypeC(0).Text) <= CInt(txtTypeC(1).Text)) Then
                MsgBox strWarning, vbCritical, strWarning2
                ValidateData = False
                txtTypeC(0).SetFocus
            Else
                varValuesA.Add txtTypeC(0).Text
                varValuesA.Add txtTypeC(1).Text
            End If
            
            If Not (CInt(txtTypeD(0).Text) <= CInt(txtTypeD(1).Text)) Then
                MsgBox strWarning, vbCritical, strWarning2
                ValidateData = False
                txtTypeD(0).SetFocus
                Exit Function
            Else
                varValuesA.Add txtTypeD(0).Text
                varValuesA.Add txtTypeD(1).Text
            End If
            
            'Test for mutually exclusive values
            For i = 1 To varValuesA.Count
                intValue = varValuesA(i)
                    For j = 1 To varValuesA.Count
                        If j <> i Then
                            If varValuesA(i) = varValuesA(j) Then
                                MsgBox strWarning3, vbCritical, strWarning4
                                ValidateData = False
                                Exit Function
                            End If
                        End If
                    Next j
            Next i
            
'            For i = 1 To varValuesB.Count
'                intValue = varValuesB(i)
'                    For j = 1 To varValuesB.Count
'                        If j <> i Then
'                            If varValuesB(i) = varValuesB(j) Then
'                                MsgBox strWarning3, vbCritical, strWarning4
'                                ValidateData = False
'                                Exit Function
'                            End If
'                        End If
'                    Next j
'            Next i
            
    End Select
    
    ValidateData = True
    
    
    
    ''Set varValuesA = Nothing
    'Set varValuesB = Nothing
Exit Function

ErrHandler:
    MsgBox "Numeric values only please.", vbCritical, "Numbers only please."
    Set varValuesA = Nothing
    Set varValuesB = Nothing
End Function

