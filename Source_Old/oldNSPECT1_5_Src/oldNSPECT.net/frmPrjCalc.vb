Option Strict Off
Option Explicit On
Friend Class frmPrjCalc
	Inherits System.Windows.Forms.Form
	
	Private m_App As ESRI.ArcGIS.Framework.IApplication 'Ref to ArcMap
	Private m_pMap As ESRI.ArcGIS.Carto.IMap
	Private m_pTable As ESRI.ArcGIS.Geodatabase.ITable
	Private m_pFields As ESRI.ArcGIS.Geodatabase.IFields
	Private m_intTypeDefType As Short '0 = alpha, 1 = numeric
	Private m_strXMLFile As String
	
	Private clsTypeDef As clsXMLCoeffTypeDef
	
	
	'UPGRADE_WARNING: Event cboLayer.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboLayer_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLayer.SelectedIndexChanged
		
		cboAttrib.Items.Clear()
		
		Dim i As Short
		Dim pFields As ESRI.ArcGIS.Geodatabase.IFields
		Dim pFeatClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
		
		pFeatClass = modUtil.GetFeatureClass((cboLayer.Text), m_App)
		m_pTable = pFeatClass
		'UPGRADE_WARNING: Couldn't resolve default property of object m_pTable.Fields. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_pFields = m_pTable.Fields
		
		For i = 1 To m_pFields.FieldCount - 1
			cboAttrib.Items.Add(m_pFields.Field(i).Name)
		Next i
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		If ValidateData Then
			m_strXMLFile = CreateXMLFile
			
			clsTypeDef.SaveFile(m_strXMLFile)
			frmPrj.grdCoeffs.set_TextMatrix(g_intCoeffRow, 6, m_strXMLFile)
			
			
			Me.Close()
		End If
		
	End Sub
	
	Private Sub frmPrjCalc_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		modUtil.AddFeatureLayerToComboBox(cboLayer, m_pMap, "poly")
		clsTypeDef = New clsXMLCoeffTypeDef
		
		If Len(g_strCoeffCalc) > 0 Then
			clsTypeDef.XML = g_strCoeffCalc
			PopulateForm()
		End If
		
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	'UPGRADE_WARNING: Event cboAttrib.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboAttrib_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboAttrib.SelectedIndexChanged
		
		Dim lngField As Integer
		Dim pField As ESRI.ArcGIS.Geodatabase.IField
		Dim lstAttributes() As String
		Dim pAttList As String
		Dim pCursor As ESRI.ArcGIS.Geodatabase.ICursor
		Dim pRow As ESRI.ArcGIS.Geodatabase.IRow
		Dim i As Short
		Dim j As Short
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		'Clear out the Type combos
		For i = 0 To cboType.UBound
			cboType(i).Items.Clear()
		Next i
		
		lngField = m_pFields.FindField(cboAttrib.Text)
		pField = m_pFields.Field(lngField)
		
		'Depending on field type, alpha, or numeric, make the correct frame and controls visilbe
		If Not pField Is Nothing Then
			
			Select Case pField.Type
				
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeSmallInteger
					fraDef(0).Visible = True
					fraDef(1).Visible = False
					m_intTypeDefType = 1
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeInteger
					fraDef(0).Visible = True
					fraDef(1).Visible = False
					m_intTypeDefType = 1
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeSingle
					fraDef(0).Visible = True
					fraDef(1).Visible = False
					m_intTypeDefType = 1
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeDouble
					fraDef(0).Visible = True
					fraDef(1).Visible = False
					m_intTypeDefType = 1
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeString
					fraDef(0).Visible = False
					fraDef(1).Visible = True
					
					pCursor = m_pTable.Search(Nothing, True)
					pRow = pCursor.NextRow
					
					While Not pRow Is Nothing
						If Not (pAttList = "") Then
							pAttList = pAttList & ","
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object pRow.Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						pAttList = pAttList & pRow.Value(lngField)
						pRow = pCursor.NextRow
					End While
					
					lstAttributes = VB6.CopyArray((Split(pAttList, ",")))
					For i = 0 To UBound(lstAttributes)
						
						For j = 0 To cboType.UBound
							
							cboType(j).Items.Add((lstAttributes(i)))
							
						Next j
						
					Next 
					
					For j = 0 To cboType.UBound
						LoadUniqueValues(cboType(j))
					Next j
					m_intTypeDefType = 0
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeDate
					fraDef(1).Visible = False
					fraDef(0).Visible = False
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeOID
					fraDef(1).Visible = False
					fraDef(0).Visible = False
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeGeometry
					fraDef(1).Visible = False
					fraDef(0).Visible = False
					
				Case ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeBlob
					fraDef(1).Visible = False
					fraDef(0).Visible = False
					
			End Select
			
		End If
		
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		'UPGRADE_ISSUE: Screen property Screen.MousePointer does not support custom mousepointers. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="45116EAB-7060-405E-8ABE-9DBB40DC2E86"'
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = vbNormal
		
	End Sub
	
	Public Function LoadUniqueValues(ByRef combo1 As System.Windows.Forms.ComboBox) As Object
		
		Dim dicUnique As Scripting.Dictionary
		dicUnique = New Scripting.Dictionary
		
		Dim i As Short
		
		For i = 1 To combo1.Items.Count
			If Not dicUnique.Exists(VB6.GetItemString(combo1, i)) Then dicUnique.Add(VB6.GetItemString(combo1, i), VB6.GetItemString(combo1, i))
		Next 
		
		combo1.Items.Clear()
		
		For i = 0 To dicUnique.Count - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object dicUnique.Keys(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			combo1.Items.Add(dicUnique.Keys(i))
		Next 
		
		'UPGRADE_NOTE: Object dicUnique may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		dicUnique = Nothing
		
		
	End Function
	
	Public Sub init(ByVal pApp As ESRI.ArcGIS.Framework.IApplication)
		
		Dim pMxDoc As ESRI.ArcGIS.ArcMapUI.IMxDocument
		
		m_App = pApp
		pMxDoc = m_App.Document
		
		m_pMap = pMxDoc.FocusMap
		
	End Sub
	
	Private Function CreateXMLFile() As String
		
		Dim i As Short
		Dim clsType As New clsXMLCoeffTypeDef
		
		With clsType
			
			.strTDLyrName = Trim(cboLayer.Text)
			.strTDLyrFileName = modUtil.GetFeatureFileName((cboLayer.Text), m_App)
			.strTDAttribute = Trim(cboAttrib.Text)
			.intTDType = m_intTypeDefType
			
			Select Case m_intTypeDefType
				Case 0 'If Alpha
					.strTDDef1 = cboType(0).Text
					.strTDDef2 = cboType(1).Text
					.strTDDef3 = cboType(2).Text
					.strTDDef4 = cboType(3).Text
				Case 1 'If numeric
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
		Dim intYesNo As Short
		Dim i As Short
		
		strLyrName = clsTypeDef.strTDLyrName
		strLyrFileName = clsTypeDef.strTDLyrFileName
		strAttribute = clsTypeDef.strTDAttribute
		
		If modUtil.LayerInMap(strLyrName, m_pMap) Then
			cboLayer.SelectedIndex = modUtil.GetCboIndex(strLyrName, cboLayer)
			cboLayer.Refresh()
			cboAttrib.SelectedIndex = modUtil.GetCboIndex(strAttribute, cboAttrib)
		Else
			If modUtil.AddFeatureLayerToMapFromFileName(strLyrFileName, m_pMap) Then
				With cboLayer
					.Items.Add(strLyrName)
					.Refresh()
					.SelectedIndex = modUtil.GetCboIndex(strLyrName, cboLayer)
				End With
			Else
				intYesNo = MsgBox("Could not find the layer: " & strLyrFileName & ".  Would you like to browse for it", MsgBoxStyle.Critical, "Missing Layer")
				
				If intYesNo = MsgBoxResult.Yes Then
					clsTypeDef.strTDLyrFileName = modUtil.AddInputFromGxBrowser(cboLayer, Me, "Feature")
					If clsTypeDef.strTDLyrFileName <> "" Then
						If modUtil.AddFeatureLayerToMapFromFileName((clsTypeDef.strTDLyrFileName), m_pMap) Then
							strLyrName = modUtil.SplitFileName((clsTypeDef.strTDLyrFileName))
							cboLayer.Items.Add(strLyrName)
							cboLayer.SelectedIndex = modUtil.GetCboIndex(strLyrName, cboLayer)
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
			
			Case 0 'If Alpha
				fraDef(0).Visible = False
				fraDef(1).Visible = True
				cboType(0).SelectedIndex = modUtil.GetCboIndex((clsTypeDef.strTDDef1), cboType(0))
				cboType(1).SelectedIndex = modUtil.GetCboIndex((clsTypeDef.strTDDef2), cboType(1))
				cboType(2).SelectedIndex = modUtil.GetCboIndex((clsTypeDef.strTDDef3), cboType(2))
				cboType(3).SelectedIndex = modUtil.GetCboIndex((clsTypeDef.strTDDef4), cboType(3))
				
			Case 1 'If numeric
				fraDef(0).Visible = True
				fraDef(1).Visible = False
				txtTypeA(0).Text = Split(clsTypeDef.strTDDef1, ",")(0)
				txtTypeA(1).Text = Split(clsTypeDef.strTDDef1, ",")(1)
				txtTypeB(0).Text = Split(clsTypeDef.strTDDef2, ",")(0)
				txtTypeB(1).Text = Split(clsTypeDef.strTDDef2, ",")(1)
				txtTypeC(0).Text = Split(clsTypeDef.strTDDef3, ",")(0)
				txtTypeC(1).Text = Split(clsTypeDef.strTDDef3, ",")(1)
				txtTypeD(0).Text = Split(clsTypeDef.strTDDef4, ",")(0)
				txtTypeD(1).Text = Split(clsTypeDef.strTDDef4, ",")(1)
				
		End Select
		
	End Sub
	
	
	
	Private Function ValidateData() As Boolean
		'Function returns true if form inputs are valid
		'For alpha values, all must be unique
		'For numeric, column 1 value must be <= column 2 value and value sets must be
		'mutually exclusive.
		
		Const strWarning As String = "The first value must be less than or equal to the second."
		Const strWarning2 As String = "Incorrect Values Found"
		Const strWarning3 As String = "Value pairs must be mutually exclusive"
		Const strWarning4 As String = "Duplicate Values Found"
		
		Dim i As Short
		Dim j As Short
		Dim strTypeValue As String 'String values of types
		Dim intValue As Short 'Int values of types
		Dim varValuesA As New Collection 'Collection of first column values
		Dim varValuesB As New Collection 'Collection of second column values
		
		
		'Check Name
		If cboLayer.Text = "" Then
			MsgBox("Please select a Layer.  If the combobox is empty, you must add a layer to the map.", MsgBoxStyle.Critical, "Select Layer")
			ValidateData = False
			Exit Function
		End If
		
		'Check Attribute
		If cboAttrib.Text = "" Then
			MsgBox("Please select an attribute. ", MsgBoxStyle.Critical, "Select Attribute")
			cboAttrib.Focus()
			ValidateData = False
			Exit Function
		End If
		
		'Based on attribute type, int or alpha, text the rest
		Select Case m_intTypeDefType
			Case 0 'Alpha
				For i = 0 To cboType.UBound
					'If blank, set focus to offending cbo
					If cboType(i).Text = "" Then
						MsgBox("Please select a value for Type " & i + 1 & ".", MsgBoxStyle.Critical, "Missing Value")
						ValidateData = False
						cboType(i).Focus()
						Exit Function
					End If
				Next i
				
				'Now test for unique values
				For i = 0 To cboType.UBound
					strTypeValue = cboType(i).Text
					For j = 0 To cboType.UBound
						If j <> i Then
							If cboType(j).Text = strTypeValue Then
								MsgBox("Type values must be unique.", MsgBoxStyle.Critical, "Duplicate Values Found")
								ValidateData = False
								cboType(j).Focus()
								Exit Function
							End If
						End If
					Next j
				Next i
				
			Case 1 'Numeric
				
				On Error GoTo ErrHandler 'Lazy way to handle strings/blanks
				
				If Not (CShort(txtTypeA(0).Text) <= CShort(txtTypeA(1).Text)) Then
					MsgBox(strWarning, MsgBoxStyle.Critical, strWarning2)
					ValidateData = False
					txtTypeA(0).Focus()
					Exit Function
				Else
					varValuesA.Add(txtTypeA(0).Text)
					varValuesA.Add(txtTypeA(1).Text)
				End If
				
				If Not (CShort(txtTypeB(0).Text) <= CShort(txtTypeB(1).Text)) Then
					MsgBox(strWarning, MsgBoxStyle.Critical, strWarning2)
					ValidateData = False
					txtTypeB(0).Focus()
					Exit Function
				Else
					varValuesA.Add(txtTypeB(0).Text)
					varValuesA.Add(txtTypeB(1).Text)
				End If
				
				If Not (CShort(txtTypeC(0).Text) <= CShort(txtTypeC(1).Text)) Then
					MsgBox(strWarning, MsgBoxStyle.Critical, strWarning2)
					ValidateData = False
					txtTypeC(0).Focus()
				Else
					varValuesA.Add(txtTypeC(0).Text)
					varValuesA.Add(txtTypeC(1).Text)
				End If
				
				If Not (CShort(txtTypeD(0).Text) <= CShort(txtTypeD(1).Text)) Then
					MsgBox(strWarning, MsgBoxStyle.Critical, strWarning2)
					ValidateData = False
					txtTypeD(0).Focus()
					Exit Function
				Else
					varValuesA.Add(txtTypeD(0).Text)
					varValuesA.Add(txtTypeD(1).Text)
				End If
				
				'Test for mutually exclusive values
				For i = 1 To varValuesA.Count()
					'UPGRADE_WARNING: Couldn't resolve default property of object varValuesA(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					intValue = varValuesA.Item(i)
					For j = 1 To varValuesA.Count()
						If j <> i Then
							'UPGRADE_WARNING: Couldn't resolve default property of object varValuesA(j). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object varValuesA(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If varValuesA.Item(i) = varValuesA.Item(j) Then
								MsgBox(strWarning3, MsgBoxStyle.Critical, strWarning4)
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
		MsgBox("Numeric values only please.", MsgBoxStyle.Critical, "Numbers only please.")
		'UPGRADE_NOTE: Object varValuesA may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		varValuesA = Nothing
		'UPGRADE_NOTE: Object varValuesB may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		varValuesB = Nothing
	End Function
End Class