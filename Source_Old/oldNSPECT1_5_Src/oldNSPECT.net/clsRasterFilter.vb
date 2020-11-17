Option Strict Off
Option Explicit On
Friend Class clsRasterFilter
	Implements ESRI.ArcGIS.Catalog.IGxObjectFilter
	Implements ESRI.ArcGIS.Catalog.IRasterFormatFilter
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsRasterFilter
	' *************************************************************************************
	' *  Description: ESRI supplied code for filtering rasters in dialog selection environs
	' *
	' *  Called By: modUtil.BrowseForOutputName
	' *************************************************************************************
	' *  Subs:
	
	' *  Misc:
	' *
	' *************************************************************************************
	
	Private m_Format As String
	Private m_Extension As String
	Private m_path As String
	
	
	'UPGRADE_NOTE: Object was upgraded to Object_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function IGxObjectFilter_CanChooseObject(ByVal Object_Renamed As ESRI.ArcGIS.Catalog.IGxObject, ByRef result As ESRI.ArcGIS.Catalog.esriDoubleClickResult) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanChooseObject
		' Only a RasterDataset can be choosen
		Dim pGxds As ESRI.ArcGIS.Catalog.IGxDataset
		IGxObjectFilter_CanChooseObject = False
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If TypeOf Object_Renamed Is ESRI.ArcGIS.Catalog.IGxDataset Then
			pGxds = Object_Renamed
			If pGxds.Type = ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTRasterDataset Then
				IGxObjectFilter_CanChooseObject = True
			End If
		End If
		
		'Cleanup
		'UPGRADE_NOTE: Object pGxds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxds = Nothing
	End Function
	
	'UPGRADE_NOTE: Object was upgraded to Object_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Function IGxObjectFilter_CanDisplayObject(ByVal Object_Renamed As ESRI.ArcGIS.Catalog.IGxObject) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanDisplayObject
		'Display specified rasterdataset
		Dim pGxds As ESRI.ArcGIS.Catalog.IGxDataset
		Dim pRasterDS As ESRI.ArcGIS.Geodatabase.IRasterDataset
		
		IGxObjectFilter_CanDisplayObject = False
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If TypeOf Object_Renamed Is ESRI.ArcGIS.Catalog.IGxFolder Then
			IGxObjectFilter_CanDisplayObject = True
		End If
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If TypeOf Object_Renamed Is ESRI.ArcGIS.Catalog.IGxDataset Then
			pGxds = Object_Renamed
			If pGxds.Type = ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTRasterDataset Then
				pRasterDS = pGxds.Dataset
				If pRasterDS.format = m_Format Then
					IGxObjectFilter_CanDisplayObject = True
				End If
			End If
		End If
		
		'Cleanup
		'UPGRADE_NOTE: Object pGxds may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxds = Nothing
		'UPGRADE_NOTE: Object pRasterDS may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pRasterDS = Nothing
		
	End Function
	
	Private Function IGxObjectFilter_CanSaveObject(ByVal Location As ESRI.ArcGIS.Catalog.IGxObject, ByVal newObjectName As String, ByRef objectAlreadyExists As Boolean) As Boolean Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.CanSaveObject
		Dim pGxFolder As ESRI.ArcGIS.Catalog.IGxFolder
		Dim sName, sPath As String
		Dim pGxFile As ESRI.ArcGIS.Catalog.IGxFile
		Dim fso As Object
		
		On Error GoTo ERH
		IGxObjectFilter_CanSaveObject = False
		'UPGRADE_WARNING: TypeOf has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If Not TypeOf Location Is ESRI.ArcGIS.Catalog.IGxFolder Then Exit Function
		pGxFolder = Location
		sName = newObjectName
		If Len(Trim(sName)) < 1 Then Exit Function
		
		' Get the path
		pGxFile = pGxFolder
		sPath = pGxFile.path
		fso = CreateObject("scripting.filesystemobject")
		
		Dim pos As Short
		pos = InStr(sName, ".")
		
		If pos > 0 Then
			sName = Left(sName, pos - 1)
		End If
		
		newObjectName = sName & m_Extension
		m_path = sPath
		'UPGRADE_WARNING: Couldn't resolve default property of object fso.FileExists. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		objectAlreadyExists = fso.FileExists(sPath & "\" & sName)
		IGxObjectFilter_CanSaveObject = True
		
		'UPGRADE_NOTE: Object pGxFolder may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxFolder = Nothing
		'UPGRADE_NOTE: Object pGxFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pGxFile = Nothing
ERH: 
	End Function
	
	Private ReadOnly Property IGxObjectFilter_Description() As String Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.Description
		Get
			
			IGxObjectFilter_Description = m_Format
			
		End Get
	End Property
	
	Private ReadOnly Property IGxObjectFilter_Name() As String Implements ESRI.ArcGIS.Catalog.IGxObjectFilter.Name
		Get
			
			IGxObjectFilter_Name = m_Format
			
		End Get
	End Property
	
	
	'UPGRADE_NOTE: format was upgraded to format_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Property format_Renamed() As String
		Get
			
			format_Renamed = m_Format
			
		End Get
		Set(ByVal Value As String)
			
			m_Format = Value
			Select Case m_Format
				Case "IMAGINE Image" : m_Extension = ".img"
				Case "TIFF" : m_Extension = ".tif"
				Case "GRID" : m_Extension = ""
			End Select
			
		End Set
	End Property
	
	Public ReadOnly Property path() As String
		Get
			
			path = m_path
			
		End Get
	End Property
	
	Private ReadOnly Property IRasterFormatFilter_Extension() As String Implements ESRI.ArcGIS.Catalog.IRasterFormatFilter.Extension
		Get
			
			IRasterFormatFilter_Extension = m_Extension
			
		End Get
	End Property
	
	Public ReadOnly Property extension() As Object
		Get
			
			'UPGRADE_WARNING: Couldn't resolve default property of object extension. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			extension = m_Extension
			
		End Get
	End Property
End Class