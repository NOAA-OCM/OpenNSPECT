VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRasterFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Option Explicit

Private m_Format As String
Private m_Extension As String
Private m_path As String

Implements IGxObjectFilter
Implements IRasterFormatFilter

Private Function IGxObjectFilter_CanChooseObject(ByVal Object As esriCatalog.IGxObject, result As esriCatalog.esriDoubleClickResult) As Boolean
    ' Only a RasterDataset can be choosen
    Dim pGxds As IGxDataset
    IGxObjectFilter_CanChooseObject = False
    If TypeOf Object Is IGxDataset Then
        Set pGxds = Object
        If pGxds.Type = esriDTRasterDataset Then
            IGxObjectFilter_CanChooseObject = True
        End If
    End If
    
    'Cleanup
    Set pGxds = Nothing
End Function

Private Function IGxObjectFilter_CanDisplayObject(ByVal Object As esriCatalog.IGxObject) As Boolean
    'Display specified rasterdataset
    Dim pGxds As IGxDataset
    Dim pRasterDS As IRasterDataset
    
    IGxObjectFilter_CanDisplayObject = False
    If TypeOf Object Is IGxFolder Then
        IGxObjectFilter_CanDisplayObject = True
    End If
    If TypeOf Object Is IGxDataset Then
        Set pGxds = Object
        If pGxds.Type = esriDTRasterDataset Then
            Set pRasterDS = pGxds.Dataset
            If pRasterDS.format = m_Format Then
                IGxObjectFilter_CanDisplayObject = True
            End If
        End If
    End If
    
    'Cleanup
    Set pGxds = Nothing
    Set pRasterDS = Nothing
    
End Function

Private Function IGxObjectFilter_CanSaveObject(ByVal Location As esriCatalog.IGxObject, ByVal newObjectName As String, objectAlreadyExists As Boolean) As Boolean
    Dim pGxFolder As IGxFolder
    Dim sName As String, sPath As String
    Dim pGxFile As IGxFile
    Dim fso As Object
    
    On Error GoTo ERH
    IGxObjectFilter_CanSaveObject = False
    If Not TypeOf Location Is IGxFolder Then Exit Function
    Set pGxFolder = Location
    sName = newObjectName
    If Len(Trim(sName)) < 1 Then Exit Function
    
    ' Get the path
    Set pGxFile = pGxFolder
    sPath = pGxFile.path
    Set fso = CreateObject("scripting.filesystemobject")
    
    Dim pos As Integer
    pos = InStr(sName, ".")
    
    If pos > 0 Then
        sName = Left(sName, pos - 1)
    End If
    
    newObjectName = sName & m_Extension
    m_path = sPath
    objectAlreadyExists = fso.FileExists(sPath & "\" & sName)
    IGxObjectFilter_CanSaveObject = True
    
    Set pGxFolder = Nothing
    Set pGxFile = Nothing
ERH:
End Function

Private Property Get IGxObjectFilter_Description() As String

    IGxObjectFilter_Description = m_Format

End Property

Private Property Get IGxObjectFilter_Name() As String
    
    IGxObjectFilter_Name = m_Format

End Property

Public Property Let format(ByVal vNewValue As String)
  
  m_Format = vNewValue
  Select Case m_Format
    Case "IMAGINE Image": m_Extension = ".img"
    Case "TIFF": m_Extension = ".tif"
    Case "GRID": m_Extension = ""
  End Select

End Property

Public Property Get format() As String
    
    format = m_Format

End Property

Public Property Get path() As String
  
  path = m_path

End Property

Private Property Get IRasterFormatFilter_Extension() As String
    
    IRasterFormatFilter_Extension = m_Extension

End Property

Public Property Get extension() As Variant
  
  extension = m_Extension

End Property

