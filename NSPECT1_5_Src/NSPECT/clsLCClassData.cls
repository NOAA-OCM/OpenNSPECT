VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLCClassData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  clsHelp
' *************************************************************************************
' *  Description:  For simplicity's sake, I created this class to handle the data operations
' *  involved in the LandClass form.  Add, Delete, copy etc
' *
' *  Called By:  Land Cover form
' *************************************************************************************
Option Explicit

Const cstTableName = "LCClass"
Const cstDefaultQuery = "SELECT * FROM LCClass WHERE "

Private rsTemp                  As ADODB.Recordset

'Special Events
Event Refresh()
Event Exists()
Event NotFound()
Event Synchronised()
Event NotSynchronised()
Event Loaded()
Event Saved()
Event Added()
Event Deleted()
Event SavedMembers()
Event Enumerate()

'Object States: Not required to be persistent
Private mstrSQL                As String
Private mblnIsDirty            As Boolean
Private mblnIsSynchronised     As Boolean
Private mblnPersist            As Boolean
Private mblnNoMatch            As Boolean

'local variable(s)
Private mlngLCClassID          As Long
Private mlngValue              As Long
Private mstrName               As String
Private mlngLCTypeID           As Long
Private sngCNA                 As Single
Private sngCNB                 As Single
Private sngCNC                 As Single
Private sngCND                 As Single
Private sngCoverFactor         As Single
Private mlngW_WL               As Long

'Object States: Not required to be persistent
Public Property Get IsDirty() As Boolean
    IsDirty = mblnIsDirty
End Property

Public Property Get IsSynchronised() As Boolean
    IsSynchronised = mblnIsSynchronised
End Property

Private Property Let IsSynchronised(ByVal Param As Boolean)
    mblnIsSynchronised = Param
    If mblnIsSynchronised Then
       RaiseEvent Synchronised
    End If
End Property

Public Property Get Persist() As Boolean
    Persist = mblnPersist
End Property

Public Property Let Persist(ByVal Param As Boolean)
  mblnPersist = Param
  If mblnPersist = True Then
      If rsTemp Is Nothing Then OpenTable
  End If
End Property

'Persistent Object properties
Public Property Let LCClassID(ByVal Param As Long)
  mlngLCClassID = Param
  mblnIsDirty = True
End Property

Public Property Get LCClassID() As Long
  LCClassID = mlngLCClassID
End Property

Public Property Let Value(ByVal Param As Long)
  mlngValue = Param
  mblnIsDirty = True
End Property

Public Property Get Value() As Long
  Value = mlngValue
End Property

Public Property Let Name(ByVal Param As String)
  mstrName = Param
  mblnIsDirty = True
End Property

Public Property Get Name() As String
  Name = mstrName
End Property

Public Property Let LCTypeID(ByVal Param As Long)
  mlngLCTypeID = Param
  mblnIsDirty = True
End Property

Public Property Get LCTypeID() As Long
  LCTypeID = mlngLCTypeID
End Property

Public Property Let CNA(ByVal Param As Single)
  sngCNA = Param
  mblnIsDirty = True
End Property

Public Property Get CNA() As Single
  CNA = sngCNA
End Property

Public Property Let CNB(ByVal Param As Single)
  sngCNB = Param
  mblnIsDirty = True
End Property

Public Property Get CNB() As Single
  CNB = sngCNB
End Property

Public Property Let CNC(ByVal Param As Single)
  sngCNC = Param
  mblnIsDirty = True
End Property

Public Property Get CNC() As Single
  CNC = sngCNC
End Property

Public Property Let CND(ByVal Param As Single)
  sngCND = Param
  mblnIsDirty = True
End Property

Public Property Get CND() As Single
  CND = sngCND
End Property

Public Property Let CoverFactor(ByVal Param As Single)
  sngCoverFactor = Param
  mblnIsDirty = True
End Property

Public Property Get CoverFactor() As Single
  CoverFactor = sngCoverFactor
End Property

Public Property Let W_WL(ByVal Param As Long)
  mlngW_WL = Param
  mblnIsDirty = True
End Property

Public Property Get W_WL() As Long
  W_WL = mlngW_WL
End Property

Public Property Let SQL(ByVal vData As String)
  mstrSQL = vData
  'Set new record source using SQL statements
  If mstrSQL = "" Then Exit Property
  If Not rsTemp Is Nothing Then CloseTable
  OpenTable mstrSQL
End Property

Public Property Get SQL() As String
    SQL = mstrSQL
End Property

Private Sub Class_Initialize()
    mblnIsDirty = False
    mblnIsSynchronised = False
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Terminate
End Sub

Public Sub Connect()
  OpenTable
  MoveFirst
End Sub

Public Sub DisConnect()
  CloseTable
End Sub

Public Function Load(Optional lngLCClassID As Variant) As Boolean

    On Error GoTo ErrorBlock

    If rsTemp Is Nothing Then OpenTable
    If IsMissing(lngLCClassID) Then lngLCClassID = mlngLCClassID

    If rsTemp.BOF And rsTemp.EOF Then Exit Function
    rsTemp.MoveFirst
    rsTemp.Find "LCClassID = " & lngLCClassID
    NoMatch = EOF

    If Not NoMatch Then
        RefreshMembers
        Load = True
        IsSynchronised = True
        mblnIsDirty = False
        RaiseEvent Loaded
    Else
        RaiseEvent NotFound
        Load = False
        IsSynchronised = False
    End If

ExitBlock:

    Exit Function

ErrorBlock:

    Err.Raise vbObjectError, TypeName(Me) & ".Load", "Data Engine Misbehaved on Load, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description
    Resume ExitBlock

End Function

Public Function Exists(Optional lngLCClassID As Variant) As Boolean

    On Error GoTo ErrorBlock

    Dim tbTemp As ADODB.Recordset
    Dim lngRecords As Long

    If rsTemp Is Nothing Then OpenTable
    If IsMissing(lngLCClassID) Then lngLCClassID = mlngLCClassID
    Set tbTemp = g_ADOConn.Execute("SELECT COUNT (LCClassID) FROM LCClass WHERE LCClassID = " & lngLCClassID)
    lngRecords = Val(tbTemp.GetString)

    If lngRecords = 1 Then
        NoMatch = False
        RaiseEvent Exists
    Else
        NoMatch = True
    End If

    Exists = Not NoMatch

ExitBlock:

    On Error Resume Next
    tbTemp.Close
    Set tbTemp = Nothing
    Exit Function

ErrorBlock:

    Err.Raise vbObjectError, TypeName(Me) & ".Exists", "Data Engine Misbehaved on Exists, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description
    Resume ExitBlock

End Function

Private Sub RefreshMembers()

    On Error Resume Next

    If rsTemp.BOF Or rsTemp.EOF Then Exit Sub

    With rsTemp

      mlngLCClassID = !LCClassID
      mlngValue = !Value
      mstrName = !Name
      mlngLCTypeID = !LCTypeID
      sngCNA = ![CN-A]
      sngCNB = ![CN-B]
      sngCNC = ![CN-C]
      sngCND = ![CN-D]
      sngCoverFactor = !CoverFactor
      mlngW_WL = !W_WL

    End With

    IsSynchronised = True
    RaiseEvent Refresh

End Sub

Public Sub AddNew()

    On Error GoTo ErrorBlock

    If rsTemp Is Nothing Then OpenTable
          
    rsTemp.AddNew
    SaveMembers
    
    rsTemp.Update
    mblnIsDirty = False
    RaiseEvent Added
    
    g_intLCClassid = rsTemp!LCClassID
    
ExitBlock:

    Exit Sub

ErrorBlock:

    Err.Raise vbObjectError, TypeName(Me) & ".AddNew", "Data Engine Misbehaved on AddNew, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description
    Resume ExitBlock

End Sub

Public Sub Delete()

On Error GoTo ErrorBlock

    If rsTemp Is Nothing Then OpenTable
          
    rsTemp.Delete adAffectCurrent
    'SaveMembers
    
    rsTemp.Update
    mblnIsDirty = False
    
ExitBlock:

    Exit Sub

ErrorBlock:

    Err.Raise vbObjectError, TypeName(Me) & ".AddNew", "Data Engine Misbehaved on AddNew, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description
    Resume ExitBlock

End Sub

Public Sub SaveChanges()

    On Error GoTo ErrorBlock

    Dim varBookmark As Variant

    If Not mblnIsSynchronised Then
        'Cannot Save Changes to object without first being loaded or synchronised
        RaiseEvent NotSynchronised
    Else
        If mblnIsDirty Then
            SaveMembers
            rsTemp.Update
            mblnIsDirty = False
            RaiseEvent Saved
        End If
    End If

ExitBlock:

    On Error Resume Next
    Exit Sub

ErrorBlock:

    Err.Raise vbObjectError, TypeName(Me) & ".SaveChanges", "Data Engine Misbehaved on SaveChanges, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description
    Resume ExitBlock

End Sub

Private Sub SaveMembers()

    On Error Resume Next

    With rsTemp
    
        !Value = mlngValue
        !Name = mstrName
        !LCTypeID = mlngLCTypeID
        ![CN-A] = sngCNA
        ![CN-B] = sngCNB
        ![CN-C] = sngCNC
        ![CN-D] = sngCND
        !CoverFactor = sngCoverFactor
        !W_WL = mlngW_WL
        !LCTypeID = mlngLCTypeID
    End With

    IsSynchronised = True
    RaiseEvent SavedMembers

End Sub

Private Sub CopyProperties(objCopy As clsLCClassData)

    On Error GoTo ErrorBlock

    With rsTemp

      mlngLCClassID = objCopy.LCClassID
      mlngValue = objCopy.Value
      mstrName = objCopy.Name
      mlngLCTypeID = objCopy.LCTypeID
      sngCNA = objCopy.CNA
      sngCNB = objCopy.CNB
      sngCNC = objCopy.CNC
      sngCND = objCopy.CND
      sngCoverFactor = objCopy.CoverFactor
      mlngW_WL = objCopy.W_WL

    End With

ExitBlock:

    Exit Sub

ErrorBlock:

    Err.Raise vbObjectError, TypeName(Me) & ".CopyProperties", "Data Engine Misbehaved on CopyProperties, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description
    Resume ExitBlock

End Sub

Public Function CopyObject(ByVal objCopy As clsLCClassData) As clsLCClassData
    If Not objCopy Is Nothing Then
        CopyProperties objCopy
    End If
End Function

Friend Sub OpenTable(Optional strSQL As String)

    On Error GoTo ErrorBlock
    
    If rsTemp Is Nothing Then
      Set rsTemp = New ADODB.Recordset
      rsTemp.CursorType = adOpenDynamic
      rsTemp.LockType = adLockOptimistic
    End If
    
    If strSQL = "" Then
      rsTemp.Open "LCClass", modUtil.g_ADOConn, adOpenDynamic
    Else
      rsTemp.Open strSQL, modUtil.g_ADOConn, adOpenDynamic
    End If

    mblnPersist = True

ExitBlock:

    Exit Sub

ErrorBlock:

    Err.Raise vbObjectError, TypeName(Me) & ".OpenTable", "Data Engine Misbehaved on OpenTable, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description
    Resume ExitBlock

End Sub

Property Let Recordset(rs As ADODB.Recordset)
    If rsTemp Is Nothing Then Exit Property
    Recordset = rs
End Property

Property Get Recordset() As ADODB.Recordset
    If rsTemp Is Nothing Then Exit Property
    Set Recordset = rsTemp
End Property

Public Sub Initialise(Optional TheRecordSet As ADODB.Recordset)
    If Not TheRecordSet Is Nothing Then
        Set rsTemp = TheRecordSet
    Else
        OpenTable
    End If
End Sub

Friend Sub CloseTable()
  On Error Resume Next
  If Not rsTemp Is Nothing Then
      rsTemp.Close
      Set rsTemp = Nothing
      mblnPersist = False
  End If
End Sub

Public Sub Terminate()
  On Error Resume Next
  CloseTable
End Sub

Property Let Filter(New_Filter As String)
    If rsTemp Is Nothing Then Exit Property
    rsTemp.Filter = New_Filter
End Property

Property Get Filter() As String
    If rsTemp Is Nothing Then Exit Property
    Filter = rsTemp.Filter
End Property

Public Sub ReQuery(Optional blnReLoad As Boolean)
    If rsTemp Is Nothing Then Exit Sub
    rsTemp.ReQuery
    If blnReLoad Then MoveFirst
End Sub

Public Sub ReLoad()
    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.BOF And rsTemp.EOF Then Exit Sub
    RefreshMembers
End Sub

Property Get RecordCount() As Long
    If rsTemp Is Nothing Then Exit Property
    If rsTemp.BOF And rsTemp.EOF Then Exit Property
    On Error Resume Next    'Ignore if no current record
    RecordCount = rsTemp.RecordCount
End Property

Property Get TableObject() As ADODB.Recordset
    If rsTemp Is Nothing Then Exit Property
    Set TableObject = rsTemp
End Property

Property Get TableSource() As String
    If rsTemp Is Nothing Then Exit Property
    TableSource = rsTemp.Source
End Property

Property Get TableName() As String
    If rsTemp Is Nothing Then Exit Property
    TableName = cstTableName
End Property

Property Let TableIndex(strIndexName As String)
    rsTemp.Index = strIndexName
End Property

Property Get TableIndex() As String
    If rsTemp Is Nothing Then Exit Property
    TableIndex = rsTemp.Index
End Property

Property Get NoMatch() As Boolean
    NoMatch = mblnNoMatch
End Property

Property Let NoMatch(New_Param As Boolean)
    mblnNoMatch = New_Param
End Property

Public Function BOF() As Boolean
    If Not rsTemp Is Nothing Then
        BOF = rsTemp.BOF
    Else
        BOF = True
    End If
End Function

Public Function EOF() As Boolean
    If Not rsTemp Is Nothing Then
        EOF = rsTemp.EOF
    Else
        EOF = True
    End If
End Function

Public Sub MoveFirst(Optional blnNoRefresh As Boolean)
    If rsTemp Is Nothing Then OpenTable
    If rsTemp.BOF And rsTemp.EOF Then Exit Sub
    rsTemp.MoveFirst
    If Not blnNoRefresh Then RefreshMembers
End Sub

Public Sub MoveLast(Optional blnNoRefresh As Boolean)
    If rsTemp Is Nothing Then OpenTable
    If rsTemp.BOF And rsTemp.EOF Then Exit Sub
    rsTemp.MoveLast
    If Not blnNoRefresh Then RefreshMembers
End Sub

Public Sub MovePrevious(Optional blnNoRefresh As Boolean)
    If rsTemp Is Nothing Then Exit Sub
    If Not rsTemp.BOF Then
        rsTemp.MovePrevious
        If Not blnNoRefresh Then RefreshMembers
    End If
End Sub

Public Sub MoveNext(Optional blnNoRefresh As Boolean)
    If rsTemp Is Nothing Then Exit Sub
    If Not rsTemp.EOF Then
        rsTemp.MoveNext
        If Not blnNoRefresh Then RefreshMembers
    End If
End Sub

Public Sub Enumerate(Optional blnReverse As Boolean)

    If rsTemp Is Nothing Then OpenTable

    If Not blnReverse Then
        MoveFirst
        RaiseEvent Enumerate
        Do Until EOF
            MoveNext
            If EOF Then Exit Sub
            RaiseEvent Enumerate
        Loop
    Else
        MoveLast
        RaiseEvent Enumerate
        Do Until BOF
            MovePrevious
            If BOF Then Exit Sub
            RaiseEvent Enumerate
        Loop
    End If

End Sub


