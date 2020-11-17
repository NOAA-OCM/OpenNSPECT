Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsLCClassData_NET.clsLCClassData")> Public Class clsLCClassData
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
	
	Const cstTableName As String = "LCClass"
	Const cstDefaultQuery As String = "SELECT * FROM LCClass WHERE "
	
	Private rsTemp As ADODB.Recordset
	
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
	Private mstrSQL As String
	Private mblnIsDirty As Boolean
	Private mblnIsSynchronised As Boolean
	Private mblnPersist As Boolean
	Private mblnNoMatch As Boolean
	
	'local variable(s)
	Private mlngLCClassID As Integer
	Private mlngValue As Integer
	Private mstrName As String
	Private mlngLCTypeID As Integer
	Private sngCNA As Single
	Private sngCNB As Single
	Private sngCNC As Single
	Private sngCND As Single
	Private sngCoverFactor As Single
	Private mlngW_WL As Integer
	
	'Object States: Not required to be persistent
	Public ReadOnly Property IsDirty() As Boolean
		Get
			IsDirty = mblnIsDirty
		End Get
	End Property
	
	
	Public Property IsSynchronised() As Boolean
		Get
			IsSynchronised = mblnIsSynchronised
		End Get
		Set(ByVal Value As Boolean)
			mblnIsSynchronised = Value
			If mblnIsSynchronised Then
				RaiseEvent Synchronised()
			End If
		End Set
	End Property
	
	
	Public Property Persist() As Boolean
		Get
			Persist = mblnPersist
		End Get
		Set(ByVal Value As Boolean)
			mblnPersist = Value
			If mblnPersist = True Then
				If rsTemp Is Nothing Then OpenTable()
			End If
		End Set
	End Property
	
	'Persistent Object properties
	
	Public Property LCClassID() As Integer
		Get
			LCClassID = mlngLCClassID
		End Get
		Set(ByVal Value As Integer)
			mlngLCClassID = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property Value() As Integer
		Get
			Value = mlngValue
		End Get
		Set(ByVal Value As Integer)
			mlngValue = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property Name() As String
		Get
			Name = mstrName
		End Get
		Set(ByVal Value As String)
			mstrName = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property LCTypeID() As Integer
		Get
			LCTypeID = mlngLCTypeID
		End Get
		Set(ByVal Value As Integer)
			mlngLCTypeID = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property CNA() As Single
		Get
			CNA = sngCNA
		End Get
		Set(ByVal Value As Single)
			sngCNA = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property CNB() As Single
		Get
			CNB = sngCNB
		End Get
		Set(ByVal Value As Single)
			sngCNB = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property CNC() As Single
		Get
			CNC = sngCNC
		End Get
		Set(ByVal Value As Single)
			sngCNC = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property CND() As Single
		Get
			CND = sngCND
		End Get
		Set(ByVal Value As Single)
			sngCND = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property CoverFactor() As Single
		Get
			CoverFactor = sngCoverFactor
		End Get
		Set(ByVal Value As Single)
			sngCoverFactor = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property W_WL() As Integer
		Get
			W_WL = mlngW_WL
		End Get
		Set(ByVal Value As Integer)
			mlngW_WL = Value
			mblnIsDirty = True
		End Set
	End Property
	
	
	Public Property SQL() As String
		Get
			SQL = mstrSQL
		End Get
		Set(ByVal Value As String)
			mstrSQL = Value
			'Set new record source using SQL statements
			If mstrSQL = "" Then Exit Property
			If Not rsTemp Is Nothing Then CloseTable()
			OpenTable(mstrSQL)
		End Set
	End Property
	
	
	Property Recordset() As ADODB.Recordset
		Get
			If rsTemp Is Nothing Then Exit Property
			Recordset = rsTemp
		End Get
		Set(ByVal Value As ADODB.Recordset)
			If rsTemp Is Nothing Then Exit Property
			Recordset.Fields = Value.Fields
		End Set
	End Property
	
	
	'UPGRADE_NOTE: Filter was upgraded to Filter_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Property Filter_Renamed() As String
		Get
			If rsTemp Is Nothing Then Exit Property
			'UPGRADE_WARNING: Couldn't resolve default property of object rsTemp.Filter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Filter_Renamed = rsTemp.Filter
		End Get
		Set(ByVal Value As String)
			If rsTemp Is Nothing Then Exit Property
			rsTemp.Filter = Value
		End Set
	End Property
	
	ReadOnly Property RecordCount() As Integer
		Get
			If rsTemp Is Nothing Then Exit Property
			If rsTemp.BOF And rsTemp.EOF Then Exit Property
			On Error Resume Next 'Ignore if no current record
			RecordCount = rsTemp.RecordCount
		End Get
	End Property
	
	ReadOnly Property TableObject() As ADODB.Recordset
		Get
			If rsTemp Is Nothing Then Exit Property
			TableObject = rsTemp
		End Get
	End Property
	
	ReadOnly Property TableSource() As String
		Get
			If rsTemp Is Nothing Then Exit Property
			'UPGRADE_WARNING: Couldn't resolve default property of object rsTemp.Source. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TableSource = rsTemp.Source
		End Get
	End Property
	
	ReadOnly Property TableName() As String
		Get
			If rsTemp Is Nothing Then Exit Property
			TableName = cstTableName
		End Get
	End Property
	
	
	Property TableIndex() As String
		Get
			If rsTemp Is Nothing Then Exit Property
			TableIndex = rsTemp.Index
		End Get
		Set(ByVal Value As String)
			rsTemp.Index = Value
		End Set
	End Property
	
	
	Property NoMatch() As Boolean
		Get
			NoMatch = mblnNoMatch
		End Get
		Set(ByVal Value As Boolean)
			mblnNoMatch = Value
		End Set
	End Property
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		mblnIsDirty = False
		mblnIsSynchronised = False
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		On Error Resume Next
		Terminate()
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Sub Connect()
		OpenTable()
		MoveFirst()
	End Sub
	
	Public Sub DisConnect()
		CloseTable()
	End Sub
	
	Public Function Load(Optional ByRef lngLCClassID As Object = Nothing) As Boolean
		
		On Error GoTo ErrorBlock
		
		If rsTemp Is Nothing Then OpenTable()
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lngLCClassID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If IsNothing(lngLCClassID) Then lngLCClassID = mlngLCClassID
		
		If rsTemp.BOF And rsTemp.EOF Then Exit Function
		rsTemp.MoveFirst()
		'UPGRADE_WARNING: Couldn't resolve default property of object lngLCClassID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		rsTemp.Find("LCClassID = " & lngLCClassID)
		NoMatch = EOF_Renamed()
		
		If Not NoMatch Then
			RefreshMembers()
			Load = True
			IsSynchronised = True
			mblnIsDirty = False
			RaiseEvent Loaded()
		Else
			RaiseEvent NotFound()
			Load = False
			IsSynchronised = False
		End If
		
ExitBlock: 
		
		Exit Function
		
ErrorBlock: 
		
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Err.Raise(vbObjectError, TypeName(Me) & ".Load", "Data Engine Misbehaved on Load, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description)
		Resume ExitBlock
		
	End Function
	
	'UPGRADE_NOTE: Exists was upgraded to Exists_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function Exists_Renamed(Optional ByRef lngLCClassID As Object = Nothing) As Boolean
		
		On Error GoTo ErrorBlock
		
		Dim tbTemp As ADODB.Recordset
		Dim lngRecords As Integer
		
		If rsTemp Is Nothing Then OpenTable()
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object lngLCClassID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If IsNothing(lngLCClassID) Then lngLCClassID = mlngLCClassID
		'UPGRADE_WARNING: Couldn't resolve default property of object lngLCClassID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tbTemp = g_ADOConn.Execute("SELECT COUNT (LCClassID) FROM LCClass WHERE LCClassID = " & lngLCClassID)
		lngRecords = Val(tbTemp.GetString)
		
		If lngRecords = 1 Then
			NoMatch = False
			RaiseEvent Exists()
		Else
			NoMatch = True
		End If
		
		Exists_Renamed = Not NoMatch
		
ExitBlock: 
		
		On Error Resume Next
		tbTemp.Close()
		'UPGRADE_NOTE: Object tbTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		tbTemp = Nothing
		Exit Function
		
ErrorBlock: 
		
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Err.Raise(vbObjectError, TypeName(Me) & ".Exists", "Data Engine Misbehaved on Exists, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description)
		Resume ExitBlock
		
	End Function
	
	Private Sub RefreshMembers()
		
		On Error Resume Next
		
		If rsTemp.BOF Or rsTemp.EOF Then Exit Sub
		
		With rsTemp
			
			mlngLCClassID = .Fields("LCClassID").Value
			mlngValue = .Fields("Value").Value
			mstrName = .Fields("Name").Value
			mlngLCTypeID = .Fields("LCTypeID").Value
			sngCNA = .Fields("CN-A").Value
			sngCNB = .Fields("CN-B").Value
			sngCNC = .Fields("CN-C").Value
			sngCND = .Fields("CN-D").Value
			sngCoverFactor = .Fields("CoverFactor").Value
			mlngW_WL = .Fields("W_WL").Value
			
		End With
		
		IsSynchronised = True
		RaiseEvent Refresh()
		
	End Sub
	
	Public Sub AddNew()
		
		On Error GoTo ErrorBlock
		
		If rsTemp Is Nothing Then OpenTable()
		
		rsTemp.AddNew()
		SaveMembers()
		
		rsTemp.Update()
		mblnIsDirty = False
		RaiseEvent Added()
		
		g_intLCClassid = rsTemp.Fields("LCClassID").Value
		
ExitBlock: 
		
		Exit Sub
		
ErrorBlock: 
		
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Err.Raise(vbObjectError, TypeName(Me) & ".AddNew", "Data Engine Misbehaved on AddNew, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description)
		Resume ExitBlock
		
	End Sub
	
	Public Sub Delete()
		
		On Error GoTo ErrorBlock
		
		If rsTemp Is Nothing Then OpenTable()
		
		rsTemp.Delete(ADODB.AffectEnum.adAffectCurrent)
		'SaveMembers
		
		rsTemp.Update()
		mblnIsDirty = False
		
ExitBlock: 
		
		Exit Sub
		
ErrorBlock: 
		
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Err.Raise(vbObjectError, TypeName(Me) & ".AddNew", "Data Engine Misbehaved on AddNew, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description)
		Resume ExitBlock
		
	End Sub
	
	Public Sub SaveChanges()
		
		On Error GoTo ErrorBlock
		
		Dim varBookmark As Object
		
		If Not mblnIsSynchronised Then
			'Cannot Save Changes to object without first being loaded or synchronised
			RaiseEvent NotSynchronised()
		Else
			If mblnIsDirty Then
				SaveMembers()
				rsTemp.Update()
				mblnIsDirty = False
				RaiseEvent Saved()
			End If
		End If
		
ExitBlock: 
		
		On Error Resume Next
		Exit Sub
		
ErrorBlock: 
		
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Err.Raise(vbObjectError, TypeName(Me) & ".SaveChanges", "Data Engine Misbehaved on SaveChanges, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description)
		Resume ExitBlock
		
	End Sub
	
	Private Sub SaveMembers()
		
		On Error Resume Next
		
		With rsTemp
			
			.Fields("Value").Value = mlngValue
			.Fields("Name").Value = mstrName
			.Fields("LCTypeID").Value = mlngLCTypeID
			.Fields("CN-A").Value = sngCNA
			.Fields("CN-B").Value = sngCNB
			.Fields("CN-C").Value = sngCNC
			.Fields("CN-D").Value = sngCND
			.Fields("CoverFactor").Value = sngCoverFactor
			.Fields("W_WL").Value = mlngW_WL
			.Fields("LCTypeID").Value = mlngLCTypeID
		End With
		
		IsSynchronised = True
		RaiseEvent SavedMembers()
		
	End Sub
	
	Private Sub CopyProperties(ByRef objCopy As clsLCClassData)
		
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
		
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Err.Raise(vbObjectError, TypeName(Me) & ".CopyProperties", "Data Engine Misbehaved on CopyProperties, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description)
		Resume ExitBlock
		
	End Sub
	
	Public Function CopyObject(ByVal objCopy As clsLCClassData) As clsLCClassData
		If Not objCopy Is Nothing Then
			CopyProperties(objCopy)
		End If
	End Function
	
	Friend Sub OpenTable(Optional ByRef strSQL As String = "")
		
		On Error GoTo ErrorBlock
		
		If rsTemp Is Nothing Then
			rsTemp = New ADODB.Recordset
			rsTemp.CursorType = ADODB.CursorTypeEnum.adOpenDynamic
			rsTemp.LockType = ADODB.LockTypeEnum.adLockOptimistic
		End If
		
		If strSQL = "" Then
			rsTemp.Open("LCClass", modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic)
		Else
			rsTemp.Open(strSQL, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic)
		End If
		
		mblnPersist = True
		
ExitBlock: 
		
		Exit Sub
		
ErrorBlock: 
		
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Err.Raise(vbObjectError, TypeName(Me) & ".OpenTable", "Data Engine Misbehaved on OpenTable, Error Number: [" & Err.Number & "]" & vbCrLf & "Description: " & Err.Description)
		Resume ExitBlock
		
	End Sub
	
	Public Sub Initialise(Optional ByRef TheRecordSet As ADODB.Recordset = Nothing)
		If Not TheRecordSet Is Nothing Then
			rsTemp = TheRecordSet
		Else
			OpenTable()
		End If
	End Sub
	
	Friend Sub CloseTable()
		On Error Resume Next
		If Not rsTemp Is Nothing Then
			rsTemp.Close()
			'UPGRADE_NOTE: Object rsTemp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			rsTemp = Nothing
			mblnPersist = False
		End If
	End Sub
	
	Public Sub Terminate()
		On Error Resume Next
		CloseTable()
	End Sub
	
	Public Sub ReQuery(Optional ByRef blnReLoad As Boolean = False)
		If rsTemp Is Nothing Then Exit Sub
		rsTemp.ReQuery()
		If blnReLoad Then MoveFirst()
	End Sub
	
	Public Sub ReLoad()
		If rsTemp Is Nothing Then Exit Sub
		If rsTemp.BOF And rsTemp.EOF Then Exit Sub
		RefreshMembers()
	End Sub
	
	Public Function BOF() As Boolean
		If Not rsTemp Is Nothing Then
			BOF = rsTemp.BOF
		Else
			BOF = True
		End If
	End Function
	
	'UPGRADE_NOTE: EOF was upgraded to EOF_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Function EOF_Renamed() As Boolean
		If Not rsTemp Is Nothing Then
			EOF_Renamed = rsTemp.EOF
		Else
			EOF_Renamed = True
		End If
	End Function
	
	Public Sub MoveFirst(Optional ByRef blnNoRefresh As Boolean = False)
		If rsTemp Is Nothing Then OpenTable()
		If rsTemp.BOF And rsTemp.EOF Then Exit Sub
		rsTemp.MoveFirst()
		If Not blnNoRefresh Then RefreshMembers()
	End Sub
	
	Public Sub MoveLast(Optional ByRef blnNoRefresh As Boolean = False)
		If rsTemp Is Nothing Then OpenTable()
		If rsTemp.BOF And rsTemp.EOF Then Exit Sub
		rsTemp.MoveLast()
		If Not blnNoRefresh Then RefreshMembers()
	End Sub
	
	Public Sub MovePrevious(Optional ByRef blnNoRefresh As Boolean = False)
		If rsTemp Is Nothing Then Exit Sub
		If Not rsTemp.BOF Then
			rsTemp.MovePrevious()
			If Not blnNoRefresh Then RefreshMembers()
		End If
	End Sub
	
	Public Sub MoveNext(Optional ByRef blnNoRefresh As Boolean = False)
		If rsTemp Is Nothing Then Exit Sub
		If Not rsTemp.EOF Then
			rsTemp.MoveNext()
			If Not blnNoRefresh Then RefreshMembers()
		End If
	End Sub
	
	'UPGRADE_NOTE: Enumerate was upgraded to Enumerate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Enumerate_Renamed(Optional ByRef blnReverse As Boolean = False)
		
		If rsTemp Is Nothing Then OpenTable()
		
		If Not blnReverse Then
			MoveFirst()
			RaiseEvent Enumerate()
			Do Until EOF_Renamed()
				MoveNext()
				If EOF_Renamed() Then Exit Sub
				RaiseEvent Enumerate()
			Loop 
		Else
			MoveLast()
			RaiseEvent Enumerate()
			Do Until BOF
				MovePrevious()
				If BOF Then Exit Sub
				RaiseEvent Enumerate()
			Loop 
		End If
		
	End Sub
End Class