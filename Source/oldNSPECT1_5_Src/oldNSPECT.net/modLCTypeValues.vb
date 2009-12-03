Option Strict Off
Option Explicit On
Module modLCTypeValues
	' *************************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  modLCTypeValues
	' *************************************************************************************************
	' *  Description: Code mod for handling validation of landcover type
	' *
	' *         Function CheckGridValuesLCtype::  checks the cover factor grid values
	' *         Function Rollback::  Delete's previous work if user so chooses.
	' *
	' *  Called By:  frmLCTypes, frmNewLCtypes
	' **************************************************************************************************
	
	Public Function CheckGridValuesLCType(ByRef aryValue() As Object) As Boolean
		
		Dim i As Short
		i = 0
		
		For i = 0 To UBound(aryValue)
			'Select Case i
			Select Case i
				
				Case 6
					'UPGRADE_WARNING: Couldn't resolve default property of object aryValue(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If aryValue(i) < 0 Or aryValue(i) > 1 Then
						MsgBox("Cover Factor must be between 0 and 1.  Please check values", MsgBoxStyle.Critical, "Check Cover Number")
						CheckGridValuesLCType = False
					Else
						CheckGridValuesLCType = True
					End If
					
			End Select
			
		Next i
		
	End Function
	
	Public Sub RollBack(ByRef strName As String)
		
		Dim strSQLSelect As String
		strSQLSelect = "SELECT * FROM LCTYPE where NAME LIKE '" & strName & "'"
		
		Dim rsRollback As ADODB.Recordset
		rsRollback = New ADODB.Recordset
		
		rsRollback.Open(strSQLSelect, modUtil.g_ADOConn, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
		rsRollback.Delete(ADODB.AffectEnum.adAffectCurrent)
		
		'clean
		'UPGRADE_NOTE: Object rsRollback may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		rsRollback = Nothing
		
	End Sub
End Module