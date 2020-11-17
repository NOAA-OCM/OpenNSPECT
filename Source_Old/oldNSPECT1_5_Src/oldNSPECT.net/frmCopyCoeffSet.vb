Option Strict Off
Option Explicit On
Friend Class frmCopyCoeffSet
	Inherits System.Windows.Forms.Form
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
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		If modUtil.UniqueName("CoefficientSet", (txtCoeffSetName.Text)) And Trim(txtCoeffSetName.Text) <> "" Then
			If g_boolCopyCoeff Then
				frmPollutants.CopyCoefficient((txtCoeffSetName.Text), (cboCoeffSet.Text))
			Else
				frmNewPollutants.CopyCoefficient((txtCoeffSetName.Text), (cboCoeffSet.Text))
			End If
		Else
			MsgBox("The name you have choosen for coefficient set is already in use.  Please pick another.", MsgBoxStyle.Critical, "Name In Use")
			With txtCoeffSetName
				.SelectionStart = 0
				.SelectionLength = Len(txtCoeffSetName.Text)
				.Focus()
			End With
			
		End If
		
	End Sub
	
	Public Sub init(ByRef rsCoeffSet As ADODB.Recordset)
		'The form is passed a recordest containing the names of all coefficient sets, allows for
		'easier populating
		
		Dim i As Short
		
		rsCoeffSet.MoveFirst()
		
		For i = 0 To rsCoeffSet.RecordCount - 1
			cboCoeffSet.Items.Add(rsCoeffSet.Fields("Name").Value)
			rsCoeffSet.MoveNext()
		Next i
		
	End Sub
End Class