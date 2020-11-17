<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PollCoeffSelectionShapefileForm
    Inherits OpenNspect.BaseDialogForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cbxSelPollutantPolysOnly = New System.Windows.Forms.CheckBox()
        Me.diaPollSFOpen = New System.Windows.Forms.OpenFileDialog()
        Me.txtPollSF = New System.Windows.Forms.TextBox()
        Me.txtPollSFBrowse = New System.Windows.Forms.Button()
        Me.btnSelPollPoly = New System.Windows.Forms.Button()
        Me.txtNumSelected = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmboPollAttrib = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'cbxSelPollutantPolysOnly
        '
        Me.cbxSelPollutantPolysOnly.AutoSize = True
        Me.cbxSelPollutantPolysOnly.Location = New System.Drawing.Point(12, 67)
        Me.cbxSelPollutantPolysOnly.Name = "cbxSelPollutantPolysOnly"
        Me.cbxSelPollutantPolysOnly.Size = New System.Drawing.Size(155, 17)
        Me.cbxSelPollutantPolysOnly.TabIndex = 0
        Me.cbxSelPollutantPolysOnly.Text = "Use selected polygons only"
        Me.cbxSelPollutantPolysOnly.UseVisualStyleBackColor = True
        Me.cbxSelPollutantPolysOnly.Visible = False
        '
        'diaPollSFOpen
        '
        Me.diaPollSFOpen.FileName = "OpenFileDialog1"
        '
        'txtPollSF
        '
        Me.txtPollSF.Location = New System.Drawing.Point(12, 25)
        Me.txtPollSF.Name = "txtPollSF"
        Me.txtPollSF.Size = New System.Drawing.Size(388, 20)
        Me.txtPollSF.TabIndex = 1
        '
        'txtPollSFBrowse
        '
        Me.txtPollSFBrowse.Location = New System.Drawing.Point(417, 25)
        Me.txtPollSFBrowse.Name = "txtPollSFBrowse"
        Me.txtPollSFBrowse.Size = New System.Drawing.Size(75, 23)
        Me.txtPollSFBrowse.TabIndex = 2
        Me.txtPollSFBrowse.Text = "Browse"
        Me.txtPollSFBrowse.UseVisualStyleBackColor = True
        '
        'btnSelPollPoly
        '
        Me.btnSelPollPoly.Location = New System.Drawing.Point(415, 63)
        Me.btnSelPollPoly.Name = "btnSelPollPoly"
        Me.btnSelPollPoly.Size = New System.Drawing.Size(75, 23)
        Me.btnSelPollPoly.TabIndex = 3
        Me.btnSelPollPoly.Text = "Select"
        Me.btnSelPollPoly.UseVisualStyleBackColor = True
        Me.btnSelPollPoly.Visible = False
        '
        'txtNumSelected
        '
        Me.txtNumSelected.Enabled = False
        Me.txtNumSelected.Location = New System.Drawing.Point(197, 64)
        Me.txtNumSelected.Name = "txtNumSelected"
        Me.txtNumSelected.Size = New System.Drawing.Size(100, 20)
        Me.txtNumSelected.TabIndex = 5
        Me.txtNumSelected.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 104)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Attribute Field*:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(303, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 13)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Polygons selected"
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(9, 137)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(430, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "* Attribute Field value should be an integer from 1 to 4, representing coefficien" & _
    "t types 1 - 4." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'cmboPollAttrib
        '
        Me.cmboPollAttrib.FormattingEnabled = True
        Me.cmboPollAttrib.Location = New System.Drawing.Point(88, 100)
        Me.cmboPollAttrib.Name = "cmboPollAttrib"
        Me.cmboPollAttrib.Size = New System.Drawing.Size(307, 21)
        Me.cmboPollAttrib.TabIndex = 9
        '
        'PollCoeffSelectionShapefileForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(504, 212)
        Me.Controls.Add(Me.cmboPollAttrib)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNumSelected)
        Me.Controls.Add(Me.btnSelPollPoly)
        Me.Controls.Add(Me.txtPollSFBrowse)
        Me.Controls.Add(Me.txtPollSF)
        Me.Controls.Add(Me.cbxSelPollutantPolysOnly)
        Me.Name = "PollCoeffSelectionShapefileForm"
        Me.Text = "Shapefile for choosing pollutant coefficient"
        Me.Controls.SetChildIndex(Me.cbxSelPollutantPolysOnly, 0)
        Me.Controls.SetChildIndex(Me.txtPollSF, 0)
        Me.Controls.SetChildIndex(Me.txtPollSFBrowse, 0)
        Me.Controls.SetChildIndex(Me.btnSelPollPoly, 0)
        Me.Controls.SetChildIndex(Me.txtNumSelected, 0)
        Me.Controls.SetChildIndex(Me.Label1, 0)
        Me.Controls.SetChildIndex(Me.Label2, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.cmboPollAttrib, 0)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbxSelPollutantPolysOnly As System.Windows.Forms.CheckBox
    Friend WithEvents diaPollSFOpen As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtPollSF As System.Windows.Forms.TextBox
    Friend WithEvents txtPollSFBrowse As System.Windows.Forms.Button
    Friend WithEvents btnSelPollPoly As System.Windows.Forms.Button
    Friend WithEvents txtNumSelected As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmboPollAttrib As System.Windows.Forms.ComboBox
End Class
