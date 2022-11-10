<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TextBoxMessage
	Inherits System.Windows.Forms.Form

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
		Me.TxtLog = New System.Windows.Forms.RichTextBox()
		Me.SuspendLayout()
		'
		'TxtLog
		'
		Me.TxtLog.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
			Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.TxtLog.Location = New System.Drawing.Point(12, 12)
		Me.TxtLog.Name = "TxtLog"
		Me.TxtLog.Size = New System.Drawing.Size(421, 322)
		Me.TxtLog.TabIndex = 0
		Me.TxtLog.Text = ""
		'
		'TextBoxMessage
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(445, 346)
		Me.Controls.Add(Me.TxtLog)
		Me.Name = "TextBoxMessage"
		Me.Text = "Message"
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents TxtLog As System.Windows.Forms.RichTextBox
End Class
