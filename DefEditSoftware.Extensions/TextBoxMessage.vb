Public Class TextBoxMessage

	Public Sub New(Message As String, Optional Title As String = "Message")


		' This call is required by the designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call.
		TxtLog.Text = Message
		Me.Text = Title
	End Sub

End Class