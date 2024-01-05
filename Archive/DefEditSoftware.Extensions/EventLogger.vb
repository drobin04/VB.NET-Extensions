Public Class EventLogger


	Public Sub EventLogger(strEvent As String, Source As String)
		Dim sSource As String
		Dim sLog As String
		Dim sEvent As String
		Dim sMachine As String

		sSource = Source
		sLog = "Application"
		sEvent = strEvent

		sMachine = "."

		'If Not EventLog.SourceExists(sSource, sMachine) Then
		'EventLog.CreateEventSource(sSource, sLog, sMachine)
		'End If


		If strEvent.Length >= 4000 Then
			strEvent = strEvent.Substring(0, 3950)

		End If


		Dim ELog As New EventLog(sLog, sMachine, sSource)
		'ELog.WriteEntry(sEvent)
		Try
			ELog.WriteEntry(sEvent, EventLogEntryType.Information, 234, CType(3, Short))
		Catch ex As Exception

		End Try


	End Sub

End Class
