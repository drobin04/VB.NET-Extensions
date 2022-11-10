
Module ExtensionFunctions

	Public Sub CreateFolderIfNonExistant(FolderPath As String)

		If Not IO.Directory.Exists(FolderPath) Then

			IO.Directory.CreateDirectory(FolderPath)

		End If

	End Sub

End Module
