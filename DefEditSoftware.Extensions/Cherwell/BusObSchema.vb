Imports DefEditSoftware.Extensions.DefExtensions


Public Class BusObSchema

	Public BusObID As String

	Public fieldDefinitions As New List(Of BusObFieldSchema)



End Class

Public Class BusObFieldSchema


	Public autoFill As String
	Public calculated As Boolean
	Public category As String
	Public decimalDigits As Integer
	Public description As String
	Public details As String
	Public displayName As String
	Public enabled As Boolean
	Public Property fieldId As String
		Get
			Return _fieldidvalue
		End Get
		Set(value As String)
			_fieldidvalue = value.Remove("BO:")


		End Set
	End Property
	Public _fieldidvalue As String


	Public hasDate As Boolean
	Public hasTime As Boolean
	Public isFullTextSearchable As Boolean
	Public maximumSize As String
	Public name As String
	Public ReadOnly2 As Boolean
	Public required As Boolean
	Public type As String
	Public typeLocalized As String
	Public validated As Boolean
	Public wholeDigits As Integer


End Class