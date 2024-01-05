Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports System.Windows.Forms.ListView
Imports System.Web.UI.WebControls
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Net.Mail
Imports System.Web.UI.HtmlControls
Imports Newtonsoft.Json.Linq

Namespace DefExtensions
    Public Module Extensions
        <Extension()>
        Public Function RemoveBetweenRecursive(OriginalText As String, TextBeforestr As String, TextAfterstr As String, KeyToEnsureProperIdentificationUsingContains As String) As String
            If OriginalText.Contains(KeyToEnsureProperIdentificationUsingContains) = False Or OriginalText.Contains(TextAfterstr) = False Then
                Throw New ArgumentException("Text to remove could not be found.")
            End If


            Do While OriginalText.Contains(KeyToEnsureProperIdentificationUsingContains) And OriginalText.Contains(TextAfterstr)


                OriginalText = OriginalText.TextBefore(TextBeforestr) & OriginalText.TextAfter(TextAfterstr)

            Loop

            Return OriginalText
        End Function


        <Extension()>
        Public Sub ShowDivLightBox(ByRef div As HtmlGenericControl)
            div.Style.Remove("display")
            div.Style.Add("display", "block")
        End Sub

        <Extension()>
        Public Sub HideDivLightBox(ByRef div As HtmlGenericControl)
            div.Style.Remove("display")
            div.Style.Add("display", "none")
        End Sub

        <Extension()>
        Public Function FormatDateForSQL(inputDate As DateTime, Optional format As String = "yyyy-MM-dd HH:mm:ss") As String
            Return inputDate.ToString(format)
        End Function

        <Extension()>
        Public Function EscapeLineBreaksToRN(text As String) As String
            Return text.Replace(vbCrLf, "\r\n")
        End Function

        <Extension()>
        Public Function UnEscapeLineBreaksFromRN(text As String) As String
            Return text.Replace("\r\n", vbCrLf)
        End Function

        <Extension()>
        Public Function UnescapeRNToHtmlLineBreaks(text As String) As String
            Return text.Replace("\r\n", "<br>")
        End Function

        <Extension()>
        Function wrapValue(value As String, group As String, separator As String) As String
            If value.Contains(separator) Then
                If value.Contains(group) Then
                    value = value.Replace(group, group + group)
                End If
                value = group & value & group
            End If
            Return value
        End Function

        <Extension()>
        Function ExportToCSV(dtable As DataTable, fileName As String) As Boolean
            Dim result As Boolean = True
            Try
                Dim sb As New StringBuilder()
                Dim separator As String = ","
                Dim group As String = """"
                Dim newLine As String = Environment.NewLine

                For Each column As DataColumn In dtable.Columns
                    sb.Append(wrapValue(column.ColumnName, group, separator) & separator)
                Next
                ' here you could add the column for the username
                sb.Append(newLine)

                For Each row As DataRow In dtable.Rows
                    For Each col As DataColumn In dtable.Columns
                        sb.Append(wrapValue(row(col).ToString(), group, separator) & separator)
                    Next
                    ' here you could extract the password for the username
                    sb.Append(newLine)
                Next
                Using fs As New StreamWriter(fileName)
                    fs.Write(sb.ToString())
                End Using
            Catch ex As Exception
                result = False
            End Try
            Return result
        End Function

        <Extension()>
        Public Function FormatTableFromQuery(dataTable As DataTable, Optional recid As Boolean = False) As String
            Dim innerHTML As String = ""
            Dim firstCol As String = "<tr>"
            Dim RecIDColumnIndex As Integer = -1
            Dim recidcolumncounter As Integer = 0
            For Each col As DataColumn In dataTable.Columns
                If recid = False And col.ColumnName = "RecID" Then
                    RecIDColumnIndex = recidcolumncounter
                    Continue For
                Else
                    recidcolumncounter += 1

                End If

                firstCol += $"<th>{col.ColumnName}</th>"
            Next

            firstCol += "</tr>"

            For Each row As DataRow In dataTable.Rows
                Dim count As Integer = 0

                innerHTML += "<tr>"

                'If recid = False Then
                '    count = 1
                'End If

                While count < dataTable.Columns.Count
                    If count = RecIDColumnIndex Then
                        Continue While

                    Else
                        innerHTML += $"<td>{row(count)}</td>"

                        count += 1
                    End If

                End While

                innerHTML += "</tr>"
            Next

            Dim outerHTML As String = $"{firstCol}{innerHTML}"

            Return $"<table cellspacing='0' class='htmlLit' style='width: 100%'>{outerHTML}</table>"
        End Function

        <Extension()>
        Public Function TryString(Text As String) As String
            Try
                Return Text
            Catch ex As Exception
                Return ""
            End Try
        End Function

        <Extension()>
        Public Function ToString(stream As IO.Stream) As String
            Using reader As New StreamReader(stream, Encoding.UTF8)
                Return reader.ReadToEnd()
            End Using


        End Function

        <Extension()>
        Public Function FirstWord(str As String) As String
            Return str.Split(" ").First
        End Function

        <Extension()>
        Public Function ToStringList(Rows As DataRowCollection, ColumnNameForStringList As String) As List(Of String)
            Dim listicle As New List(Of String)
            For Each row As DataRow In Rows
                listicle.Add(row.GetValueNotNull(ColumnNameForStringList))
            Next
            Return listicle
        End Function


        <Extension()>
        Public Sub SendExternalEmail(SendTo As String, Subject As String, Body As String, Host As String, User As String, Pass As String, Optional html As Boolean = False)
            Try
                Dim eMail As MailMessage = New MailMessage(User, SendTo, Subject, Body)
                Dim eMailHost As SmtpClient = New SmtpClient(Host, 25)

                eMailHost.UseDefaultCredentials = False
                eMailHost.Credentials = New System.Net.NetworkCredential(User, Pass)
                eMail.IsBodyHtml = html
                eMailHost.EnableSsl = True
                eMailHost.Send(eMail)
            Catch ex As Exception
            End Try
        End Sub

        <Extension()>
        Public Function GetValueNotNull(row As DataRow, ColumnName As String) As String
            Try
                If row(ColumnName) Is Nothing Then
                    Return ""
                Else
                    Return row(ColumnName).ToString
                End If
            Catch ex As Exception
                Return ""
            End Try
        End Function

        <Extension()>
        Public Function Contains(text As String, containedtext As String, Optional CaseSensitive As Boolean = False) As Boolean
            Dim str As String = text
            Dim contained As String = containedtext
            If CaseSensitive = False Then
                str = str.ToLower
                contained = contained.ToLower
            End If
            Return str.Contains(contained)
        End Function
        <Extension()>
        Public Function HasValue(text As String) As Boolean
            Try
                If text <> "" Then
                    Return True
                Else
                    Return False
                End If
            Catch
                Return False
            End Try
        End Function

        <Extension()>
        Public Function CommentOutXML(text As String) As String
            Return text.Replace("<", "&lt;").Replace(">", "&gt;")
        End Function

        <Extension()>
        Public Sub TransmitFileToUser(Filepath As String, FileDisplayName As String)
            System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream"
            System.Web.HttpContext.Current.Response.AddHeader("Content-Disposition", String.Format("attachment; filename=""{0}""", FileDisplayName))
            Try
                FileSystem.FileClose(Filepath)
            Catch ex As Exception

            End Try

            System.Web.HttpContext.Current.Response.TransmitFile(Filepath)
            Try
                System.Web.HttpContext.Current.Response.[End]()
            Catch ex As System.Threading.ThreadAbortException

            End Try
        End Sub

        <Extension()>
        Public Function ToXmlElement(text As String) As XmlElement
            Dim doc As New XmlDocument()
            doc.LoadXml(text)

            Return doc.DocumentElement
        End Function

        <Extension()>
        Public Sub ExecuteSingleNonQuery(dbPath As String, dbQuery As String)
            Dim databaseCon As New SQLiteDatabase(dbPath)

            databaseCon.ExecuteNonQuery(dbQuery)
        End Sub

        <Extension()>
        Public Function PrettyPrintXML(ByVal xml As String) As String
            Dim result As String = ""
            Dim mStream As MemoryStream = New MemoryStream()
            Dim writer As XmlTextWriter = New XmlTextWriter(mStream, Encoding.Unicode)
            Dim document As XmlDocument = New XmlDocument()

            Try
                document.LoadXml(xml)
                writer.Formatting = System.Xml.Formatting.Indented
                document.WriteContentTo(writer)
                writer.Flush()
                mStream.Flush()
                mStream.Position = 0
                Dim sReader As StreamReader = New StreamReader(mStream)
                Dim formattedXml As String = sReader.ReadToEnd()
                result = formattedXml
            Catch __unusedXmlException1__ As XmlException
            End Try
            Try
                mStream.Close()
                writer.Close()
            Catch ex As Exception

            End Try

            Return result
            'Credit to S M Kamran @ Stack Overflow
            'https://stackoverflow.com/questions/1123718/format-xml-string-to-print-friendly-xml-string
        End Function

        <Extension()>
        Public Sub AutoResizeHeightByNumberOfItems(ByRef lstbox As System.Web.UI.WebControls.ListBox)
            lstbox.Height = (lstbox.Items.Count * 17) + 8
        End Sub

        <Extension()>
        Public Function FindAllIndexesOf(ByVal str As String, ByVal value As String) As List(Of Integer)
            If String.IsNullOrEmpty(value) Then Throw New ArgumentException("the string to find may not be empty", "value")
            Dim indexes As List(Of Integer) = New List(Of Integer)()
            Dim index As Integer = 0

            While True
                index = str.IndexOf(value, index)
                If index = -1 Then Return indexes
                indexes.Add(index)
                index += value.Length
            End While
        End Function

        <Extension()>
        Public Function TextAfterRecursive(text As String, TextToBeginFrom As String) As String
            Dim newtext As String = text
            Do Until Not newtext.Contains(TextToBeginFrom)
                newtext = newtext.TextAfter(TextToBeginFrom)
            Loop
            Return newtext
        End Function

        <Extension()>
        Public Function HasItems(listitems As ListItemCollection) As Boolean
            If listitems.Count <> 0 Then
                Return True
            Else
                Return False

            End If
        End Function

        <Extension()>
        Public Function FormatForFileName(text As String) As String
            Return text.Remove("/").Remove("\").Remove(":").Remove(" ").Remove(",").Remove("""").Remove("*")
        End Function

        <Extension()>
        Public Function FormatBoolForSQL(ByRef bool As Boolean) As Integer

            Select Case bool
                Case True
                    Return 1
                Case False
                    Return 0
                Case Else
                    Return 0

            End Select
        End Function


        <Extension()>
        Public Sub WriteLine(ByRef Text As String, AppendedText As String)

            If Text = "" Then
                Text = AppendedText

            Else
                Text += vbCrLf & AppendedText
            End If

        End Sub

        <Extension()>
        Public Function FormatForSQLStringData(UnEscapedText As String) As String
            If UnEscapedText IsNot Nothing Then
                Return UnEscapedText.Replace("'", "''").Remove(vbCrLf)

            Else
                Return ""

            End If


        End Function

        <Extension>
        Public Function UnFormatForSQLStringData(EscapedText As String) As String
            If EscapedText IsNot Nothing Then
                Return EscapedText.Replace("''", "'")
            Else
                Return ""
            End If
        End Function

        <Extension()>
        Public Function GetFieldListFromBusObSchema(BusObSchemaString As String) As List(Of BusObFieldSchema)
            Dim schema As BusObSchema

            schema = JsonConvert.DeserializeObject(Of BusObSchema)(BusObSchemaString.Replace("readOnly", "readOnly2"))





            Return schema.fieldDefinitions

        End Function

        <Extension>
        Public Sub Minimize(window As System.Windows.Window)

            window.WindowState = FormWindowState.Minimized

        End Sub

        <Extension>
        Public Sub ExtendXAxis(window As System.Windows.Window, AdditionalXAxisPixels As Integer)

            window.Width += AdditionalXAxisPixels

        End Sub

        <Extension>
        Public Function GetLastItemsText(lvItemCollection As ListViewItemCollection) As String
            Dim indexoflastitem As Integer

            indexoflastitem = lvItemCollection.Count - 1 'Count is 1-based, or starts at 1. Index is 0-based, or starts at 0. Subtract 1 to make up difference
            Return lvItemCollection.Item(indexoflastitem).Text

        End Function

        <Extension>
        Public Function PrefixAsBearerToken(AccessToken As String) As String
            Dim newaccesstoken As String
            If AccessToken.Contains("Bearer") Then
                newaccesstoken = AccessToken

            Else
                newaccesstoken = "Bearer " & AccessToken

            End If


            Return newaccesstoken

        End Function

        <Extension()>
        Public Function SplitStringByString(Text As String, Delimiter As String) As String()
            Return Regex.Split(Text, Delimiter)

        End Function

        <Extension()>
        Public Function GetCherwellFieldName(CherwellFieldJSONFromAPI As String) As String
            Dim CherwellFieldName As String = CherwellFieldJSONFromAPI.TextAfter("""name"": ").TextBefore(",")

            Return CherwellFieldName.Replace("""", "")
        End Function

        <Extension()>
        Public Function GetCherwellFieldValue(CherwellFieldJSONFromAPI As String) As String
            Dim CherwellFieldName As String = CherwellFieldJSONFromAPI.TextAfter("""value"": ").TextBefore("}")

            Return CherwellFieldName.Replace("""", "")
        End Function

        <Extension()>
        Public Function PrefixWithText(Text As String, Prefix As String) As String


            Try
                If Text.Contains(Prefix) Then
                    Return Text
                Else
                    Return Prefix & Text
                End If
            Catch
                Return ""
            End Try
        End Function

        <Extension()>
        Public Function TextAfter(Text As String, StartingPoint As String) As String
            Dim IndexOfStartingPoint As Integer

            IndexOfStartingPoint = Text.IndexOf(StartingPoint)
            Try
                Return Text.Substring(IndexOfStartingPoint + StartingPoint.Length)
            Catch
                Return ""
            End Try
        End Function

        <Extension()>
        Public Function TextBefore(Text As String, TextDelimiter As String) As String
            Dim IndexOfDelimiter As Integer = 0
            Dim TextBeforeDelimiter As String = ""
            Try
                If Text.Contains(TextDelimiter) Then
                    IndexOfDelimiter = Text.IndexOf(TextDelimiter)
                    TextBeforeDelimiter = Text.Substring(0, IndexOfDelimiter)

                Else
                    TextBeforeDelimiter = Text

                End If

            Catch

            End Try

            Return TextBeforeDelimiter
        End Function

        <Extension()>
        Public Function Substring(Text As String, StartingPoint As String) As String
            Dim IndexOfStartingPoint As Integer

            IndexOfStartingPoint = Text.IndexOf(StartingPoint)
            Try
                Return Text.Substring(IndexOfStartingPoint + StartingPoint.Length)
            Catch
                Return ""
            End Try
        End Function

        <Extension()> Public Function TextBetween(Text As String, strAfter As String, strBefore As String) As String

            Return Text.TextAfter(strAfter).TextBefore(strBefore)

        End Function

        <Extension()> Public Function GetFieldListFromCherwellAPIBusinessObject(BusinessObjectSearchResultJSON As String) As String
            Dim FieldList As String = ""

            FieldList = BusinessObjectSearchResultJSON.TextBetween("""fields"":[", "],""links"":")

            Return FieldList
        End Function
        <Extension()>
        Public Function GetBusObSummariesFromArray(BusinessObjectSummaries As String) As BusObSummary()
            Dim Summary() As BusObSummary

            Summary = JsonConvert.DeserializeObject(Of BusObSummary())(BusinessObjectSummaries)

            Return Summary
        End Function

        <Extension()>
        Public Function GetBusObSummaryFromJSON(BusObSummaryJSON As String) As BusObSummary

            Dim Summary As BusObSummary

            Summary = JsonConvert.DeserializeObject(Of BusObSummary)(BusObSummaryJSON)

            Return Summary
        End Function

        <Extension()>
        Public Function GetBusObSummaryArrayFromJSON(SummaryJSON As String) As BusObSummary()

            Return SummaryJSON.GetBusObSummariesFromArray()

        End Function

        <Extension()> Public Sub LogToEventLog(LogText As String, Optional Source As String = "Def-Tools Class Library")
            Try
                Dim Logger As New EventLogger
                Logger.EventLogger(LogText, Source)
            Catch ex As Exception
                Console.Write(ex.ToString)

            End Try

        End Sub

        <Extension()> Public Function Remove(OriginalText As String, TextToBeRemoved As String) As String
            Return OriginalText.Replace(TextToBeRemoved, "")


        End Function



    End Module

End Namespace
