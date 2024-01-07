Imports System.Data
Imports System.IO
Imports System.Net.Mail
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Data.SQLite

' Documentation

'NUGET PACKAGES NEEDED
' FOR SQLITE FUNCTIONALITY: System.Data.SQLite
' FOR JSON FUNCTIONALITY: Newtonsoft.JSON

' To use this set of extensions, either copy this file into your project and add the needed NuGet packages,
' Or build this into a class library, and reference it within your project as a project reference.
' If this is inside another project, you may need to import it in your code by writing 'imports Extensions'.
' It may be helpful to declare a namespace so you can import it as 'imports mynamespace.Extensions'


Public Module Extensions

    '<summary>
    ' Removes text between two points, wherever they're found 
    '</summary>
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

    '<summary>
    'Formats a date for insertion into a MS SQL SERVER DATE / DATETIME column.
    '</summary>
    <Extension()>
    Public Function FormatDateForSQL(inputDate As DateTime, Optional format As String = "yyyy-MM-dd HH:mm:ss") As String
        Return inputDate.ToString(format)
    End Function

    '<summary>
    ' Escapes line breaks in a string into \r\n tags, as needed for some API calls.
    '</summary>
    <Extension()>
    Public Function EscapeLineBreaksToRN(text As String) As String
        Return text.Replace(vbCrLf, "\r\n")
    End Function

    '<summary>
    ' Unescapes \r\n tags into real line breaks
    '</summary>
    <Extension()>
    Public Function UnEscapeLineBreaksFromRN(text As String) As String
        Return text.Replace("\r\n", vbCrLf)
    End Function

    '<summary>
    ' Unescapes \r\n into HTML <br /> tags for line breaks in a web page
    '</summary>
    <Extension()>
    Public Function UnescapeRNToHtmlLineBreaks(text As String) As String
        Return text.Replace("\r\n", "<br />")
    End Function

    '<summary>
    ' Not likely useful for most people; Replaces an existing separator (comma? Semicolon, etc) with a set of new separator tags around it. Maybe useful for putting XML tags around something.
    '</summary>
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

    '<summary>
    ' Exports a DataTable object (result from sql database query, for example) into a CSV file. 
    '</summary>
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

    '<summary>
    'Formats a DataTable object (result from a SQL Server query result, for example) into HTML for output on a web page
    '</summary>
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

    '<summary>
    ' Convert an io.stream into text using reader.ReadToEnd(). Convenience / user would expect a functioning .ToString().
    '</summary>
    <Extension()>
    Public Function ToString(stream As IO.Stream) As String
        Using reader As New StreamReader(stream, Encoding.UTF8)
            Return reader.ReadToEnd()
        End Using


    End Function
    '<summary>
    'Returns the first word of a string
    '</summary>
    <Extension()>
    Public Function FirstWord(str As String) As String
        Return str.Split(" ").First
    End Function
    '<summary>
    'Turns a DataRow collection into a List( of Strings)
    '</summary>
    <Extension()>
    Public Function ToStringList(Rows As DataRowCollection, ColumnNameForStringList As String) As List(Of String)
        Dim listicle As New List(Of String)
        For Each row As DataRow In Rows
            listicle.Add(row.GetValueNotNull(ColumnNameForStringList))
        Next
        Return listicle
    End Function

    '<summary>
    'Sends an email
    '</summary>
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

    '<summary>
    ' Gets a value from a SQL datarow column; returns an empty string if it's null, preventing null reference errors

    '</summary>
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

    '<summary>
    'Check quickly if a string contains another string, as well as checking whether it's case sensitive or not
    '</summary>
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

    '<summary>
    'Replace XML/HTML < and > tags with &gt; and &lt; tags
    '</summary>
    <Extension()>
    Public Function CommentOutXML(text As String) As String
        Return text.Replace("<", "&lt;").Replace(">", "&gt;")
    End Function

    '<summary>
    'Converts a string holding XML into an XmlElement object
    '</summary>
    <Extension()>
    Public Function ToXmlElement(text As String) As XmlElement
        Dim doc As New XmlDocument()
        doc.LoadXml(text)

        Return doc.DocumentElement
    End Function

    ''<summary>
    ''Executes a query in a SQLite database. THIS STRING, BEFORE THE PERIOD, SHOULD BE THE FILE PATH OF THE SQLITE DB. Query as the parameter after it.
    ''</summary>
    '<Extension()>
    'Public Sub ExecuteSingleNonQuery(dbPath As String, dbQuery As String)
    '    Dim databaseCon As New SQLiteDatabase(dbPath)

    '    databaseCon.ExecuteNonQuery(dbQuery)
    'End Sub

    '<summary>
    'Breaks out XML in a string, into proper line breaks and indentation for easy viewing.
    '</summary>
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

    '<summary>
    'Returns a List(Of Integers) holding the index positions wherever a substring is found. 
    '</summary>
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
        Return indexes
    End Function

    '<summary>
    ' Returns the text after the absolute last instance a given character or string. I.e after a \ for the file name of a file path.
    '</summary>
    <Extension()>
    Public Function TextAfterRecursive(text As String, TextToBeginFrom As String) As String
        Dim newtext As String = text
        Do Until Not newtext.Contains(TextToBeginFrom)
            newtext = newtext.TextAfter(TextToBeginFrom)
        Loop
        Return newtext
    End Function

    '<Extension()>
    'Public Function HasItems(listitems As ListItemCollection) As Boolean
    '    If listitems.Count <> 0 Then
    '        Return True
    '    Else
    '        Return False

    '    End If
    'End Function

    '<summary>
    'Formats a string for text to be escaped for proper saving into a file name
    '</summary>
    <Extension()>
    Public Function FormatForFileName(text As String) As String
        Return text.Remove("/").Remove("\").Remove(":").Remove(" ").Remove(",").Remove("""").Remove("*")
    End Function

    '<summary>
    'Converts a boolean into a 0 or 1, for saving into a boolean field in sql. 
    '</summary>
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

    '<summary>
    'Appends text to an existing string, including a newline after the already existing text
    '</summary>
    <Extension()>
    Public Sub WriteLine(ByRef Text As String, AppendedText As String)

        If Text = "" Then
            Text = AppendedText

        Else
            Text += vbCrLf & AppendedText
        End If

    End Sub

    '<summary>
    'Escapes single quotes in a string that would break SQL syntax
    '</summary>
    <Extension()>
    Public Function FormatForSQLStringData(UnEscapedText As String) As String
        If UnEscapedText IsNot Nothing Then
            Return UnEscapedText.Replace("'", "''").Remove(vbCrLf)

        Else
            Return ""

        End If


    End Function
    '<summary>
    'Un escapes single quotes in a sql string.
    '</summary>
    <Extension>
    Public Function UnFormatForSQLStringData(EscapedText As String) As String
        If EscapedText IsNot Nothing Then
            Return EscapedText.Replace("''", "'")
        Else
            Return ""
        End If
    End Function

    '<Extension>
    '    Public Sub Minimize(window As System.Windows.Window)

    '        window.WindowState = FormWindowState.Minimized

    '    End Sub

    '<Extension>
    '    Public Sub ExtendXAxis(window As System.Windows.Window, AdditionalXAxisPixels As Integer)

    '        window.Width += AdditionalXAxisPixels

    '    End Sub

    '<Extension>
    '    Public Function GetLastItemsText(lvItemCollection As ListViewItemCollection) As String
    '        Dim indexoflastitem As Integer

    '        indexoflastitem = lvItemCollection.Count - 1 'Count is 1-based, or starts at 1. Index is 0-based, or starts at 0. Subtract 1 to make up difference
    '        Return lvItemCollection.Item(indexoflastitem).Text

    '    End Function


    '<summary>
    ' This is for prefixing a value received from an OAuth token endpoint for authenticating, as a Bearer token.
    ' Often you get just the access token itself, whereas it's expected you send it back as 'Bearer ' and then the token..
    ' This checks if the Bearer tag is present, and if not, prefixes it
    '</summary>
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
    '<summary>
    'Splits a string by another string of text, returns an array. 
    '</summary>
    <Extension()>
    Public Function SplitStringByString(Text As String, Delimiter As String) As String()
        ' Splits a string, by another string
        Return Regex.Split(Text, Delimiter)

    End Function
    '<summary>
    'Prefixes a string with some new text behind / at the beginning of it. 
    '</summary>
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

    '<summary>
    ' Returns the text occurring after a specific string, as opposed to needing to get text after a given index using substring()
    '</summary>
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

    '<summary>
    ' Returns the text occurring before a specific string
    '</summary>
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

    '<summary>
    ' Returns the text between two points... Test and review this code, it doesn't look right. 
    '</summary>
    '<Extension()>
    '    Public Function Substring(Text As String, StartingPoint As String) As String
    '        Dim IndexOfStartingPoint As Integer

    '        IndexOfStartingPoint = Text.IndexOf(StartingPoint)
    '        Try
    '            Return Text.Substring(IndexOfStartingPoint + StartingPoint.Length)
    '        Catch
    '            Return ""
    '        End Try
    '    End Function


    '<summary>
    ' returns the text between two specified substrings, using strings instead of indexes. 
    '</summary>
    <Extension()>
    Public Function TextBetween(Text As String, strAfter As String, strBefore As String) As String

        Return Text.TextAfter(strAfter).TextBefore(strBefore)

    End Function

    '<summary>
    'Remove a string wherever it's found from an existing string
    '</summary>
    <Extension()> Public Function Remove(OriginalText As String, TextToBeRemoved As String) As String
        Return OriginalText.Replace(TextToBeRemoved, "")


    End Function



End Module
