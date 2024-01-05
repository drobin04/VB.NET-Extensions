
Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.SQLite
Imports System.Windows.Forms

Public Class SQLiteDatabase

    Private dbConnection As String

    Public Sub New(ByVal inputFile As String)
        dbConnection = String.Format("Data Source={0}", inputFile)
    End Sub

    Public Sub New(ByVal connectionOpts As Dictionary(Of String, String))
        Dim str As String = ""
        For Each row As KeyValuePair(Of String, String) In connectionOpts
            str += String.Format("{0}={1}; ", row.Key, row.Value)
        Next

        str = str.Trim().Substring(0, str.Length - 1)
        dbConnection = str
    End Sub

    Public Function GetDataTable(ByVal sql As String) As DataTable
        Dim dt As DataTable = New DataTable()
        System.Diagnostics.Debug.WriteLine("Executing: " & sql)
        Try
            Dim cnn As SQLiteConnection = New SQLiteConnection(dbConnection, True)
            cnn.Open()
            Dim mycommand As SQLiteCommand = New SQLiteCommand(cnn)
            mycommand.CommandText = sql
            Dim reader As SQLiteDataReader = mycommand.ExecuteReader()
            dt.Load(reader)
            reader.Close()
            cnn.Close()
        Catch e As Exception
            Throw New Exception(e.Message)
        End Try

        Return dt
    End Function

    Public Function ExecuteNonQuery(ByVal sql As String) As Integer
        Dim cnn As SQLiteConnection = New SQLiteConnection(dbConnection)
        'System.Diagnostics.Debug.WriteLine("Executing: " & sql)
        cnn.Open()
        Dim mycommand As SQLiteCommand = New SQLiteCommand(cnn)
        mycommand.CommandText = sql
        Dim rowsUpdated As Integer = mycommand.ExecuteNonQuery()
        cnn.Close()
        Return rowsUpdated
    End Function

    Public Function ExecuteScalar(ByVal sql As String) As String
        Dim cnn As SQLiteConnection = New SQLiteConnection(dbConnection)
        cnn.Open()
        Dim mycommand As SQLiteCommand = New SQLiteCommand(cnn)
        mycommand.CommandText = sql
        Dim value As Object = mycommand.ExecuteScalar()
        cnn.Close()
        If value IsNot Nothing Then
            Return value.ToString()
        End If

        Return ""
    End Function

    Public Function Update(ByVal tableName As String, ByVal data As Dictionary(Of String, String), ByVal where As String) As Boolean
        Dim vals As String = ""
        Dim returnCode As Boolean = True
        If data.Count >= 1 Then
            For Each val As KeyValuePair(Of String, String) In data
                vals += String.Format(" {0} = '{1}',", val.Key.ToString(), val.Value.ToString())
            Next

            vals = vals.Substring(0, vals.Length - 1)
        End If

        Try
            Me.ExecuteNonQuery(String.Format("update {0} set {1} where {2};", tableName, vals, where))
        Catch
            returnCode = False
        End Try

        Return returnCode
    End Function

    Public Function Delete(ByVal tableName As String, ByVal where As String) As Boolean
        Dim returnCode As Boolean = True
        Try
            Me.ExecuteNonQuery(String.Format("delete from {0} where {1};", tableName, where))
        Catch fail As Exception

            returnCode = False
            Throw
        End Try

        Return returnCode
    End Function

    Public Function Insert(ByVal tableName As String, ByVal data As Dictionary(Of String, String)) As Boolean
        Dim columns As String = ""
        Dim values As String = ""
        Dim returnCode As Boolean = True
        For Each val As KeyValuePair(Of String, String) In data
            columns += String.Format(" {0},", val.Key.ToString())
            values += String.Format(" '{0}',", val.Value)
        Next

        columns = columns.Substring(0, columns.Length - 1)
        values = values.Substring(0, values.Length - 1)
        Try
            Me.ExecuteNonQuery(String.Format("insert into {0}({1}) values({2});", tableName, columns, values))
        Catch fail As Exception

            returnCode = False
            Throw

        End Try

        Return returnCode
    End Function

    Public Function ClearDB() As Boolean
        Dim tables As DataTable
        Try
            tables = Me.GetDataTable("select NAME from SQLITE_MASTER where type='table' order by NAME;")
            For Each table As DataRow In tables.Rows
                Me.ClearTable(table("NAME").ToString())
            Next

            Return True
        Catch
            Return False
        End Try
    End Function

    Public Function ClearTable(ByVal table As String) As Boolean
        Try
            Me.ExecuteNonQuery(String.Format("delete from {0};", table))
            Return True
        Catch
            Return False
        End Try
    End Function
End Class

