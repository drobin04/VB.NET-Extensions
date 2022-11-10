


Imports System.Security.Cryptography
Imports System.Text

Public Module ExtensionModule


    Public Sub DisplayError(message As String)
        Dim msg As New TextBoxMessage(message)
        msg.Show()

    End Sub

    Public Sub DisplayError(ex As Exception)
        Dim msg As New TextBoxMessage(ex.ToString)
        msg.Show()

    End Sub

    Public Function DeleteFileIfExists(filePath As String) As Boolean
        Dim test As Boolean = False
        Try
            GC.Collect()
            GC.WaitForPendingFinalizers()

            If IO.File.Exists(filePath) Then IO.File.Delete(filePath)

            test = True
        Catch
            test = False
        End Try

        Return test
    End Function


    Public Function CalculateMD5Hash(ByVal input As String) As String
        'Dim md5 As MD5 = System.Security.Cryptography.MD5.Create()
        'Dim inputBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(input)
        'Dim hash As Byte() = md5.ComputeHash(inputBytes)
        'Dim sb As StringBuilder = New StringBuilder()

        'For i As Integer = 0 To hash.Length - 1
        '    sb.Append(hash(i).ToString("X2"))
        'Next

        'Return sb.ToString()

        Dim md5 As MD5 = MD5.Create()
        Dim hashData As Byte() = md5.ComputeHash(Encoding.[Default].GetBytes(input))
        Dim returnValue As StringBuilder = New StringBuilder()

        For i As Integer = 0 To hashData.Length - 1
            returnValue.Append(hashData(i).ToString())
        Next

        Return returnValue.ToString()

    End Function

End Module
