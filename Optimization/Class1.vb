Imports System.Data.SqlClient
Public Class Class1
    Public Property connect1 As String = "Data Source=(Local)\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"
    Public Function cutkey() As String
        Dim cutkey1 As String = ""
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=(Local)\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT CutGLib FROM activationCode"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
        While readerObj.Read
            cutkey1 = readerObj("CutGLib").ToString
        End While

        Return cutkey1
    End Function

    Public Property version As String = "Easy Cut V1.1"
End Class