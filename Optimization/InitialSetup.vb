Imports System.Data.SqlClient

Public Class InitialSetup
    Dim connectionstring As Class1 = New Class1
    Private Sub InitialSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Text = "Before running this you must make sure that microsoft SQL 2014 is installed and a server named SQLEXPRESS is created."
        TextBox1.Text = "https://www.microsoft.com/en-us/download/details.aspx?id=42299"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim str As String
        '"Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"
        Dim myConn As SqlConnection = New SqlConnection("Server=(local)\SQLEXPRESS;Integrated Security=True;database=master")

        str = "Create Database TestDB"

        Dim myCommand As SqlCommand = New SqlCommand(str, myConn)

        Try
            myConn.Open()
            myCommand.ExecuteNonQuery()
            MessageBox.Show("Database is created successfully",
                        "MyProgram", MessageBoxButtons.OK,
                         MessageBoxIcon.Information)

            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "CREATE TABLE parts" + "(partID nvarchar(50), description nvarchar(50), color nvarchar(50), size FLOAT, count int, internalID int, shopNumber nvarchar(50), itemNumber float, itemQuantity int, context1 nvarchar(50), context2 float, context3 nvarchar(50))"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE TABLE stockNew" + "(stockID1 nvarchar(50), stockID2 nvarchar(50), stockID3 nvarchar(50), description nvarchar(MAX), color nvarchar(50), size FLOAT, count int, internalID int, context1 nvarchar(50), context2 float, context3 nvarchar(50))"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "CREATE TABLE stockUsed" + "(stockID1 nvarchar(50), stockID2 nvarchar(50), stockID3 nvarchar(50), description nvarchar(MAX), color nvarchar(50), size FLOAT, count int, internalID int, location nvarchar(50), context1 nvarchar(50), context2 float, context3 nvarchar(50))"
            cmd.ExecuteNonQuery()
            MessageBox.Show("Tables have been created successfully",
                        "MyProgram", MessageBoxButtons.OK,
                         MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        Finally
            If (myConn.State = ConnectionState.Open) Then
                myConn.Close()
            End If
            Me.Close()
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.No
        Me.Close()
    End Sub
End Class