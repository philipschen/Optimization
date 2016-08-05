Imports System.Data.SqlClient

Public Class InitialSetup
    Dim connectionstring As Class1 = New Class1
    Private Sub InitialSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Text = "Before running this you must make sure that microsoft SQL 2014 is installed and a server named SQLEXPRESS is created." + vbCrLf + vbCrLf + "More detailed instructions can be found in the InstallationGuide and ReadMe"
        TextBox1.Text = "https://www.microsoft.com/en-us/download/details.aspx?id=42299"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim str As String
        '"Data Source=(Local)\SQLEXPRESS;Initial Catalog=TestDB;Integrated Security=True"
        Dim myConn As SqlConnection = New SqlConnection("Server=(local)\SQLEXPRESS;Integrated Security=True;database=master")
        str = "Create Database TestDB"
        Dim myCommand As SqlCommand = New SqlCommand(str, myConn)

        Try
            myConn.Open()
            myCommand.ExecuteNonQuery()
            Dim result1 As DialogResult = MessageBox.Show("Database is created successfully",
                        "MyProgram", MessageBoxButtons.OK,
                         MessageBoxIcon.Information)
            If result1 = DialogResult.OK Then
            End If

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