Imports System.Data.SqlClient
Public Class Form1
    Dim connectionstring As Class1 = New Class1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = connectionstring.version
        '  PictureBox1.Image = Image.FromFile("test.png")
        '  PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tables1 As NavInput = New NavInput()
        tables1.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim excellin1 As BomInput = New BomInput()
        excellin1.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "DELETE FROM Nav"
        cmd.ExecuteNonQuery()
    End Sub
End Class

