Imports System.Data.SqlClient
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"
        Button3.Visible = True
        '  PictureBox1.Image = Image.FromFile("C:\Users\Veritias\Documents\Visual Studio 2015\Projects\Optimization\Optimization\test.png")
        '  PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize

    End Sub
    Private Sub BStock_Click(sender As Object, e As EventArgs) _
   Handles BStock.Click
        Dim frm3 As Form3 = New Form3()
        frm3.Show()
    End Sub

    Private Sub BPart_Click(sender As Object, e As EventArgs) Handles BPart.Click
        Dim frm2 As Form2 = New Form2()
        frm2.Show()
    End Sub

    Private Sub set_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=Parts;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "DELETE FROM Parts"

        cmd.ExecuteNonQuery()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ext1 As Extrusions = New Extrusions()
        ext1.Show()
    End Sub
End Class

