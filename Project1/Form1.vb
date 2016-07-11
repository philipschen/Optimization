Imports System.Data.SqlClient
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"
        Button3.Visible = True
        '  PictureBox1.Image = Image.FromFile("C:\Users\Veritias\Documents\Visual Studio 2015\Projects\Optimization\Optimization\test.png")
        '  PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tables1 As NavInput = New NavInput()
        tables1.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Dim excellin1 As StockInputExcell = New StockInputExcell()
        'excellin1.Show()
    End Sub
End Class

