Imports System.Data.SqlClient
Public Class Form1
    Dim connectionstring As Class1 = New Class1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"
        Button3.Visible = True
        PictureBox1.Image = Image.FromFile("temp.jpg")
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

    End Sub

    Private Sub BStock_Click(sender As Object, e As EventArgs) Handles BStock.Click
        Dim ext1 As Extrusions = New Extrusions()
        ext1.Show()
    End Sub

    Private Sub BPart_Click(sender As Object, e As EventArgs) Handles BPart.Click
        Dim cutman1 As CutManagement = New CutManagement()
        cutman1.Show()
    End Sub

    Private Sub set_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim result1 As DialogResult = MessageBox.Show("Are you certain you want to delete all tables?", "Confirm Deletion", MessageBoxButtons.YesNo)
        If result1 = DialogResult.Yes Then
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = connectionstring.connect1
            con.Open()
            cmd.Connection = con
            cmd.CommandText = "DELETE FROM stockNew"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "DELETE FROM stockUsed"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "DELETE FROM parts"
            cmd.ExecuteNonQuery()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim tables1 As DirectlyEdit = New DirectlyEdit()
        tables1.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim excellin1 As StockInputExcell = New StockInputExcell()
        excellin1.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim excellin1 As PartInputExcell = New PartInputExcell()
        excellin1.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim excellin1 As UsedStockInputExcel = New UsedStockInputExcel()
        excellin1.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim excellin1 As NavInput = New NavInput()
        excellin1.Show()
    End Sub
End Class

