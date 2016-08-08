Imports System.Data.SqlClient

Public Class CutGLibActivation
    Dim connectionstring As Class1 = New Class1
    Private Sub CutGLibActivation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Text = "1. Go to the following URL and Download CutGLib." + vbCrLf + "2. Run The .EXE File from it." + vbCrLf + "3. Launch The Get License Application and follow the instructions." + vbCrLf + vbCrLf + "More detailed instructions found in the InstallationGuide and ReadME"
        TextBox1.Text = "http://www.optimalon.com/CutGLib_install.htm"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Or TextBox2.Text Is Nothing Then
            MessageBox.Show("Please copy the activation code into the box.")
        Else
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = connectionstring.connect1
            con.Open()
            cmd.Connection = con

            cmd.CommandText = "DELETE FROM activationCode"
            cmd.ExecuteNonQuery()

            cmd.CommandText = "INSERT into activationCode Values('" + TextBox2.Text + "')"
            cmd.ExecuteNonQuery()
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.No
        Me.Close()
    End Sub
End Class