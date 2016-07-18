Public Class Amount
    Public Property amount As Integer

    Private Sub chooseStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Easy Cut V1.0"

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        amount = TextBox1.Text
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub
End Class