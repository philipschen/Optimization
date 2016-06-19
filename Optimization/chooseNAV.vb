Public Class chooseNAV
    Dim connectionstring As Class1 = New Class1
    Public Property choice As Integer
    Private Sub chooseNAV_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
        Me.ControlBox = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        choice = 0
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        choice = 1
        Me.Close()
    End Sub
End Class