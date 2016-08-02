Public Class Confirm1
    Dim connectionstring As Class1 = New Class1
    Private Sub Confirm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
        Me.ControlBox = False
        AcceptButton = Button1
        CancelButton = Button2
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If String.Equals(TextBox1.Text.ToUpper, "YES") Then
            DialogResult = DialogResult.OK
            Me.Close()
        Else
            TextBox1.Text = ""
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class