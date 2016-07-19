Public Class partColor
    Dim connectionstring As Class1 = New Class1
    Private Sub partColor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
        Me.ControlBox = False
        AcceptButton = Button1
        CancelButton = Button2
        ListBox1.Items.Add("BRONZE")
        ListBox1.Items.Add("WHITE")
        ListBox1.Items.Add("SILVER")
        ListBox1.Items.Add("BLACK")
        ListBox1.Items.Add("OTHER")
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

End Class