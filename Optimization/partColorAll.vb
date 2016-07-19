Public Class partColorAll
    Dim connectionstring As Class1 = New Class1
    Private Sub partColorAll_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
        Me.ControlBox = False
        AcceptButton = Button1
        CancelButton = Button2
        ListBox1.Items.Add("BRONZE")
        ListBox1.Items.Add("WHITE")
        ListBox1.Items.Add("SILVER")
        ListBox1.Items.Add("BLACK")
        ListBox1.Items.Add("OTHER")
        Button3.Visible = False
        ListBox1.Visible = False
        Label1.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Button3.Visible = True
        ListBox1.Visible = True
        Label1.Visible = True
        Button2.Visible = False
        Button1.Visible = False


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class