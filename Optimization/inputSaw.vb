Public Class inputSaw
    Dim connectionstring As Class1 = New Class1
    Private Sub inputSaw_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
        Me.ControlBox = False
        AcceptButton = Button1
        CancelButton = Button2

        ListBox1.Items.Add("SAW 1")
        ListBox1.Items.Add("SAW 2")
        ListBox1.Items.Add("SAW 3")
        ListBox1.Items.Add("SAW 4")
        ListBox1.Items.Add("SAW 5")
        ListBox1.Items.Add("SAW 6")
        ListBox1.SelectedIndex = 0
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