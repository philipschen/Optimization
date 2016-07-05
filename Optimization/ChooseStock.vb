Public Class ChooseStock
    Public Property des As ArrayList
    Public Property id As ArrayList
    Public Property color As ArrayList
    Public Property part1 As String

    Private Sub ChooseStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AcceptButton = Button1
        CancelButton = Button2
        For it = 0 To des.Count - 1
            ListBox1.Items.Add(color(it) + " " + des(it) + " " + id(it))
        Next
        ListBox1.SelectedIndex = 0
        TextBox1.Text = part1
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