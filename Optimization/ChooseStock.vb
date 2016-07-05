Public Class ChooseStock
    Public Property des As ArrayList
    Public Property id As ArrayList
    Public Property color As ArrayList
    Public Property part1 As String
    Public Property saw As ArrayList

    Private Sub ChooseStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Easy Cut V1.0"
        Me.ControlBox = False
        AcceptButton = Button1
        CancelButton = Button2
        For it = 0 To des.Count - 1
            ListBox1.Items.Add(color(it) + " " + des(it) + " " + id(it) + " Saw: " + saw(it))
        Next
        ListBox1.SelectedIndex = 0
        TextBox1.Text = part1
        TextBox2.Text = "F201 = BRONZE" + vbCrLf + "F202 = WHITE" + vbCrLf + "FSP2 = OTHER"
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If saw(ListBox1.SelectedIndex) = "" Then
            Dim frm1 As inputSaw = New inputSaw()
            If frm1.ShowDialog() = DialogResult.OK Then
                saw.RemoveAt(ListBox1.SelectedIndex)
                saw.Insert(ListBox1.SelectedIndex, frm1.ListBox1.Text)
            End If
        End If
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub


End Class