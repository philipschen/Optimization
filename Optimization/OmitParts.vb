Public Class OmitParts
    Dim connectionstring As Class1 = New Class1
    Public Property des As ArrayList
    Public Property omitted As ArrayList = New ArrayList

    Private Sub OmitParts_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
        Me.ControlBox = False
        For it = 0 To des.Count - 1
            ListBox1.Items.Add(des(it))
        Next
        ListBox1.SelectedIndex = 0
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ListBox1.SelectedIndex >= 0 Then
            ListBox2.Items.Add(ListBox1.SelectedItem)
            omitted.Add(ListBox1.SelectedItem)
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox2.SelectedIndex >= 0 Then
            ListBox1.Items.Add(ListBox2.SelectedItem)
            omitted.Remove(ListBox2.SelectedItem)
            ListBox2.Items.RemoveAt(ListBox2.SelectedIndex)
        End If
    End Sub


End Class