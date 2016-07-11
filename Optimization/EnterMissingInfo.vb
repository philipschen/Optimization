Public Class EnterMissingInfo
    Public Property des As ArrayList
    Public Property id As ArrayList
    Public Property color As String
    Public Property part1 As String
    Public Property saw As String
    Public Property size1 As String

    Private Sub ChooseStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Easy Cut V1.0"
        Me.ControlBox = False
        AcceptButton = Button1
        'CancelButton = Button2
        TextBox3.Visible = False
        ListBox1.Items.Add("SAW 1")
        ListBox1.Items.Add("SAW 2")
        ListBox1.Items.Add("SAW 3")
        ListBox1.Items.Add("SAW 4")
        ListBox1.Items.Add("SAW 5")
        ListBox1.Items.Add("SAW 6")
        ListBox1.Items.Add("DON'T KNOW")

        ListBox2.Items.Add("180")
        ListBox2.Items.Add("200")
        ListBox2.Items.Add("216")
        ListBox2.Items.Add("230")
        ListBox2.Items.Add("240")
        ListBox2.Items.Add("Other")
        ListBox2.SelectedIndex = 0

        ListBox3.Items.Add("BRONZE")
        ListBox3.Items.Add("WHITE")
        ListBox3.Items.Add("SILVER")
        ListBox3.Items.Add("BLACK")
        ListBox3.Items.Add("OTHER")
        ListBox1.Items.Add("DON'T KNOW")

        TextBox1.Text = part1
        TextBox2.Text = "F201 = BRONZE" + vbCrLf + "F202 = WHITE" + vbCrLf + "FSP2 = OTHER"
    End Sub
    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        If ListBox2.SelectedItem = "OTHER" Then
            TextBox3.Visible = True
        Else
            TextBox3.Visible = False
        End If
    End Sub
    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListBox2.SelectedValue >= 0 Then
            If ListBox1.SelectedIndex >= 0 And ListBox1.SelectedItem <> "DON'T KNOW" Then
                saw = ListBox1.Text
            Else
                saw = ""
            End If

            If ListBox2.SelectedIndex >= 0 And ListBox2.SelectedItem <> "OTHER" Then
                size1 = ListBox2.Text
            Else
                size1 = TextBox3.Text
            End If

            If ListBox3.SelectedIndex >= 0 And ListBox1.SelectedItem <> "DON'T KNOW" Then
                color = ListBox3.Text
            Else
                color = ""
            End If

            Me.DialogResult = DialogResult.OK
            Me.Close()
        Else
            MessageBox.Show("Enter A Value for Size")
        End If
    End Sub
End Class