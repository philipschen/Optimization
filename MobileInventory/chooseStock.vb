﻿Public Class chooseStock
    Public Property des As ArrayList
    Public Property id As ArrayList
    Public Property part1 As String

    Private Sub chooseStock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Easy Cut V1.0"
        Me.ControlBox = False
        AcceptButton = Button1
        CancelButton = Button2
        For it = 0 To des.Count - 1
            ListBox1.Items.Add(des(it) + " " + id(it))
        Next
        ListBox1.SelectedIndex = 0
        TextBox1.Text = part1
        TextBox2.Text = "F201 = BRONZE" + vbCrLf + "F202 = WHITE" + vbCrLf + "FSP2 = OTHER"
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