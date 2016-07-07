Imports System.Data.SqlClient
Public Class NewExtrusions
    Private Sub NewExtrusions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"

        ListBox1.Items.Add("Bronze")
        ListBox1.Items.Add("White")
        ListBox1.Items.Add("Silver")
        ListBox1.Items.Add("Other")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        If (ListBox1.SelectedIndex >= 0) Then
            Dim stockID1 As String = ""
            Dim stockID2 As String = ""
            Dim stockID3 As String = ""

            stockID1 = TextBox1.Text
            stockID2 = TextBox2.Text
            stockID3 = TextBox3.Text
            Dim description As String = TextBox4.Text
            Dim stockColor As String = ListBox1.SelectedItems(0)
            Dim stockSize As String = TextBox5.Text
            Dim stockCount As String = TextBox6.Text
            Dim vUpdate As Boolean = True
            Dim internalID As ArrayList = New ArrayList

            cmd.CommandText = "SELECT internalID  FROM stockNew"
            cmd.ExecuteNonQuery()
            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            'This will loop through all returned records 
            While readerObj.Read
                Dim temp1 As Integer = Convert.ToInt64(readerObj("internalID").ToString)
                internalID.Add(temp1)
            End While
            readerObj.Close()
            If (String.Equals(stockID2, "") Or String.Equals(description, "") Or String.Equals(stockSize, "") Or String.Equals(stockCount, "")) Then
                Dim warning1 As finishInput = New finishInput()
                warning1.ShowDialog()
            Else
                Dim rn As New Random
                Dim it1 As Integer
                Dim boolinternalID As Boolean = True
                Dim inputinternal As Integer

                inputinternal = rn.Next(10000, 99999)
                For it1 = 0 To internalID.Count - 1
                    If internalID(it1) = inputinternal Then
                        it1 = 0
                        inputinternal = rn.Next(10000, 99999)
                    End If
                Next
                Dim temp1 = Convert.ToString(inputinternal)

                cmd.CommandText = "INSERT INTO stockNew VALUES('" + stockID1.ToUpper() + "', '" + stockID2.ToUpper() + "' , '" + stockID3.ToUpper() + "', '" + description.ToUpper() + "' , '" + stockColor.ToUpper() + "', " + stockSize.ToUpper() + " , " + stockCount.ToUpper() + ", " + temp1.ToUpper() + ", '' , 0, '')"
                cmd.ExecuteNonQuery()
                Me.Close()
            End If
        ElseIf ListBox1.SelectedIndex < 0 Then
            Dim warning1 As inputColor = New inputColor()
            warning1.ShowDialog()
        End If



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

End Class