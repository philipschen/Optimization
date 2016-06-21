
Imports System.Data.SqlClient
Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'PartsDataSet.Parts' table. You can move, or remove it, as needed.
        Me.PartsTableAdapter.Fill(Me.PartsDataSet.Parts)
        'TODO: This line of code loads data into the 'PartsDataSet.Parts' table. You can move, or remove it, as needed.
        Me.PartsTableAdapter.Fill(Me.PartsDataSet.Parts)
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"
        CheckBox1.Visible = False
        CheckBox2.Visible = False
        CheckBox3.Visible = False
        CheckBox4.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        TextBox1.Visible = False
        TextBox2.Visible = False
        ListBox1.Visible = False
        ListBox2.Visible = False
        Button3.Visible = False

        CheckBox1.Checked = True

        ListBox2.Items.Add("Bronze")
        ListBox2.Items.Add("White")
        ListBox2.Items.Add("Silver")

        ListBox1.SelectedIndex = 0
        ListBox2.SelectedIndex = 0

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=Parts;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        Dim adp As SqlDataAdapter = New SqlDataAdapter("SELECT stockID, Color, Size, Count FROM dbo.Parts", "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=Parts;Integrated Security=True")
        Dim ds As DataSet = New DataSet()
        adp.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)

    End Sub

    Private Sub Manual_Click(sender As Object, e As EventArgs) _
   Handles Button1.Click
        CheckBox1.Visible = True
        CheckBox2.Visible = True
        CheckBox3.Visible = True
        CheckBox4.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        Label6.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        ListBox1.Visible = True
        ListBox2.Visible = True
        Button3.Visible = True
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ListBox1.Items.Add("THX-01-001")
            ListBox1.Items.Add("THX-01-01F")
            ListBox1.Items.Add("THX-01-01XL")
            ListBox1.Items.Add("THX-01-01N")
            ListBox1.Items.Add("THX-01-002")
            ListBox1.Items.Add("THX-01-02F")
            ListBox1.Items.Add("THX-01-02L")
            ListBox1.Items.Add("THX-01-02N")
            ListBox1.Items.Add("THX-01-02S")
            ListBox1.Items.Add("THX-01-03A")
            ListBox1.Items.Add("THX-01-03AF")
            ListBox1.Items.Add("THX-01-03XLA")
            ListBox1.Items.Add("THX-01-3AN")
            ListBox1.Items.Add("THX-99-02B")
            ListBox1.Items.Add("THX-01-05A")
            ListBox1.Items.Add("THX-01-004")
            ListBox1.Items.Add("THX-99-003")
            ListBox1.Items.Add("THX-01-007A")
            ListBox1.Items.Add("THX-01-06A")
            ListBox1.Items.Add("THX-01-08A")
            ListBox1.Items.Add("THX-01-03AEL")
            ListBox1.Items.Add("THX-01-009")
            ListBox1.Items.Add("THX-99-15")
            ListBox1.Items.Add("THX-99-15A")
            ListBox1.Items.Add("P01-THX001")
            ListBox1.Items.Add("2000-A01-000000")
        End If

        If CheckBox1.Checked = False Then
            ListBox1.Items.Remove("THX-01-001")
            ListBox1.Items.Remove("THX-01-01F")
            ListBox1.Items.Remove("THX-01-01XL")
            ListBox1.Items.Remove("THX-01-01N")
            ListBox1.Items.Remove("THX-01-002")
            ListBox1.Items.Remove("THX-01-02F")
            ListBox1.Items.Remove("THX-01-02L")
            ListBox1.Items.Remove("THX-01-02N")
            ListBox1.Items.Remove("THX-01-02S")
            ListBox1.Items.Remove("THX-01-03A")
            ListBox1.Items.Remove("THX-01-03AF")
            ListBox1.Items.Remove("THX-01-03XLA")
            ListBox1.Items.Remove("THX-01-3AN")
            ListBox1.Items.Remove("THX-99-02B")
            ListBox1.Items.Remove("THX-01-05A")
            ListBox1.Items.Remove("THX-01-004")
            ListBox1.Items.Remove("THX-99-003")
            ListBox1.Items.Remove("THX-01-007A")
            ListBox1.Items.Remove("THX-01-06A")
            ListBox1.Items.Remove("THX-01-08A")
            ListBox1.Items.Remove("THX-01-03AEL")
            ListBox1.Items.Remove("THX-01-009")
            ListBox1.Items.Remove("THX-99-15")
            ListBox1.Items.Remove("THX-99-15A")
            ListBox1.Items.Remove("P01-THX001")
            ListBox1.Items.Remove("2000-A01-000000")
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ListBox1.Items.Add("THX-03A-01")
            ListBox1.Items.Add("THX-03A-01F")
            ListBox1.Items.Add("THX-03A-01XL")
            ListBox1.Items.Add("THX-03A-1N")
            ListBox1.Items.Add("THX-03A-1L")
            ListBox1.Items.Add("THX-03A-04")
            ListBox1.Items.Add("THX-03A-06")
            ListBox1.Items.Add("THX-03A-06A")
            ListBox1.Items.Add("THX-03A-03")
            ListBox1.Items.Add("THX-03A-IFN-X")
            ListBox1.Items.Add("THX-03A-1NC")
            ListBox1.Items.Add("THX-03A-01")
            ListBox1.Items.Add("THX-03A-001")
            ListBox1.Items.Add("THX-03A-003")
        End If

        If CheckBox2.Checked = False Then
            ListBox1.Items.Remove("THX-03A-01")
            ListBox1.Items.Remove("THX-03A-01F")
            ListBox1.Items.Remove("THX-03A-01XL")
            ListBox1.Items.Remove("THX-03A-1N")
            ListBox1.Items.Remove("THX-03A-1L")
            ListBox1.Items.Remove("THX-03A-04")
            ListBox1.Items.Remove("THX-03A-06")
            ListBox1.Items.Remove("THX-03A-06A")
            ListBox1.Items.Remove("THX-03A-03")
            ListBox1.Items.Remove("THX-03A-IFN-X")
            ListBox1.Items.Remove("THX-03A-1NC")
            ListBox1.Items.Remove("THX-03A-01")
            ListBox1.Items.Remove("THX-03A-001")
            ListBox1.Items.Remove("THX-03A-003")
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ListBox1.Items.Add("TDX-12-01")
            ListBox1.Items.Add("TDX-12-02")
            ListBox1.Items.Add("TDX-12-01")
            ListBox1.Items.Add("TDX-12-43")
            ListBox1.Items.Add("TDX-12-06")
            ListBox1.Items.Add("TDX-12-47")
            ListBox1.Items.Add("TDX-12-48A")
            ListBox1.Items.Add("TDX-12-09")
            ListBox1.Items.Add("TDX-12-04")
            ListBox1.Items.Add("TDX-12-05")
            ListBox1.Items.Add("TDX-12-10A")
            ListBox1.Items.Add("TDX-12-A02")
            ListBox1.Items.Add("TDX-12-44")
            ListBox1.Items.Add("TDX-12-A08A")
            ListBox1.Items.Add("SNX-005")
            ListBox1.Items.Add("SNX-005L")
            ListBox1.Items.Add("1240-572-000325")
            ListBox1.Items.Add("TDX-12-A08")
            ListBox1.Items.Add("TDX-12-17")
            ListBox1.Items.Add("TDX-12-11B")
        End If

        If CheckBox3.Checked = False Then
            ListBox1.Items.Remove("TDX-12-01")
            ListBox1.Items.Remove("TDX-12-02")
            ListBox1.Items.Remove("TDX-12-01")
            ListBox1.Items.Remove("TDX-12-43")
            ListBox1.Items.Remove("TDX-12-06")
            ListBox1.Items.Remove("TDX-12-47")
            ListBox1.Items.Remove("TDX-12-48A")
            ListBox1.Items.Remove("TDX-12-09")
            ListBox1.Items.Remove("TDX-12-04")
            ListBox1.Items.Remove("TDX-12-05")
            ListBox1.Items.Remove("TDX-12-10A")
            ListBox1.Items.Remove("TDX-12-A02")
            ListBox1.Items.Remove("TDX-12-44")
            ListBox1.Items.Remove("TDX-12-A08A")
            ListBox1.Items.Remove("SNX-005")
            ListBox1.Items.Remove("SNX-005L")
            ListBox1.Items.Remove("1240-572-000325")
            ListBox1.Items.Remove("TDX-12-A08")
            ListBox1.Items.Remove("TDX-12-17")
            ListBox1.Items.Remove("TDX-12-11B")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            ListBox1.Items.Add("TDX-12-801")
            ListBox1.Items.Add("TDX-12-816")
            ListBox1.Items.Add("TDX-12-804")
            ListBox1.Items.Add("TDX-12-802")
            ListBox1.Items.Add("TDX-12-803")
            ListBox1.Items.Add("TDX-12-809")
            ListBox1.Items.Add("TDX-12-805")
            ListBox1.Items.Add("TDX-12-806B")
            ListBox1.Items.Add("TDX-12-807A")
            ListBox1.Items.Add("TDX-12-808")
            ListBox1.Items.Add("TDX-12-810")
            ListBox1.Items.Add("THX-62-19")
            ListBox1.Items.Add("TDX-12-811")
            ListBox1.Items.Add("TDX-12-814")
            ListBox1.Items.Add("TDX-12-16")
            ListBox1.Items.Add("TDX-12-817")
            ListBox1.Items.Add("1280-ASTR-01")
        End If

        If CheckBox4.Checked = False Then
            ListBox1.Items.Remove("TDX-12-801")
            ListBox1.Items.Remove("TDX-12-816")
            ListBox1.Items.Remove("TDX-12-804")
            ListBox1.Items.Remove("TDX-12-802")
            ListBox1.Items.Remove("TDX-12-803")
            ListBox1.Items.Remove("TDX-12-809")
            ListBox1.Items.Remove("TDX-12-805")
            ListBox1.Items.Remove("TDX-12-806B")
            ListBox1.Items.Remove("TDX-12-807A")
            ListBox1.Items.Remove("TDX-12-808")
            ListBox1.Items.Remove("TDX-12-810")
            ListBox1.Items.Remove("THX-62-19")
            ListBox1.Items.Remove("TDX-12-811")
            ListBox1.Items.Remove("TDX-12-814")
            ListBox1.Items.Remove("TDX-12-16")
            ListBox1.Items.Remove("TDX-12-817")
            ListBox1.Items.Remove("1280-ASTR-01")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) _
   Handles Button3.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=Parts;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        Dim stockID As String = ListBox1.SelectedItems(0)
        Dim stockColor As String = ListBox2.SelectedItems(0)
        Dim stockSize As String = TextBox2.Text
        Dim stockCount As String = TextBox1.Text
        Dim vUpdate As Boolean = True

        cmd.CommandText = "SELECT stockID, Color, Size, Count FROM Parts"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
        'This will loop through all returned records 
        While readerObj.Read
            If ((String.Equals(stockID, readerObj("stockID").ToString)) And (String.Equals(stockColor, readerObj("Color").ToString)) And (String.Equals(stockSize, readerObj("Size").ToString))) Then
                Dim temp1 As Integer = Convert.ToInt32(stockCount)
                Dim temp2 As String = readerObj("Count").ToString
                Dim temp3 As Integer = Convert.ToInt32(temp2)
                temp3 = temp3 + temp1
                temp2 = Convert.ToString(temp3)
                readerObj.Close()
                cmd.CommandText = "UPDATE Parts SET Count = " + temp2 + " WHERE stockID = '" + stockID + "' AND Color = '" + stockColor + "' And Size = " + stockSize + ""
                cmd.ExecuteNonQuery()
                vUpdate = False
                Exit While
            End If
            If readerObj.IsDBNull(1) Then
                readerObj.Close()
                Exit While
            End If
        End While

        If vUpdate Then
            readerObj.Close()
            cmd.CommandText = "INSERT INTO Parts VALUES('" + stockID + "', '" + stockColor + "' , " + stockSize + ", " + stockCount + ")"
            cmd.ExecuteNonQuery()
        End If '

        Label5.Text = "Added " + stockCount + " " + stockID
        Label5.Visible = True

        Dim adp As SqlDataAdapter = New SqlDataAdapter("SELECT stockID, Color, Size, Count FROM dbo.Parts", "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=Parts;Integrated Security=True")
        Dim ds As DataSet = New DataSet()
        adp.Fill(ds)
        DataGridView1.DataSource = ds.Tables(0)

    End Sub

End Class