Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Public Class PartInputExcell
    Private Sub this_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Easy Cut V1.0"
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "Excell files (*.xls, *.xlsx)|*.xls;*.xlsx"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True
        fd.Multiselect = False
        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName

            Dim appXL As Excel.Application
            Dim wbXl As Excel.Workbook
            Dim shXL As Excel.Worksheet
            ' Start Excel and get Application object.
            appXL = CreateObject("Excel.Application")
            'appXL.Visible = True
            ' Add a new workbook.
            wbXl = appXL.Workbooks.Open(strFileName)
            shXL = wbXl.ActiveSheet

            Dim internalID As ArrayList = New ArrayList
            Dim usedID As ArrayList = New ArrayList

            cmd.CommandText = "SELECT internalID  FROM parts"
            cmd.ExecuteNonQuery()
            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            While readerObj.Read
                Dim temp1 As Integer = Convert.ToInt64(readerObj("internalID").ToString)
                internalID.Add(temp1)
            End While

            Dim color1 As String = ""

            Dim set1 As Double = 0
            Dim set2 As Integer = 0
            Dim colorall As Boolean = False

            Dim frm1 As partColorAll = New partColorAll()
            If frm1.ShowDialog() = DialogResult.OK Then
                color1 = frm1.ListBox1.SelectedItem
                colorall = True
            Else
                color1 = ""
            End If

            For it = 0 To shXL.UsedRange.Rows.Count - 1
                Dim temp As String = Convert.ToString(it)
                Dim temp1 As String = Convert.ToString(shXL.UsedRange.Rows.Count - 1)
                Label1.Text = "Loading " + temp + "/" + temp1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(it).Cells(0).Value = shXL.Cells(2 + it, 8).Value
                DataGridView1.Rows(it).Cells(1).Value = shXL.Cells(2 + it, 9).Value

                Dim inputinternal As Integer
                Dim rn As New Random
                inputinternal = rn.Next(100000, 999999)
                For it1 = 0 To internalID.Count - 1
                    If internalID(it1) = inputinternal Or usedID.Contains(inputinternal) Then
                        it1 = 0
                        inputinternal = rn.Next(100000, 999999)
                    End If
                Next
                usedID.Add(inputinternal)

                Dim frm2 As partColor = New partColor()
                If Not colorall And (shXL.Cells(2 + it, 3).Value <> set1 Or shXL.Cells(2 + it, 4).Value <> set2) Then
                    If frm2.ShowDialog() = DialogResult.OK Then
                        color1 = frm2.ListBox1.SelectedItem
                    Else
                        color1 = ""
                    End If
                End If
                set1 = shXL.Cells(2 + it, 3).Value
                set2 = shXL.Cells(2 + it, 4).Value


                DataGridView1.Rows(it).Cells(2).Value = color1
                DataGridView1.Rows(it).Cells(3).Value = shXL.Cells(2 + it, 11).Value
                DataGridView1.Rows(it).Cells(4).Value = shXL.Cells(2 + it, 10).Value
                DataGridView1.Rows(it).Cells(5).Value = inputinternal
                DataGridView1.Rows(it).Cells(6).Value = shXL.Cells(2 + it, 1).Value
                DataGridView1.Rows(it).Cells(7).Value = shXL.Cells(2 + it, 3).Value
                DataGridView1.Rows(it).Cells(8).Value = shXL.Cells(2 + it, 4).Value
                DataGridView1.Rows(it).Cells(9).Value = ""
                DataGridView1.Rows(it).Cells(10).Value = 0
                DataGridView1.Rows(it).Cells(11).Value = ""

                DataGridView1.FirstDisplayedScrollingRowIndex = it
                If it > 0 Then
                    DataGridView1.Rows(it - 1).Cells(0).Selected = False
                End If
            Next

            DataGridView1.AutoResizeColumns()
            readerObj.Close()
            wbXl.Close()
            DataGridView1.Rows(0).Cells(0).Selected = True
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        For it = 0 To DataGridView1.RowCount - 1
            If Not (DataGridView1.Rows(it).Cells(1).Value Is Nothing) Then
                cmd.CommandText = "INSERT INTO parts VALUES('" + DataGridView1.Rows(it).Cells(0).Value.ToString + "', '" + DataGridView1.Rows(it).Cells(1).Value.ToString + "' , '" + DataGridView1.Rows(it).Cells(2).Value.ToString + "', " + DataGridView1.Rows(it).Cells(3).Value.ToString + " , " + DataGridView1.Rows(it).Cells(4).Value.ToString + ", " + DataGridView1.Rows(it).Cells(5).Value.ToString + " , '" + DataGridView1.Rows(it).Cells(6).Value.ToString + "', " + DataGridView1.Rows(it).Cells(7).Value.ToString + ", " + DataGridView1.Rows(it).Cells(8).Value.ToString + ",'" + DataGridView1.Rows(it).Cells(9).Value.ToString + "'," + DataGridView1.Rows(it).Cells(10).Value.ToString + ",'" + DataGridView1.Rows(it).Cells(11).Value.ToString + " ')"
                cmd.ExecuteNonQuery()
            End If
        Next

    End Sub
End Class