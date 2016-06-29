Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Public Class StockInputExcell
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
            For it = 0 To shXL.UsedRange.Rows.Count - 1
                DataGridView1.Rows.Add()
                DataGridView1.Rows(it).Cells(0).Value = ""
                DataGridView1.Rows(it).Cells(1).Value = shXL.Cells(2 + it, 2).Value
                DataGridView1.Rows(it).Cells(2).Value = shXL.Cells(2 + it, 3).Value
                DataGridView1.Rows(it).Cells(3).Value = shXL.Cells(2 + it, 4).Value
                DataGridView1.Rows(it).Cells(4).Value = ""
                DataGridView1.Rows(it).Cells(5).Value = 0
                DataGridView1.Rows(it).Cells(6).Value = shXL.Cells(2 + it, 7).Value
                DataGridView1.Rows(it).Cells(7).Value = 0
                DataGridView1.Rows(it).Cells(8).Value = ""
                DataGridView1.Rows(it).Cells(9).Value = 0
                DataGridView1.Rows(it).Cells(10).Value = ""
            Next

            DataGridView1.AutoResizeColumns()

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        For it = 0 To DataGridView1.RowCount - 2
            cmd.CommandText = "INSERT INTO stockNew VALUES('" + DataGridView1.Rows(it).Cells(0).Value.ToString + "', '" + DataGridView1.Rows(it).Cells(1).Value.ToString + "' , '" + DataGridView1.Rows(it).Cells(2).Value.ToString + "', '" + DataGridView1.Rows(it).Cells(3).Value.ToString + "' , '" + DataGridView1.Rows(it).Cells(4).Value.ToString + "', " + DataGridView1.Rows(it).Cells(5).Value.ToString + " , " + DataGridView1.Rows(it).Cells(6).Value.ToString + ", " + DataGridView1.Rows(it).Cells(7).Value.ToString + ", '" + DataGridView1.Rows(it).Cells(8).Value.ToString + "'," + DataGridView1.Rows(it).Cells(9).Value.ToString + ",'" + DataGridView1.Rows(it).Cells(10).Value.ToString + "')"
            cmd.ExecuteNonQuery()
        Next

    End Sub
End Class