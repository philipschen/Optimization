Imports System.Data.SqlClient
Public Class DirectlyEdit
    Private Sub DirectlyEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'OptimizationDatabaseDataSet.stockNew' table. You can move, or remove it, as needed.
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
        Me.DataGridView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridView1.AutoResizeColumns()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "DELETE FROM stockNew"
        cmd.ExecuteNonQuery()

        For it = 0 To DataGridView1.RowCount - 1
            cmd.CommandText = "INSERT INTO stockNew VALUES('" + DataGridView1.Rows(it).Cells(0).Value.ToString + "', '" + DataGridView1.Rows(it).Cells(1).Value.ToString + "' , '" + DataGridView1.Rows(it).Cells(2).Value.ToString + "', '" + DataGridView1.Rows(it).Cells(3).Value.ToString + "' , '" + DataGridView1.Rows(it).Cells(4).Value.ToString + "', " + DataGridView1.Rows(it).Cells(5).Value.ToString + " , " + DataGridView1.Rows(it).Cells(6).Value.ToString + ", " + DataGridView1.Rows(it).Cells(7).Value.ToString + ", '" + DataGridView1.Rows(it).Cells(8).Value.ToString + "'," + DataGridView1.Rows(it).Cells(9).Value.ToString + ",'" + DataGridView1.Rows(it).Cells(10).Value.ToString + "')"
            cmd.ExecuteNonQuery()
        Next
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
    End Sub


End Class