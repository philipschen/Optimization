Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Public Class DirectlyEdit
    Dim connectionstring As Class1 = New Class1
    Private Sub DirectlyEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'OptimizationDatabaseDataSet3.stockUsed' table. You can move, or remove it, as needed.
        Me.StockUsedTableAdapter.Fill(Me.OptimizationDatabaseDataSet3.stockUsed)
        Me.Text = "Easy Cut V1.0"
        'TODO: This line of code loads data into the 'OptimizationDatabaseDataSet1.parts' table. You can move, or remove it, as needed.
        Me.PartsTableAdapter.Fill(Me.OptimizationDatabaseDataSet1.parts)
        'TODO: This line of code loads data into the 'OptimizationDatabaseDataSet.stockNew' table. You can move, or remove it, as needed.
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)

        Me.DataGridView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridView1.AutoResizeColumns()
        Me.DataGridView2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridView2.AutoResizeColumns()
        Me.DataGridView3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridView3.AutoResizeColumns()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = connectionstring.version
        cmd.ExecuteNonQuery()

        For it = 0 To DataGridView1.RowCount - 1
            cmd.CommandText = "INSERT INTO stockNew VALUES('" + DataGridView1.Rows(it).Cells(0).Value.ToString.ToUpper() + "', '" + DataGridView1.Rows(it).Cells(1).Value.ToString.ToUpper() + "' , '" + DataGridView1.Rows(it).Cells(2).Value.ToString.ToUpper() + "', '" + DataGridView1.Rows(it).Cells(3).Value.ToString.ToUpper() + "' , '" + DataGridView1.Rows(it).Cells(4).Value.ToString.ToUpper() + "', " + DataGridView1.Rows(it).Cells(5).Value.ToString.ToUpper() + " , " + DataGridView1.Rows(it).Cells(6).Value.ToString.ToUpper() + ", " + DataGridView1.Rows(it).Cells(7).Value.ToString.ToUpper() + ", '" + DataGridView1.Rows(it).Cells(8).Value.ToString.ToUpper() + "'," + DataGridView1.Rows(it).Cells(9).Value.ToString.ToUpper() + ",'" + DataGridView1.Rows(it).Cells(10).Value.ToString.ToUpper() + "')"
            cmd.ExecuteNonQuery()
        Next
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "DELETE FROM parts"
        cmd.ExecuteNonQuery()

        For it = 0 To DataGridView2.RowCount - 1
            cmd.CommandText = "INSERT INTO parts VALUES('" + DataGridView1.Rows(it).Cells(0).Value.ToString.ToUpper() + "', '" + DataGridView1.Rows(it).Cells(1).Value.ToString.ToUpper() + "' , '" + DataGridView1.Rows(it).Cells(2).Value.ToString.ToUpper() + "', " + DataGridView1.Rows(it).Cells(3).Value.ToString.ToUpper() + " , " + DataGridView1.Rows(it).Cells(4).Value.ToString.ToUpper() + ", " + DataGridView1.Rows(it).Cells(5).Value.ToString.ToUpper() + " , '" + DataGridView1.Rows(it).Cells(6).Value.ToString.ToUpper() + "', " + DataGridView1.Rows(it).Cells(7).Value.ToString.ToUpper() + ", " + DataGridView1.Rows(it).Cells(8).Value.ToString.ToUpper() + ",'" + DataGridView1.Rows(it).Cells(9).Value.ToString.ToUpper() + "'," + DataGridView1.Rows(it).Cells(10).Value.ToString.ToUpper() + ",'" + DataGridView1.Rows(it).Cells(11).Value.ToString.ToUpper() + " ')"
            cmd.ExecuteNonQuery()
        Next
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "DELETE FROM stockUsed"
        cmd.ExecuteNonQuery()

        For it = 0 To DataGridView2.RowCount - 1
            cmd.CommandText = "INSERT INTO stockUsed VALUES('" + DataGridView1.Rows(it).Cells(0).Value.ToString.ToUpper() + "', '" + DataGridView1.Rows(it).Cells(1).Value.ToString.ToUpper() + "' , '" + DataGridView1.Rows(it).Cells(2).Value.ToString.ToUpper() + "', '" + DataGridView1.Rows(it).Cells(3).Value.ToString.ToUpper() + "' , '" + DataGridView1.Rows(it).Cells(4).Value.ToString.ToUpper() + "', " + DataGridView1.Rows(it).Cells(5).Value.ToString.ToUpper() + " , " + DataGridView1.Rows(it).Cells(6).Value.ToString.ToUpper() + ", " + DataGridView1.Rows(it).Cells(7).Value.ToString.ToUpper() + ", '" + DataGridView1.Rows(it).Cells(8).Value.ToString.ToUpper() + "'," + DataGridView1.Rows(it).Cells(9).Value.ToString.ToUpper() + ",'" + DataGridView1.Rows(it).Cells(10).Value.ToString.ToUpper() + "')"
            cmd.ExecuteNonQuery()
        Next
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con
        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, count, internalID, context1, context3 FROM stockNew"
        cmd.ExecuteNonQuery()

        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        Dim appXL As Excel.Application
        Dim wbXl As Excel.Workbook
        Dim shXL As Excel.Worksheet
        Dim raXL As Excel.Range
        ' Start Excel and get Application object.
        appXL = CreateObject("Excel.Application")
        appXL.Visible = True
        ' Add a new workbook.
        wbXl = appXL.Workbooks.Add
        shXL = wbXl.ActiveSheet

        ' Add table headers going cell by cell.
        shXL.Cells(1, 1).Value = "ID1"
        shXL.Cells(1, 2).Value = "ID2"
        shXL.Cells(1, 3).Value = "ID3"
        shXL.Cells(1, 4).Value = "Description"
        shXL.Cells(1, 5).Value = "Color"
        shXL.Cells(1, 6).Value = "Size"
        shXL.Cells(1, 7).Value = "Count"
        shXL.Cells(1, 8).Value = "Saw Number"
        shXL.Cells(1, 9).Value = "Location of Used Parts"
        ' Format A1:D1 as bold, vertical alignment = center.
        With shXL.Range("A1", "I1")
            .Font.Bold = True
            .ColumnWidth = 20
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End With

        Dim it1 As Integer = 2
        While readerObj.Read
            ' Fill values.
            With shXL
                .Cells(it1, 1).Value = readerObj("stockID1").ToString
                .Cells(it1, 2).Value = readerObj("stockID2").ToString
                .Cells(it1, 3).Value = readerObj("stockID3").ToString
                .Cells(it1, 4).Value = readerObj("description").ToString
                .Cells(it1, 5).Value = readerObj("color").ToString
                .Cells(it1, 6).Value = readerObj("size").ToString
                .Cells(it1, 7).Value = readerObj("count").ToString
                .Cells(it1, 8).Value = readerObj("context1").ToString
                .Cells(it1, 9).Value = readerObj("context3").ToString

            End With
            it1 += 1

        End While

        ' AutoFit columns
        raXL = shXL.Range("A1", "I1")
        raXL.EntireColumn.AutoFit()
        ' Make sure Excel is visible and give the user control
        appXL.Visible = True
        appXL.UserControl = True
        ' Release object references.
        raXL = Nothing
        shXL = Nothing
        wbXl = Nothing
        appXL.Quit()
        appXL = Nothing
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con
        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, count, internalID, location, context1, FROM stockUsed"
        cmd.ExecuteNonQuery()

        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        Dim appXL As Excel.Application
        Dim wbXl As Excel.Workbook
        Dim shXL As Excel.Worksheet
        Dim raXL As Excel.Range
        ' Start Excel and get Application object.
        appXL = CreateObject("Excel.Application")
        appXL.Visible = True
        ' Add a new workbook.
        wbXl = appXL.Workbooks.Add
        shXL = wbXl.ActiveSheet

        Dim it1 As Integer = 2
        While readerObj.Read
            ' Add table headers going cell by cell.
            shXL.Cells(1, 1).Value = "ID1"
            shXL.Cells(1, 2).Value = "ID2"
            shXL.Cells(1, 3).Value = "ID3"
            shXL.Cells(1, 4).Value = "Description"
            shXL.Cells(1, 5).Value = "Color"
            shXL.Cells(1, 6).Value = "Size"
            shXL.Cells(1, 7).Value = "Count"
            shXL.Cells(1, 8).Value = "Location"
            shXL.Cells(1, 9).Value = "Saw Number"

            ' Format A1:G1 as bold, vertical alignment = center.
            With shXL.Range("A1", "I1")
                .Font.Bold = True
                .ColumnWidth = 20
                .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End With

            ' Fill values.
            With shXL
                .Cells(it1, 1).Value = readerObj("stockID1").ToString
                .Cells(it1, 2).Value = readerObj("stockID2").ToString
                .Cells(it1, 3).Value = readerObj("stockID3").ToString
                .Cells(it1, 4).Value = readerObj("description").ToString
                .Cells(it1, 5).Value = readerObj("color").ToString
                .Cells(it1, 6).Value = readerObj("size").ToString
                .Cells(it1, 7).Value = readerObj("count").ToString
                .Cells(it1, 8).Value = readerObj("location").ToString
                .Cells(it1, 9).Value = readerObj("context1").ToString
            End With
            it1 += 1

        End While
        readerObj.Close()
        ' AutoFit columns A:D.
        raXL = shXL.Range("A1", "I1")
        raXL.EntireColumn.AutoFit()
        ' Make sure Excel is visible and give the user control
        ' of Excel's lifetime.
        appXL.Visible = True
        appXL.UserControl = True
        ' Release object references.
        raXL = Nothing
        shXL = Nothing
        wbXl = Nothing
        appXL.Quit()
        appXL = Nothing
    End Sub
End Class