Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

Public Class NavInput
    Private Sub this_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Easy Cut V1.0"
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=InventoryManagement;Integrated Security=True"
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


            Dim totalcount As String = Convert.ToString(shXL.UsedRange.Rows.Count - 1)
            For it = 0 To shXL.UsedRange.Rows.Count - 1
                Dim c0 As String = shXL.Cells(2 + it, 1).Value
                Dim c1 As String = shXL.Cells(2 + it, 2).Value
                Dim c2 As String = shXL.Cells(2 + it, 10).Value
                Dim c3 As String = shXL.Cells(2 + it, 7).Value
                Dim c4 As String = shXL.Cells(2 + it, 8).Value
                Dim c5 As String = shXL.Cells(2 + it, 16).Value
                Dim c6 As String = ""
                Dim c7 As String = 0
                Dim c8 As String = ""

                Dim temp As String = Convert.ToString(it)
                Label1.Text = "Loading " + temp + "/" + totalcount

                Dim in0 As String = ""
                Dim in1 As String = ""
                Dim in2 As String = ""
                Dim in3 As String = ""
                Dim in4 As String = ""
                Dim in5 As String = ""
                Dim in6 As String = ""
                Dim in7 As String = ""
                Dim in8 As String = ""

                If c0 IsNot Nothing Then
                    in0 = c0
                End If
                If c1 IsNot Nothing Then
                    in1 = c1
                End If
                If c2 IsNot Nothing Then
                    in2 = c2
                End If
                If c3 IsNot Nothing Then
                    in3 = c3
                End If
                If c4 IsNot Nothing Then
                    in4 = c4
                End If
                If c5 IsNot Nothing Then
                    in5 = c5
                End If
                If c6 IsNot Nothing Then
                    in6 = c6
                End If
                If c7 IsNot Nothing Then
                    in7 = c7
                End If
                If c8 IsNot Nothing Then
                    in8 = c8
                End If

                DataGridView1.Rows.Add()
                DataGridView1.Rows(it).Cells(0).Value = in0
                DataGridView1.Rows(it).Cells(1).Value = in1
                DataGridView1.Rows(it).Cells(2).Value = in2
                DataGridView1.Rows(it).Cells(3).Value = in3
                DataGridView1.Rows(it).Cells(4).Value = in4
                DataGridView1.Rows(it).Cells(5).Value = in5
                DataGridView1.Rows(it).Cells(6).Value = in6
                DataGridView1.Rows(it).Cells(7).Value = in7
                DataGridView1.Rows(it).Cells(8).Value = in8
            Next
            Label1.Text = "Loading Complete"
            DataGridView1.AutoResizeColumns()
            'readerObj.Close()
            shXL = Nothing
            wbXl = Nothing
            appXL.Quit()
            appXL = Nothing
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label1.Text = "Uploading..."
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=InventoryManagement;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        Dim con1 As New SqlConnection
        Dim cmd1 As New SqlCommand
        con1.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=InventoryManagement;Integrated Security=True"
        con1.Open()
        cmd1.Connection = con1
        Dim ar0 As ArrayList = New ArrayList
        Dim ar1 As ArrayList = New ArrayList
        Dim ar2 As ArrayList = New ArrayList
        Dim ar3 As ArrayList = New ArrayList
        Dim ar4 As ArrayList = New ArrayList
        Dim ar5 As ArrayList = New ArrayList
        Dim ar6 As ArrayList = New ArrayList
        Dim ar7 As ArrayList = New ArrayList
        Dim ar8 As ArrayList = New ArrayList
        For it = 0 To DataGridView1.RowCount - 1
            If Not String.Equals(DataGridView1.Rows(it).Cells(3).Value, "") And DataGridView1.Rows(it).Cells(3).Value IsNot Nothing Then
                ar0.Add(DataGridView1.Rows(it).Cells(0).Value.ToString.ToUpper())
                ar1.Add(DataGridView1.Rows(it).Cells(1).Value.ToString.ToUpper())
                ar2.Add(DataGridView1.Rows(it).Cells(2).Value.ToString.ToUpper())
                ar3.Add(DataGridView1.Rows(it).Cells(3).Value.ToString.ToUpper())
                ar4.Add(DataGridView1.Rows(it).Cells(4).Value.ToString.ToUpper())
                ar5.Add(DataGridView1.Rows(it).Cells(5).Value.ToString.ToUpper())
                ar6.Add(DataGridView1.Rows(it).Cells(6).Value.ToString.ToUpper())
                ar7.Add(DataGridView1.Rows(it).Cells(7).Value.ToString.ToUpper())
                ar8.Add(DataGridView1.Rows(it).Cells(8).Value.ToString.ToUpper())
            End If
        Next
        cmd.CommandText = "SELECT ID1, count FROM Nav"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
        While readerObj.Read
            For it = 0 To ar1.Count - 1
                If String.Equals(ar1(it), readerObj("ID1").ToString) Then

                    cmd1.CommandText = "UPDATE Nav SET count = " + ar5(it) + " WHERE ID1 = " + readerObj("ID1").ToString
                    cmd1.ExecuteNonQuery()
                    ar0.RemoveAt(it)
                    ar1.RemoveAt(it)
                    ar2.RemoveAt(it)
                    ar3.RemoveAt(it)
                    ar4.RemoveAt(it)
                    ar5.RemoveAt(it)
                    ar6.RemoveAt(it)
                    ar7.RemoveAt(it)
                    ar8.RemoveAt(it)
                    Exit For
                End If
            Next
        End While

        For it = 0 To ar1.Count - 1
            cmd1.CommandText = "INSERT INTO Nav VALUES('" + ar0(it) + "', '" + ar1(it) + "' , '" + ar2(it) + "', @description" + it.ToString + " , '" + ar4(it) + "', " + ar5(it) + " , '" + ar6(it) + "', " + ar7(it) + ", '" + ar8(it) + "')"
            cmd1.Parameters.Add("@description" + it.ToString + "", SqlDbType.NVarChar).Value = ar3(it)
            cmd1.ExecuteNonQuery()
        Next

        Label1.Text = "Upload Complete"
        MessageBox.Show("Upload Complete!")
        Me.Close()
    End Sub

End Class