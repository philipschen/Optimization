Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Public Class PartInputExcell
    Dim connectionstring As Class1 = New Class1
    Private Sub this_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
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
            Dim shopnumbers As ArrayList = New ArrayList
            cmd.CommandText = "SELECT internalID, shopNumber  FROM parts"
            cmd.ExecuteNonQuery()
            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            While readerObj.Read
                Dim temp1 As Integer = Convert.ToInt64(readerObj("internalID").ToString)
                internalID.Add(temp1)
                If Not shopnumbers.Contains(readerObj("shopNumber").ToString) Then
                    shopnumbers.Add(readerObj("shopNumber").ToString)
                End If
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


            Dim alreadyentered As Boolean = False
            For it = 0 To shXL.UsedRange.Rows.Count - 1
                Dim temp1 As String = Convert.ToString(shXL.UsedRange.Rows.Count - 1)
                Label1.Text = "Loading " + it.ToString + "/" + temp1

                DataGridView1.Rows.Add()
                DataGridView1.Rows(it).Cells(0).Value = shXL.Cells(2 + it, 8).Value
                DataGridView1.Rows(it).Cells(1).Value = shXL.Cells(2 + it, 9).Value

                Dim inputinternal As Integer
                Dim rn As New Random
                inputinternal = rn.Next(1000000, 9999999)
                For it1 = 0 To internalID.Count - 1
                    If internalID(it1) = inputinternal Or usedID.Contains(inputinternal) Then
                        it1 = 0
                        inputinternal = rn.Next(1000000, 9999999)
                    End If
                Next
                usedID.Add(inputinternal)

                Dim cell3 As String = shXL.Cells(2 + it, 3).Value
                Dim cell4 As String = shXL.Cells(2 + it, 4).Value

                If Not colorall And (cell3 <> set1 Or cell4 <> set2) Then
                    DataGridView1.FirstDisplayedScrollingRowIndex = it
                    If it > 0 Then
                        DataGridView1.Rows(it - 1).Cells(0).Selected = False
                    End If
                    Dim frm2 As partColor = New partColor()
                    If frm2.ShowDialog() = DialogResult.OK Then

                        color1 = frm2.ListBox1.SelectedItem
                    Else
                        color1 = ""
                    End If
                    DataGridView1.FirstDisplayedScrollingRowIndex = 0
                End If

                set1 = cell3
                set2 = cell4


                DataGridView1.Rows(it).Cells(2).Value = color1
                DataGridView1.Rows(it).Cells(3).Value = shXL.Cells(2 + it, 11).Value
                DataGridView1.Rows(it).Cells(4).Value = shXL.Cells(2 + it, 10).Value
                DataGridView1.Rows(it).Cells(5).Value = inputinternal
                DataGridView1.Rows(it).Cells(6).Value = shXL.Cells(2 + it, 1).Value
                If shopnumbers.Contains(DataGridView1.Rows(it).Cells(6).Value) Then
                    Dim result1 As DialogResult = MessageBox.Show("This Shop Number BOM has already been entered do you want to Continue?", "ShopNumber Already Entered", MessageBoxButtons.YesNo)
                    If result1 = DialogResult.Yes Then
                        shopnumbers.Remove(DataGridView1.Rows(it).Cells(6).Value)
                    ElseIf result1 = DialogResult.No Then
                        alreadyentered = True
                        Exit For
                    End If
                End If
                DataGridView1.Rows(it).Cells(7).Value = cell3
                DataGridView1.Rows(it).Cells(8).Value = cell4
                DataGridView1.Rows(it).Cells(9).Value = ""
                DataGridView1.Rows(it).Cells(10).Value = it
                DataGridView1.Rows(it).Cells(11).Value = ""
            Next
            If alreadyentered Then
                Me.Close()
            Else

                Label1.Text = "Loading Complete"
                DataGridView1.AutoResizeColumns()
                readerObj.Close()

                shXL = Nothing
                wbXl = Nothing
                appXL.Quit()
                appXL = Nothing
                GC.Collect()
                DataGridView1.Rows(0).Cells(0).Selected = True
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label1.Text = "Uploading..."
        'DataGridView1.ClearSelection()
        'DataGridView1.Rows(0).Cells(0).Selected = True
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        For it = 0 To DataGridView1.RowCount - 1
            If Not (DataGridView1.Rows(it).Cells(1).Value Is Nothing) Then
                cmd.CommandText = "INSERT INTO parts VALUES('" + DataGridView1.Rows(it).Cells(0).Value.ToString.ToUpper() + "', '" + DataGridView1.Rows(it).Cells(1).Value.ToString.ToUpper() + "' , '" + DataGridView1.Rows(it).Cells(2).Value.ToString.ToUpper() + "', " + DataGridView1.Rows(it).Cells(3).Value.ToString.ToUpper() + " , " + DataGridView1.Rows(it).Cells(4).Value.ToString.ToUpper() + ", " + DataGridView1.Rows(it).Cells(5).Value.ToString.ToUpper() + " , '" + DataGridView1.Rows(it).Cells(6).Value.ToString.ToUpper() + "', " + DataGridView1.Rows(it).Cells(7).Value.ToString.ToUpper() + ", " + DataGridView1.Rows(it).Cells(8).Value.ToString.ToUpper() + ",'" + DataGridView1.Rows(it).Cells(9).Value.ToString.ToUpper() + "'," + DataGridView1.Rows(it).Cells(10).Value.ToString.ToUpper() + ",'" + DataGridView1.Rows(it).Cells(11).Value.ToString.ToUpper() + " ')"
                cmd.ExecuteNonQuery()
            End If
        Next
        Label1.Text = "Upload Complete"
        Dim frmNewExtr As uploadComplete = New uploadComplete()
        frmNewExtr.ShowDialog()
        Me.Close()
    End Sub
End Class