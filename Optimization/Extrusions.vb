Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

Public Class Extrusions

    Dim stockID1 As ArrayList = New ArrayList
    Dim stockID2 As ArrayList = New ArrayList
    Dim stockID3 As ArrayList = New ArrayList
    Dim description As ArrayList = New ArrayList
    Dim color As ArrayList = New ArrayList
    Dim size1 As ArrayList = New ArrayList
    Dim internalID As ArrayList = New ArrayList
    Dim Used As ArrayList = New ArrayList

    Dim ostockID1 As ArrayList = New ArrayList
    Dim ostockID2 As ArrayList = New ArrayList
    Dim ostockID3 As ArrayList = New ArrayList
    Dim odescription As ArrayList = New ArrayList
    Dim ocolor As ArrayList = New ArrayList
    Dim osize1 As ArrayList = New ArrayList
    Dim ointernalID As ArrayList = New ArrayList
    Dim oUsed As ArrayList = New ArrayList

    '
    ' Page 1
    '

    Private Sub Extrusions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'OptimizationDatabaseDataSet.stockNew' table. You can move, or remove it, as needed.
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"

        ListBox1.Items.Add("Bronze")
        ListBox1.Items.Add("White")
        ListBox1.Items.Add("Silver")
        ListBox1.Items.Add("Other")
        ListBox2.Items.Add("New Part")

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, internalID FROM stockNew"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        '
        ' Reads data from database 
        '
        While readerObj.Read
            Dim tempSize As Integer = Convert.ToInt32(readerObj("Size").ToString)

            If Not ComboBox1.Items.Contains(readerObj("stockID1").ToString) Then
                ComboBox1.Items.Add(readerObj("stockID1").ToString)
            End If
            If Not ComboBox2.Items.Contains(readerObj("stockID2").ToString) Then
                ComboBox2.Items.Add(readerObj("stockID2").ToString)
            End If
            If Not ComboBox3.Items.Contains(readerObj("stockID3").ToString) Then
                ComboBox3.Items.Add(readerObj("stockID3").ToString)
            End If
            description.Add(readerObj("description").ToString)
            stockID1.Add(readerObj("stockID1").ToString)
            stockID2.Add(readerObj("stockID2").ToString)
            stockID3.Add(readerObj("stockID3").ToString)
            color.Add(readerObj("color").ToString)
            size1.Add(readerObj("size").ToString)
            internalID.Add(readerObj("internalID").ToString)

            If readerObj.IsDBNull(1) Then
                readerObj.Close()
                Exit While
            End If
        End While

    End Sub

    Private Sub comboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        For it1 = 0 To description.Count - 1
            If String.Equals(ComboBox1.SelectedItem, stockID1(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            End If
            If Not String.Equals(ComboBox1.SelectedItem, stockID1(it1)) Then
                ListBox2.Items.Remove(description(it1) + " " + color(it1) + " ")
                Used.Remove(it1)
            End If
        Next

    End Sub

    Private Sub comboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged

        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox2.SelectedItem, stockID2(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            End If
            If Not String.Equals(ComboBox2.SelectedItem, stockID2(it1)) Then
                ListBox2.Items.Remove(description(it1) + " " + color(it1) + " ")
                Used.Remove(it1)
            End If
        Next

    End Sub

    Private Sub comboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged

        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox3.SelectedItem, stockID3(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            End If
            If Not String.Equals(ComboBox3.SelectedItem, stockID3(it1)) Then
                ListBox2.Items.Remove(description(it1) + " " + color(it1) + " ")
                Used.Remove(it1)
            End If
        Next

    End Sub

    Private Sub ListBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedValueChanged
        Dim pString As String = ""
        If ListBox2.SelectedIndex = 0 Then
            pString = "Cannot find piece, Click Input to input new stock"
        End If
        If ListBox2.SelectedIndex > 0 Then
            pString = description(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + color(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + size1(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + stockID1(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + stockID2(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + stockID3(Used(ListBox2.SelectedIndex - 1))

        End If
        RichTextBox1.Text = pString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT count, internalID FROM stockNew"
        cmd.ExecuteNonQuery()

        If ListBox2.SelectedItem = "New Part" Then
            Dim frmNewExtr As NewExtrusions = New NewExtrusions()
            frmNewExtr.ShowDialog()

        ElseIf ListBox2.SelectedIndex >= 0 Then

            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            While readerObj.Read
                If String.Equals(internalID(Used(ListBox2.SelectedIndex - 1)), readerObj("internalID").ToString) Then
                    Dim temp1 As Integer = Convert.ToInt32(TextBox1.Text)
                    Dim temp2 As String = readerObj("count").ToString
                    Dim temp3 As Integer = Convert.ToInt32(temp2)
                    temp3 = temp3 + temp1
                    temp2 = Convert.ToString(temp3)
                    readerObj.Close()

                    cmd.CommandText = "UPDATE stockNew SET count = " + temp2 + " WHERE internalID = " + internalID(Used(ListBox2.SelectedIndex - 1))
                    cmd.ExecuteNonQuery()
                    Exit While
                End If
            End While
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT count, internalID FROM stockNew"
        cmd.ExecuteNonQuery()

        If ListBox2.SelectedItem = "New Part" Then
            Dim frmNewExtr As NewExtrusions = New NewExtrusions()
            frmNewExtr.ShowDialog()

        ElseIf ListBox2.SelectedIndex >= 0 Then

            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            While readerObj.Read
                If String.Equals(internalID(Used(ListBox2.SelectedIndex - 1)), readerObj("internalID").ToString) Then
                    Dim temp1 As Integer = Convert.ToInt32(TextBox1.Text)
                    Dim temp2 As String = readerObj("count").ToString
                    Dim temp3 As Integer = Convert.ToInt32(temp2)
                    temp3 = temp3 - temp1
                    temp2 = Convert.ToString(temp3)
                    readerObj.Close()

                    cmd.CommandText = "UPDATE stockNew SET count = " + temp2 + " WHERE internalID = " + internalID(Used(ListBox2.SelectedIndex - 1))
                    cmd.ExecuteNonQuery()
                    Exit While
                End If
            End While
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, count, internalID FROM stockNew"
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

            ' Format A1:D1 as bold, vertical alignment = center.
            With shXL.Range("A1", "G1")
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
            End With
            it1 += 1

        End While

        ' AutoFit columns A:D.
        raXL = shXL.Range("A1", "G1")
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

    '
    ' Page 2
    '


End Class