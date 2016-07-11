Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

Public Class Extrusions
    Dim connectionstring As Class1 = New Class1
    Dim stockID1 As ArrayList = New ArrayList
    Dim stockID2 As ArrayList = New ArrayList
    Dim stockID3 As ArrayList = New ArrayList
    Dim description As ArrayList = New ArrayList
    Dim color As ArrayList = New ArrayList
    Dim size1 As ArrayList = New ArrayList
    Dim internalID As ArrayList = New ArrayList
    Dim sawnumber As ArrayList = New ArrayList
    Dim Used As ArrayList = New ArrayList
    Dim oUsed As ArrayList = New ArrayList

    Private Sub Extrusions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'OptimizationDatabaseDataSet3.stockUsed' table. You can move, or remove it, as needed.
        Me.StockUsedTableAdapter.Fill(Me.OptimizationDatabaseDataSet3.stockUsed)
        'TODO: This line of code loads data into the 'OptimizationDatabaseDataSet.stockNew' table. You can move, or remove it, as needed.
        Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
        Me.DataGridView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridView1.AutoResizeColumns()
        Me.DataGridView2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridView2.AutoResizeColumns()
        ' Set the caption bar text of the form.  
        Me.Text = connectionstring.version

        Label19.Visible = False
        ListBox1.Visible = False
        ListBox4.Visible = False
        Label1.Visible = False
        Label12.Visible = False
        'ListBox1.Items.Add("Bronze")
        'ListBox1.Items.Add("White")
        'ListBox1.Items.Add("Silver")
        'ListBox1.Items.Add("Other")
        ListBox2.Items.Add("New Part")

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, internalID, context1 FROM stockNew"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        ' PAGE 1
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
            If Not ComboBox7.Items.Contains(readerObj("description").ToString) Then
                ComboBox7.Items.Add(readerObj("description").ToString)
            End If
            description.Add(readerObj("description").ToString)
            stockID1.Add(readerObj("stockID1").ToString)
            stockID2.Add(readerObj("stockID2").ToString)
            stockID3.Add(readerObj("stockID3").ToString)
            color.Add(readerObj("color").ToString)
            size1.Add(readerObj("size").ToString)
            internalID.Add(readerObj("internalID").ToString)
            sawnumber.Add(readerObj("context1").ToString)

            ' PAGE 2
            If Not ComboBox6.Items.Contains(readerObj("stockID1").ToString) Then
                ComboBox6.Items.Add(readerObj("stockID1").ToString)
            End If
            If Not ComboBox5.Items.Contains(readerObj("stockID2").ToString) Then
                ComboBox5.Items.Add(readerObj("stockID2").ToString)
            End If
            If Not ComboBox4.Items.Contains(readerObj("stockID3").ToString) Then
                ComboBox4.Items.Add(readerObj("stockID3").ToString)
            End If
            If Not ComboBox8.Items.Contains(readerObj("description").ToString) Then
                ComboBox8.Items.Add(readerObj("description").ToString)
            End If
        End While
        readerObj.Close()

    End Sub

    '
    ' Page 1
    '
    Private Sub comboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged
        ListBox2.Items.Clear()
        ListBox2.Items.Add("New Part")
        For it1 = 0 To description.Count - 1
            If String.Equals(ComboBox1.SelectedItem, stockID1(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            End If
            If Not String.Equals(ComboBox1.SelectedItem, stockID1(it1)) Then
                Used.Remove(it1)
            End If
        Next

    End Sub
    Private Sub comboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged
        ListBox2.Items.Clear()
        ListBox2.Items.Add("New Part")
        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox2.SelectedItem, stockID2(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            End If
            If Not String.Equals(ComboBox2.SelectedItem, stockID2(it1)) Then
                Used.Remove(it1)
            End If
        Next

    End Sub
    Private Sub comboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged
        ListBox2.Items.Clear()
        ListBox2.Items.Add("New Part")
        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox3.SelectedItem, stockID3(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)

            ElseIf Not String.Equals(ComboBox3.SelectedItem, stockID3(it1)) Then
                Used.Remove(it1)
            End If
        Next
    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged

        ListBox2.Items.Clear()
        ListBox2.Items.Add("New Part")
        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox7.SelectedItem, description(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            ElseIf Not String.Equals(ComboBox7.SelectedItem, stockID3(it1)) Then
                Used.Remove(it1)
            End If
        Next
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
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
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT count, internalID FROM stockNew"
        cmd.ExecuteNonQuery()

        If ListBox2.SelectedItem = "New Part" Then
            Dim frmNewExtr As NewExtrusions = New NewExtrusions()
            frmNewExtr.ShowDialog()

        ElseIf ListBox2.SelectedIndex >= 0 Then
            If TextBox1.Text Is Nothing OrElse TextBox1.Text = "" Then
                MessageBox.Show("Enter a count")
            Else
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
                readerObj.Close()
                Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT count, internalID FROM stockNew"
        cmd.ExecuteNonQuery()


        If ListBox2.SelectedIndex >= 0 Then
            If TextBox1.Text Is Nothing OrElse TextBox1.Text = "" Then
                MessageBox.Show("Enter a count")
            Else
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
                readerObj.Close()
                Me.StockNewTableAdapter.Fill(Me.OptimizationDatabaseDataSet.stockNew)
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label19.Visible = True
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
        ' Make sure Excel is visible and give the user control of Excel's lifetime.
        appXL.Visible = True
        appXL.UserControl = True
        ' Release object references.
        raXL = Nothing
        shXL = Nothing
        wbXl = Nothing
        appXL.Quit()
        appXL = Nothing
        GC.Collect()
        Label19.Visible = False
    End Sub

    '
    ' Page 2
    '
    Private Sub comboBox6_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedValueChanged
        ListBox3.Items.Clear()
        oUsed.Clear()
        For it1 = 0 To description.Count - 1
            If String.Equals(ComboBox6.SelectedItem, stockID1(it1)) And Not oUsed.Contains(it1) Then
                ListBox3.Items.Add(description(it1) + " " + color(it1) + " ")
                oUsed.Add(it1)
            End If
            'If Not String.Equals(ComboBox6.SelectedItem, stockID1(it1)) Then
            '    oUsed.Remove(it1)
            'End If
        Next

    End Sub

    Private Sub comboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        ListBox3.Items.Clear()
        oUsed.Clear()
        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox3.SelectedItem, stockID2(it1)) And Not oUsed.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                oUsed.Add(it1)
            End If
            'If Not String.Equals(ComboBox3.SelectedItem, stockID2(it1)) Then
            '    oUsed.Remove(it1)
            'End If
        Next

    End Sub

    Private Sub comboBox4_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedValueChanged
        ListBox3.Items.Clear()
        oUsed.Clear()
        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox4.SelectedItem, stockID3(it1)) And Not oUsed.Contains(it1) Then
                ListBox3.Items.Add(description(it1) + " " + color(it1) + " ")
                oUsed.Add(it1)
            End If
            'ElseIf Not String.Equals(ComboBox4.SelectedItem, stockID3(it1)) Then
            '    oUsed.Remove(it1)
            'End If
        Next
    End Sub
    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        ListBox3.Items.Clear()
        oUsed.Clear()
        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox8.SelectedItem, description(it1)) And Not oUsed.Contains(it1) Then
                ListBox3.Items.Add(description(it1) + " " + color(it1) + " ")
                oUsed.Add(it1)
            End If
            'ElseIf Not String.Equals(ComboBox8.SelectedItem, stockID3(it1)) Then
            '    oUsed.Remove(it1)
            'End If
        Next
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) AndAlso Not e.KeyChar = "." Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub
    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            e.Handled = True
        End If
    End Sub
    Private Sub ListBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedValueChanged
        Dim pString As String = ""
        pString = description(oUsed(ListBox3.SelectedIndex)) + vbCrLf + color(oUsed(ListBox3.SelectedIndex)) + vbCrLf + size1(oUsed(ListBox3.SelectedIndex)) + vbCrLf + stockID1(oUsed(ListBox3.SelectedIndex)) + vbCrLf + stockID2(oUsed(ListBox3.SelectedIndex)) + vbCrLf + stockID3(oUsed(ListBox3.SelectedIndex))
        RichTextBox2.Text = pString
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If TextBox2 Is Nothing Or TextBox3.Text Is Nothing OrElse TextBox2.Text = "" OrElse TextBox3.Text = "" Then
            MessageBox.Show("Enter a size and count")
        Else
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = connectionstring.connect1
            con.Open()
            cmd.Connection = con
            Dim con1 As New SqlConnection
            Dim cmd1 As New SqlCommand
            con1.ConnectionString = connectionstring.connect1
            con1.Open()
            cmd1.Connection = con1

            Dim context2 As ArrayList = New ArrayList
            Dim usedID As ArrayList = New ArrayList

            cmd.CommandText = "SELECT context2  FROM stockUsed"
            cmd.ExecuteNonQuery()
            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            While readerObj.Read
                Dim temp2 As Integer = Convert.ToInt64(readerObj("context2").ToString)
                context2.Add(temp2)
            End While
            readerObj.Close()

            Dim stockID2 As ArrayList = New ArrayList
            Dim internalID As ArrayList = New ArrayList
            cmd.CommandText = "SELECT stockID2, internalID  FROM stockNew"
            cmd.ExecuteNonQuery()
            readerObj = cmd.ExecuteReader
            While readerObj.Read
                stockID2.Add(readerObj("stockID2").ToString)
                internalID.Add(readerObj("internalID").ToString)
            End While
            readerObj.Close()

            Dim inputinternal As Integer
            Dim rn As New Random
            inputinternal = rn.Next(1000000, 9999999)
            For it1 = 0 To context2.Count - 1
                If context2(it1) = inputinternal Or usedID.Contains(inputinternal) Then
                    it1 = 0
                    inputinternal = rn.Next(1000000, 9999999)
                End If
            Next
            usedID.Add(inputinternal)

            Dim in0 As String = stockID1(oUsed(ListBox3.SelectedIndex))
            Dim in1 As String = stockID2(oUsed(ListBox3.SelectedIndex))
            Dim in2 As String = stockID3(oUsed(ListBox3.SelectedIndex))
            Dim in3 As String = description(oUsed(ListBox3.SelectedIndex))
            Dim in4 As String = color(oUsed(ListBox3.SelectedIndex))
            Dim in5 As String = TextBox2.Text
            Dim in6 As String = TextBox3.Text
            Dim in7 As String = ""
            Dim in8 As String = ""
            Dim in9 As String = sawnumber(oUsed(ListBox3.SelectedIndex))
            Dim in10 As String = inputinternal
            Dim in11 As String = ""
            Dim validrow As Boolean = False
            Dim notempty As Boolean = False
            Dim found As Boolean = False


            found = False
            For it1 = 0 To stockID2.Count - 1
                If String.Equals(stockID2(it1), in1) Then
                    in7 = internalID(it1)
                    found = True
                End If
            Next
            If Not found Then
                Dim exitfrm As CannotFind = New CannotFind()
                exitfrm.ShowDialog()
            End If
            cmd.CommandText = "SELECT stockID2, size, count, context2  FROM stockUsed"
            cmd.ExecuteNonQuery()
            readerObj = cmd.ExecuteReader
            Dim inlist As Boolean = False
            While readerObj.Read
                If String.Equals(in1, readerObj("stockID2").ToString) And String.Equals(in5, readerObj("size").ToString) Then
                    Dim temp1 As Integer = Convert.ToInt32(in6)
                    Dim temp2 As String = readerObj("count").ToString
                    Dim temp3 As Integer = Convert.ToInt32(temp2)
                    temp3 = temp3 + temp1
                    cmd1.CommandText = "UPDATE stockUsed SET count = " + temp3.ToString + " WHERE context2 = " + readerObj("context2").ToString
                    cmd1.ExecuteNonQuery()
                    inlist = True
                End If
            End While
            readerObj.Close()

            If Not inlist Then
                If in6 > 0 Then
                    cmd1.CommandText = "INSERT INTO stockUsed VALUES('" + in0 + "', '" + in1 + "', '" + in2 + "', '" + in3 + "', '" + "" + "', " + in5 + ", " + in6 + ", " + in7 + ", '" + "" + "', '" + "" + "', " + in10 + ", '" + "" + "')"
                    cmd1.ExecuteNonQuery()
                End If
            End If
                Me.StockUsedTableAdapter.Fill(Me.OptimizationDatabaseDataSet3.stockUsed)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox2 Is Nothing Or TextBox3.Text Is Nothing OrElse TextBox2.Text = "" OrElse TextBox3.Text = "" Then
            MessageBox.Show("Enter a size and count")
        Else
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = connectionstring.connect1
            con.Open()
            cmd.Connection = con
            Dim con1 As New SqlConnection
            Dim cmd1 As New SqlCommand
            con1.ConnectionString = connectionstring.connect1
            con1.Open()
            cmd1.Connection = con1

            Dim context2 As ArrayList = New ArrayList
            Dim usedID As ArrayList = New ArrayList

            cmd.CommandText = "SELECT context2  FROM stockUsed"
            cmd.ExecuteNonQuery()
            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            While readerObj.Read
                Dim temp2 As Integer = Convert.ToInt64(readerObj("context2").ToString)
                context2.Add(temp2)
            End While
            readerObj.Close()

            Dim stockID2 As ArrayList = New ArrayList
            Dim internalID As ArrayList = New ArrayList
            cmd.CommandText = "SELECT stockID2, internalID  FROM stockNew"
            cmd.ExecuteNonQuery()
            readerObj = cmd.ExecuteReader
            While readerObj.Read
                stockID2.Add(readerObj("stockID2").ToString)
                internalID.Add(readerObj("internalID").ToString)
            End While
            readerObj.Close()

            Dim inputinternal As Integer
            Dim rn As New Random
            inputinternal = rn.Next(1000000, 9999999)
            For it1 = 0 To context2.Count - 1
                If context2(it1) = inputinternal Or usedID.Contains(inputinternal) Then
                    it1 = 0
                    inputinternal = rn.Next(1000000, 9999999)
                End If
            Next
            usedID.Add(inputinternal)

            Dim in0 As String = stockID1(oUsed(ListBox3.SelectedIndex))
            Dim in1 As String = stockID2(oUsed(ListBox3.SelectedIndex))
            Dim in2 As String = stockID3(oUsed(ListBox3.SelectedIndex))
            Dim in3 As String = description(oUsed(ListBox3.SelectedIndex))
            Dim in4 As String = color(oUsed(ListBox3.SelectedIndex))
            Dim in5 As String = TextBox2.Text
            Dim in6 As String = TextBox3.Text
            Dim in7 As String = ""
            Dim in8 As String = ""
            Dim in9 As String = sawnumber(oUsed(ListBox3.SelectedIndex))
            Dim in10 As String = inputinternal
            Dim in11 As String = ""
            Dim validrow As Boolean = False
            Dim notempty As Boolean = False
            Dim found As Boolean = False

            found = False
            For it1 = 0 To stockID2.Count - 1
                If String.Equals(stockID2(it1), in1) Then
                    in7 = internalID(it1)
                    found = True
                End If
            Next
            If Not found Then
                Dim exitfrm As CannotFind = New CannotFind()
                exitfrm.ShowDialog()
            End If

            cmd.CommandText = "SELECT stockID2, size, count, context2  FROM stockUsed"
            cmd.ExecuteNonQuery()
            readerObj = cmd.ExecuteReader

            While readerObj.Read
                If String.Equals(in1, readerObj("stockID2").ToString) And String.Equals(in5, readerObj("size").ToString) Then
                    Dim temp1 As Integer = Convert.ToInt32(in6)
                    Dim temp2 As String = readerObj("count").ToString
                    Dim temp3 As Integer = Convert.ToInt32(temp2)
                    temp3 = temp3 - temp1
                    If temp3.ToString > 0 Then
                        cmd1.CommandText = "UPDATE stockUsed SET count = " + temp3.ToString + " WHERE context2 = " + readerObj("context2").ToString
                        cmd1.ExecuteNonQuery()
                    Else
                        cmd1.CommandText = "DELETE FROM stockUsed WHERE context2 = " + readerObj("context2").ToString
                        cmd1.ExecuteNonQuery()
                    End If
                End If
            End While
            readerObj.Close()
            Me.StockUsedTableAdapter.Fill(Me.OptimizationDatabaseDataSet3.stockUsed)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Label19.Visible = True
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con
        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, count, internalID, location, context1 FROM stockUsed"
        cmd.ExecuteNonQuery()

        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        Dim appXL As Excel.Application
        Dim wbXl As Excel.Workbook
        Dim shXL As Excel.Worksheet
        Dim raXL As Excel.Range
        ' Start Excel and get Application object.
        appXL = CreateObject("Excel.Application")
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
        ' AutoFit columns
        raXL = shXL.Range("A1", "I1")
        raXL.EntireColumn.AutoFit()
        ' Make sure Excel is visible and give the user control of Excel's lifetime.
        appXL.Visible = True
        appXL.UserControl = True
        ' Release object references.
        raXL = Nothing
        shXL = Nothing
        wbXl = Nothing
        appXL.Quit()
        appXL = Nothing
        GC.Collect()
        Label19.Visible = False
    End Sub

End Class