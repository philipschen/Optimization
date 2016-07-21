Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form1
    Dim connectionstring As Class1 = New Class1
    Dim ar0 As ArrayList = New ArrayList
    Dim ar1 As ArrayList = New ArrayList
    Dim ar2 As ArrayList = New ArrayList
    Dim ar3 As ArrayList = New ArrayList
    Dim ar4 As ArrayList = New ArrayList
    Dim ar5 As ArrayList = New ArrayList
    Dim ar6 As ArrayList = New ArrayList
    Dim ar7 As ArrayList = New ArrayList
    Dim ar8 As ArrayList = New ArrayList
    Private Sub this_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = connectionstring.version
        Button3.Visible = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ar0.Clear()
        ar1.Clear()
        ar2.Clear()
        ar3.Clear()
        ar4.Clear()
        ar5.Clear()
        ar6.Clear()
        ar7.Clear()
        ar8.Clear()
        DataGridView1.Rows.Clear()
        ListBox1.Items.Clear()
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""

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
            'appXL.DisplayAlerts = False
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

            If Not wbXl Is Nothing Then wbXl.Close()
            appXL.Quit()

            shXL = Nothing
            wbXl = Nothing
            appXL = Nothing
            GC.Collect()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView1.RowCount < 2 Then
            MessageBox.Show("First Open Inventory NAV")
        Else
            ar0.Clear()
            ar1.Clear()
            ar2.Clear()
            ar3.Clear()
            ar4.Clear()
            ar5.Clear()
            ar6.Clear()
            ar7.Clear()
            ar8.Clear()
            ListBox1.Items.Clear()
            RichTextBox1.Text = ""
            RichTextBox2.Text = ""
            RichTextBox3.Text = ""
            Label1.Text = "Upload Complete, Go to Step 2"
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
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ar1.Count = 0 Then
            MessageBox.Show("First Enter Inventory NAV")
        Else
            ListBox1.Items.Clear()
            RichTextBox1.Text = ""
            RichTextBox2.Text = ""
            RichTextBox3.Text = ""

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
                Dim shXLName As ArrayList = New ArrayList
                ' Start Excel and get Application object.
                appXL = CreateObject("Excel.Application")
                'appXL.DisplayAlerts = False
                'appXL.Visible = True
                ' Add a new workbook.
                wbXl = appXL.Workbooks.Open(strFileName)
                'shXL = wbXl.ActiveSheet
                For Each shXL In wbXl.Worksheets
                    shXLName.Add(shXL.Name)
                Next shXL

                Dim partID As ArrayList = New ArrayList
                Dim partDescription As ArrayList = New ArrayList
                Dim partcount As ArrayList = New ArrayList
                Dim partsize As ArrayList = New ArrayList

                Dim fpartID As ArrayList = New ArrayList
                Dim fpartDescription As ArrayList = New ArrayList
                Dim fpartcount As ArrayList = New ArrayList
                Dim fpartsize As ArrayList = New ArrayList
                Dim ftotalsize As ArrayList = New ArrayList

                Dim ID2 As ArrayList = New ArrayList
                Dim Description As ArrayList = New ArrayList
                Dim count As ArrayList = New ArrayList
                Dim uom As ArrayList = New ArrayList
                Dim amount As ArrayList = New ArrayList
                Dim enough As ArrayList = New ArrayList

                '
                ' Populates list of all parts needed
                '

                Dim frm2 As Amount = New Amount
                If frm2.ShowDialog = DialogResult.OK Then
                    Label7.Text = "Starting Calculation..."
                    Dim amountmade As Integer = 0
                    amountmade = frm2.amount
                    For it1 = 0 To shXLName.Count - 1
                        Dim start1 As Boolean = False
                        shXL = wbXl.Sheets(shXLName(it1))
                        shXL.Activate()
                        Dim result1 As DialogResult = MessageBox.Show("Does this order include: " + shXLName(it1), "Confirm parts", MessageBoxButtons.YesNo)

                        If result1 = DialogResult.Yes Then
                            Dim totalcount As String = Convert.ToString(shXL.UsedRange.Rows.Count - 1)
                            For it = 0 To shXL.UsedRange.Rows.Count - 1
                                Dim c0 As String = shXL.Cells(1 + it, 1).Value
                                Dim c1 As String = shXL.Cells(1 + it, 2).Value
                                Dim c2 As String = shXL.Cells(1 + it, 3).Value
                                Dim c3 As String = shXL.Cells(1 + it, 6).Value

                                If c0 IsNot Nothing AndAlso c0.Contains("Part") Then
                                    start1 = True
                                End If
                                If start1 AndAlso IsNumeric(c2) Then
                                    If c0 IsNot Nothing Then
                                        partID.Add(c0)
                                    Else
                                        partID.Add("NO-ID")
                                    End If
                                    partDescription.Add(c1)
                                    partcount.Add(c2)

                                    If IsNumeric(c3) Then
                                        partsize.Add(c3)
                                    Else
                                        partsize.Add(0)
                                    End If
                                End If
                            Next
                        End If
                    Next
                    '
                    ' Merges Duplicates
                    '
                    For it = 0 To partID.Count - 1
                        If Not fpartID.Contains(partID(it)) Then
                            fpartID.Add(partID(it))
                            fpartDescription.Add(partDescription(it))
                            Dim tempcount As Integer = 0
                            Dim temp2 As Double = 0
                            For it1 = 0 To partID.Count - 1
                                If String.Equals(partID(it), partID(it1)) Then
                                    tempcount += partcount(it1)

                                    If partsize(it1) <> 0 Then
                                        temp2 += partsize(it1) * partcount(it1)

                                    End If
                                End If
                            Next
                            fpartcount.Add(tempcount)
                            ftotalsize.Add(temp2)
                        End If
                    Next
                    '
                    ' Grabs all stock used
                    '
                    For it = 0 To fpartID.Count - 1
                        Dim cID2 As ArrayList = New ArrayList
                        Dim cDescription As ArrayList = New ArrayList
                        Dim ccount As ArrayList = New ArrayList
                        Dim cuom As ArrayList = New ArrayList
                        For it1 = 0 To ar1.Count - 1
                            If ar1(it1).contains(fpartID(it)) Then
                                cID2.Add(ar1(it1))
                                cDescription.Add(ar3(it1))
                                ccount.Add(ar5(it1))
                                cuom.Add(ar2(it1))
                            End If
                        Next
                        If cID2.Count > 1 Then
                            Dim frm1 As chooseStock = New chooseStock()
                            Dim usedstock
                            frm1.id = cID2
                            frm1.des = cDescription
                            frm1.part1 = fpartDescription(it) + "  ID: " + fpartID(it)

                            If frm1.ShowDialog() = DialogResult.OK Then
                                usedstock = frm1.ListBox1.SelectedIndex
                                ID2.Add(cID2(usedstock))
                                Description.Add(cDescription(usedstock))
                                count.Add(ccount(usedstock))
                                uom.Add(cuom(usedstock))
                            Else
                                ID2.Add(fpartID(it))
                                Description.Add(fpartDescription(it))
                                count.Add(-1)
                                uom.Add("")
                            End If
                        ElseIf cID2.Count = 1 Then
                            ID2.Add(cID2(0))
                            Description.Add(cDescription(0))
                            count.Add(ccount(0))
                            uom.Add(cuom(0))
                        Else
                            ID2.Add(fpartID(it))
                            Description.Add(fpartDescription(it))
                            count.Add(-1)
                            uom.Add("")
                        End If
                    Next

                    For it = 0 To ID2.Count - 1
                        ListBox1.Items.Add(ID2(it) + "  " + Description(it))
                    Next

                    For it = 0 To ID2.Count - 1
                        If ftotalsize(it) = 0 Or Description(it).Contains("GLASS") Then
                            amount.Add(amountmade * fpartcount(it))
                        ElseIf String.Equals(uom(it), "BAR") Then
                            amount.Add(Math.Round(amountmade * ftotalsize(it) / 150) + 1)
                        Else
                            amount.Add(Math.Round(amountmade * ftotalsize(it)))
                        End If
                    Next

                    Dim pstring As String = ""
                    Dim pstring2 As String = ""
                    Dim pstring3 As String = ""

                    For it = 0 To amount.Count - 1
                        If amount(it) < count(it) Then
                            enough.Add(True)
                            pstring2 = pstring2 + "In Stock: " + ID2(it) + "   " + Description(it) + vbCrLf
                            pstring2 = pstring2 + "   In stock: " + count(it).ToString + "  Need: " + amount(it).ToString + vbCrLf
                        ElseIf count(it) <> -1 Then
                            enough.Add(False)
                            Dim temp As Integer = amount(it) - count(it)
                            pstring = pstring + "Limited by: " + ID2(it) + "   " + Description(it) + vbCrLf
                            pstring = pstring + "   Instock: " + count(it).ToString + "  Need: " + amount(it).ToString + "  Order: " + temp.ToString + vbCrLf
                        Else
                            enough.Add(False)
                            Dim temp As Integer = amount(it) - count(it) - 1
                            pstring3 = pstring3 + "Not in NAV: " + ID2(it) + "   " + Description(it) + vbCrLf
                            pstring3 = pstring3 + "  Need: " + amount(it).ToString + "  Order: " + temp.ToString + vbCrLf
                        End If
                    Next

                    RichTextBox1.Text = pstring
                    RichTextBox2.Text = pstring2
                    RichTextBox3.Text = pstring3
                    Label7.Text = "Calculation Complete"

                    'If Not wbXl Is Nothing Then wbXl.Close()
                    appXL.Quit()

                    shXL = Nothing
                    wbXl = Nothing
                    appXL = Nothing
                    GC.Collect()
                End If
            End If
        End If

    End Sub

End Class