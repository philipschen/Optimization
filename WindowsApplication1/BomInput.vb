Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Public Class BomInput
    Dim connectionstring As Class1 = New Class1
    Private Sub this_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Easy Cut V1.0"
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
            Dim shXLName As ArrayList = New ArrayList
            ' Start Excel and get Application object.
            appXL = CreateObject("Excel.Application")
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
                                partID.Add(c0)
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
                                Console.WriteLine(tempcount)
                                If partsize(it1) <> 0 Then
                                    temp2 += partsize(it1) * partcount(it1)
                                    Console.WriteLine(temp2)
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
                    'Console.WriteLine(fpartcount(it))
                    Dim cID2 As ArrayList = New ArrayList
                    Dim cDescription As ArrayList = New ArrayList
                    Dim ccount As ArrayList = New ArrayList
                    Dim cuom As ArrayList = New ArrayList

                    cmd.CommandText = "SELECT ID1, ID2, UOM, Description, Count FROM Nav"
                    cmd.ExecuteNonQuery()
                    Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
                    While readerObj.Read
                        If (readerObj("ID2").ToString.Contains(fpartID(it))) Then
                            'OrElse readerObj("Description").ToString.Contains(partDescription(it))
                            cuom.Add(readerObj("UOM"))
                            cID2.Add(readerObj("ID2"))
                            cDescription.Add(readerObj("Description"))
                            ccount.Add(readerObj("Count"))
                        End If
                    End While
                    readerObj.Close()

                    If cID2.Count > 1 Then
                        Dim frm1 As chooseStock = New chooseStock()
                        Dim usedstock
                        frm1.des = cID2
                        frm1.id = cDescription
                        frm1.part1 = fpartDescription(it) + "  ID: " + fpartID(it)

                        If frm1.ShowDialog() = DialogResult.OK Then
                            usedstock = frm1.ListBox1.SelectedIndex
                            ID2.Add(cID2(usedstock))
                            Description.Add(cDescription(usedstock))
                            count.Add(ccount(usedstock))
                            uom.Add(cuom(usedstock))
                        End If
                    ElseIf cID2.Count = 1 Then
                        ID2.Add(cID2(0))
                        Description.Add(cDescription(0))
                        count.Add(ccount(0))
                        uom.Add(cuom(0))
                    Else
                        ID2.Add(fpartID(it))
                        Description.Add(fpartDescription(it))
                        count.Add(0)
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

                For it = 0 To amount.Count - 1
                    'If String.Equals(uom(it), "BAR") Then
                    If amount(it) < count(it) Then
                        enough.Add(True)
                        pstring2 = pstring2 + "In Stock: " + ID2(it) + "   " + Description(it) + vbCrLf
                        pstring2 = pstring2 + "   In stock: " + count(it).ToString + "  Need: " + amount(it).ToString + vbCrLf
                    Else
                        enough.Add(False)
                        Dim temp As Integer = amount(it) - count(it)
                        pstring = pstring + "Limited by: " + ID2(it) + "   " + Description(it) + vbCrLf
                        pstring = pstring + "   Instock: " + count(it).ToString + "  Need: " + amount(it).ToString + "  Order: " + temp.ToString + vbCrLf
                    End If

                Next

                RichTextBox1.Text = pstring
                RichTextBox2.Text = pstring2
                Label1.Text = "Loading Complete"


                shXL = Nothing
                wbXl = Nothing
                appXL.Quit()
                appXL = Nothing
            End If
        End If

    End Sub

End Class