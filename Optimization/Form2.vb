Imports System.Data.SqlClient
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf

Public Class Form2

    Protected partID As ArrayList = New ArrayList()
    Protected partColor As ArrayList = New ArrayList()
    Protected partSize As ArrayList = New ArrayList()
    Protected partCount As ArrayList = New ArrayList()
    Protected partComboID As ArrayList = New ArrayList()


    Private Sub Form2_Load(sender As Object, e As EventArgs) _
   Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"

        CheckBox1.Checked = True
        ListBox2.Items.Add("Bronze")
        ListBox2.Items.Add("White")
        ListBox2.Items.Add("Silver")

        ListBox1.SelectedIndex = 0
        ListBox2.SelectedIndex = 0


    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ListBox1.Items.Add("THX-01-001")
            ListBox1.Items.Add("THX-01-01F")
            ListBox1.Items.Add("THX-01-01XL")
            ListBox1.Items.Add("THX-01-01N")
            ListBox1.Items.Add("THX-01-002")
            ListBox1.Items.Add("THX-01-02F")
            ListBox1.Items.Add("THX-01-02L")
            ListBox1.Items.Add("THX-01-02N")
            ListBox1.Items.Add("THX-01-02S")
            ListBox1.Items.Add("THX-01-03A")
            ListBox1.Items.Add("THX-01-03AF")
            ListBox1.Items.Add("THX-01-03XLA")
            ListBox1.Items.Add("THX-01-3AN")
            ListBox1.Items.Add("THX-99-02B")
            ListBox1.Items.Add("THX-01-05A")
            ListBox1.Items.Add("THX-01-004")
            ListBox1.Items.Add("THX-99-003")
            ListBox1.Items.Add("THX-01-007A")
            ListBox1.Items.Add("THX-01-06A")
            ListBox1.Items.Add("THX-01-08A")
            ListBox1.Items.Add("THX-01-03AEL")
            ListBox1.Items.Add("THX-01-009")
            ListBox1.Items.Add("THX-99-15")
            ListBox1.Items.Add("THX-99-15A")
            ListBox1.Items.Add("P01-THX001")
            ListBox1.Items.Add("2000-A01-000000")
        End If

        If CheckBox1.Checked = False Then
            ListBox1.Items.Remove("THX-01-001")
            ListBox1.Items.Remove("THX-01-01F")
            ListBox1.Items.Remove("THX-01-01XL")
            ListBox1.Items.Remove("THX-01-01N")
            ListBox1.Items.Remove("THX-01-002")
            ListBox1.Items.Remove("THX-01-02F")
            ListBox1.Items.Remove("THX-01-02L")
            ListBox1.Items.Remove("THX-01-02N")
            ListBox1.Items.Remove("THX-01-02S")
            ListBox1.Items.Remove("THX-01-03A")
            ListBox1.Items.Remove("THX-01-03AF")
            ListBox1.Items.Remove("THX-01-03XLA")
            ListBox1.Items.Remove("THX-01-3AN")
            ListBox1.Items.Remove("THX-99-02B")
            ListBox1.Items.Remove("THX-01-05A")
            ListBox1.Items.Remove("THX-01-004")
            ListBox1.Items.Remove("THX-99-003")
            ListBox1.Items.Remove("THX-01-007A")
            ListBox1.Items.Remove("THX-01-06A")
            ListBox1.Items.Remove("THX-01-08A")
            ListBox1.Items.Remove("THX-01-03AEL")
            ListBox1.Items.Remove("THX-01-009")
            ListBox1.Items.Remove("THX-99-15")
            ListBox1.Items.Remove("THX-99-15A")
            ListBox1.Items.Remove("P01-THX001")
            ListBox1.Items.Remove("2000-A01-000000")
        End If

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            ListBox1.Items.Add("THX-03A-01")
            ListBox1.Items.Add("THX-03A-01F")
            ListBox1.Items.Add("THX-03A-01XL")
            ListBox1.Items.Add("THX-03A-1N")
            ListBox1.Items.Add("THX-03A-1L")
            ListBox1.Items.Add("THX-03A-04")
            ListBox1.Items.Add("THX-03A-06")
            ListBox1.Items.Add("THX-03A-06A")
            ListBox1.Items.Add("THX-03A-03")
            ListBox1.Items.Add("THX-03A-IFN-X")
            ListBox1.Items.Add("THX-03A-1NC")
            ListBox1.Items.Add("THX-03A-01")
            ListBox1.Items.Add("THX-03A-001")
            ListBox1.Items.Add("THX-03A-003")
        End If

        If CheckBox2.Checked = False Then
            ListBox1.Items.Remove("THX-03A-01")
            ListBox1.Items.Remove("THX-03A-01F")
            ListBox1.Items.Remove("THX-03A-01XL")
            ListBox1.Items.Remove("THX-03A-1N")
            ListBox1.Items.Remove("THX-03A-1L")
            ListBox1.Items.Remove("THX-03A-04")
            ListBox1.Items.Remove("THX-03A-06")
            ListBox1.Items.Remove("THX-03A-06A")
            ListBox1.Items.Remove("THX-03A-03")
            ListBox1.Items.Remove("THX-03A-IFN-X")
            ListBox1.Items.Remove("THX-03A-1NC")
            ListBox1.Items.Remove("THX-03A-01")
            ListBox1.Items.Remove("THX-03A-001")
            ListBox1.Items.Remove("THX-03A-003")
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            ListBox1.Items.Add("TDX-12-01")
            ListBox1.Items.Add("TDX-12-02")
            ListBox1.Items.Add("TDX-12-01")
            ListBox1.Items.Add("TDX-12-43")
            ListBox1.Items.Add("TDX-12-06")
            ListBox1.Items.Add("TDX-12-47")
            ListBox1.Items.Add("TDX-12-48A")
            ListBox1.Items.Add("TDX-12-09")
            ListBox1.Items.Add("TDX-12-04")
            ListBox1.Items.Add("TDX-12-05")
            ListBox1.Items.Add("TDX-12-10A")
            ListBox1.Items.Add("TDX-12-A02")
            ListBox1.Items.Add("TDX-12-44")
            ListBox1.Items.Add("TDX-12-A08A")
            ListBox1.Items.Add("SNX-005")
            ListBox1.Items.Add("SNX-005L")
            ListBox1.Items.Add("1240-572-000325")
            ListBox1.Items.Add("TDX-12-A08")
            ListBox1.Items.Add("TDX-12-17")
            ListBox1.Items.Add("TDX-12-11B")
        End If

        If CheckBox3.Checked = False Then
            ListBox1.Items.Remove("TDX-12-01")
            ListBox1.Items.Remove("TDX-12-02")
            ListBox1.Items.Remove("TDX-12-01")
            ListBox1.Items.Remove("TDX-12-43")
            ListBox1.Items.Remove("TDX-12-06")
            ListBox1.Items.Remove("TDX-12-47")
            ListBox1.Items.Remove("TDX-12-48A")
            ListBox1.Items.Remove("TDX-12-09")
            ListBox1.Items.Remove("TDX-12-04")
            ListBox1.Items.Remove("TDX-12-05")
            ListBox1.Items.Remove("TDX-12-10A")
            ListBox1.Items.Remove("TDX-12-A02")
            ListBox1.Items.Remove("TDX-12-44")
            ListBox1.Items.Remove("TDX-12-A08A")
            ListBox1.Items.Remove("SNX-005")
            ListBox1.Items.Remove("SNX-005L")
            ListBox1.Items.Remove("1240-572-000325")
            ListBox1.Items.Remove("TDX-12-A08")
            ListBox1.Items.Remove("TDX-12-17")
            ListBox1.Items.Remove("TDX-12-11B")
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            ListBox1.Items.Add("TDX-12-801")
            ListBox1.Items.Add("TDX-12-816")
            ListBox1.Items.Add("TDX-12-804")
            ListBox1.Items.Add("TDX-12-802")
            ListBox1.Items.Add("TDX-12-803")
            ListBox1.Items.Add("TDX-12-809")
            ListBox1.Items.Add("TDX-12-805")
            ListBox1.Items.Add("TDX-12-806B")
            ListBox1.Items.Add("TDX-12-807A")
            ListBox1.Items.Add("TDX-12-808")
            ListBox1.Items.Add("TDX-12-810")
            ListBox1.Items.Add("THX-62-19")
            ListBox1.Items.Add("TDX-12-811")
            ListBox1.Items.Add("TDX-12-814")
            ListBox1.Items.Add("TDX-12-16")
            ListBox1.Items.Add("TDX-12-817")
            ListBox1.Items.Add("1280-ASTR-01")
        End If

        If CheckBox4.Checked = False Then
            ListBox1.Items.Remove("TDX-12-801")
            ListBox1.Items.Remove("TDX-12-816")
            ListBox1.Items.Remove("TDX-12-804")
            ListBox1.Items.Remove("TDX-12-802")
            ListBox1.Items.Remove("TDX-12-803")
            ListBox1.Items.Remove("TDX-12-809")
            ListBox1.Items.Remove("TDX-12-805")
            ListBox1.Items.Remove("TDX-12-806B")
            ListBox1.Items.Remove("TDX-12-807A")
            ListBox1.Items.Remove("TDX-12-808")
            ListBox1.Items.Remove("TDX-12-810")
            ListBox1.Items.Remove("THX-62-19")
            ListBox1.Items.Remove("TDX-12-811")
            ListBox1.Items.Remove("TDX-12-814")
            ListBox1.Items.Remove("TDX-12-16")
            ListBox1.Items.Remove("TDX-12-817")
            ListBox1.Items.Remove("1280-ASTR-01")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim temp1 As String = Convert.ToString(TextBox2.Text)
        Dim temp2 As String = Convert.ToString(TextBox1.Text)
        Dim it1 As String = ""
        Dim it2 As Integer = 0
        Dim repeat As Boolean = False

        For Each it1 In partID
            If (String.Equals(partID(it2), ListBox1.SelectedItems(0)) And (String.Equals(partColor(it2), ListBox2.SelectedItems(0))) And (String.Equals(partSize(it2), temp2))) Then
                Dim tempCount As Integer = Convert.ToInt32(partCount(it2))
                tempCount += TextBox2.Text
                partCount.RemoveAt(it2)
                Dim StempCount As String = Convert.ToString(tempCount)
                partCount.Add(StempCount)
                repeat = True
            End If
            it2 += 1
        Next
        If Not repeat Then
            partID.Add(ListBox1.SelectedItems(0))
            partColor.Add(ListBox2.SelectedItems(0))
            partSize.Add(temp2)
            partCount.Add(temp1)
            partComboID.Add(ListBox1.SelectedItems(0) + " " + ListBox2.SelectedItems(0))
        End If

        Dim it As Integer
        Dim pString As String = ""

        For it = 0 To partID.Count - 1
            Dim tempSize As String = Convert.ToString(partSize(it))
            Dim tempCount As String = Convert.ToString(partCount(it))
            pString = pString + partID(it) + "  " + partColor(it) + "  Size: " + tempSize + "  Count: " + tempCount + vbCrLf
            RichTextBox1.Text = pString
        Next

        ListBox3.Items.Clear()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim pString As String = ""
        partID.Clear()
        partColor.Clear()
        partSize.Clear()
        partCount.Clear()
        RichTextBox1.Text = pString
        ListBox3.Items.Clear()
        PictureBox2.Visible = False

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Calculator As CutGLib.CutEngine
        Calculator = New CutGLib.CutEngine
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=Parts;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        Dim it1 As String = ""
        Dim it As Integer = 0
        Dim it2 As Integer = 0
        Dim stockID As ArrayList = New ArrayList()
        Dim stockColor As ArrayList = New ArrayList()
        Dim stockSize As ArrayList = New ArrayList()
        Dim stockCount As ArrayList = New ArrayList()
        Dim stockComboID As ArrayList = New ArrayList()
        Dim cutID As ArrayList = New ArrayList()
        Dim cutStock As ArrayList = New ArrayList()
        Dim cutPos As ArrayList = New ArrayList()
        Dim cutLength As ArrayList = New ArrayList()
        Dim cutPicID As ArrayList = New ArrayList()
        Dim pagecount As Integer

        cmd.CommandText = "SELECT stockID, Color, Size, Count FROM Parts"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        '
        ' Reads data from database 
        '

        While readerObj.Read
            For Each it1 In partID
                Dim tempSize As Integer = Convert.ToInt32(readerObj("Size").ToString)

                If (String.Equals(partID(it), readerObj("stockID").ToString)) And (String.Equals(partColor(it), readerObj("Color").ToString) And tempSize > partSize(it)) Then
                    stockID.Add(readerObj("stockID").ToString)
                    stockColor.Add(readerObj("Color").ToString)
                    stockSize.Add(readerObj("Size").ToString)
                    stockCount.Add(readerObj("Count").ToString)
                    stockComboID.Add(readerObj("stockID").ToString + " " + readerObj("Color").ToString)
                    it += 1
                End If
            Next
            If readerObj.IsDBNull(1) Then
                readerObj.Close()
                Exit While
            End If
        End While

        '
        ' Adds parts to Calculator
        '

        it1 = ""
        For Each it1 In stockID
            Dim temp1 As Integer = Convert.ToInt32(stockCount(it2))
            Dim temp2 As Double = Convert.ToInt32(stockSize(it2))

            Calculator.AddLinearStock(temp2, temp1, stockComboID(it2))
            it2 += 1
        Next
        it1 = ""
        it2 = 0
        For Each it1 In partID
            Dim temp1 As Integer = Convert.ToInt32(partCount(it2))
            Dim temp2 As Double = Convert.ToInt32(partSize(it2))

            Calculator.AddLinearPart(temp2, temp1, partComboID(it2))
            it2 += 1
        Next

        it1 = ""
        it2 = 0
        For Each it1 In partID
            Console.WriteLine("part: " + partID(it2) + "  " + partColor(it2) + "  Size: " + partSize(it2) + "  Count: " + partCount(it2))
            it2 += 1
        Next
        it1 = ""
        it2 = 0
        For Each it1 In stockID
            Console.WriteLine("stock: " + stockID(it2) + "  " + stockColor(it2) + "  Size: " + stockSize(it2) + "  Count: " + stockCount(it2) + vbCrLf)
            it2 += 1
        Next

        '
        'Calculates Cuts 
        '

        Dim result As String
        result = Calculator.ExecuteLinear()

        '
        'Formats Outputs
        '

        If (result = "") Then
            Dim StockNo As Integer
            Dim Length As Double
            Dim X As Double
            Dim iPart As Integer
            Dim id As String = ""

            Dim document As PdfDocument = New PdfDocument
            document.Info.Title = "Cut Instructions"

            Dim font As XFont = New XFont("Verdana", 14, XFontStyle.Bold)
            Dim font2 As XFont = New XFont("Verdana", 16, XFontStyle.Bold)
            Dim filename As String = "CutInstructions.pdf"

            'For iPart = 0 To Calculator.PartCount - 1
            '    ' Get sizes and location of the part with index Iter
            '    ' in case of incomplete optimization the function returns FALSE
            '    Dim tempPart As String = Convert.ToString(iPart)

            '    If Calculator.GetResultLinearPart(iPart, StockNo, Length, X, id) Then
            '        cutStock.Add(StockNo)
            '        cutLength.Add(Length)
            '        cutPos.Add(X)
            '        cutID.Add(id)

            '        'Console.WriteLine("Part ID={0};  stock={1};  X={2}  Length={3} ", cutID(iPart), cutStock(iPart), cutPos(iPart), cutLength(iPart))

            '        'If StockNo > pagecount Then
            '        'pagecount = StockNo
            '        'End If

            '    Else
            '        Console.WriteLine("Source part {0} was Not placed", iPart)
            '    End If

            '    Dim tempPart1 As String = Convert.ToString(iPart)


            'Next iPart


            '
            'Printing output
            '

            Dim StockIndex, VStockCount, ViPart, iLayout, partCount, partIndex, tmp, iStock As Integer
            Dim partLength, VX, StockLength As Double
            Dim StockActive As Boolean
            Dim cutId1 As String = ""

            Console.WriteLine("Created {0} different layouts", Calculator.LayoutCount)
            ' Iterate by each layout and output information about each layout,
            ' such as number and length of used stocks and part indices cut from the stocks
            For iLayout = 0 To Calculator.LayoutCount - 1
                Calculator.GetLayoutInfo(iLayout, StockIndex, VStockCount)
                ' StockIndex is global index of the first stock used in the layout iLayout
                ' VStockCount is quantity of stocks of the same length as StockIndex used for this layout
                If VStockCount > 0 Then

                    Console.WriteLine("Layout={0}:  Start Stock={1};  Count of Stock={2}", iLayout, StockIndex, VStockCount)
                    ' Output information about each stock, such as stock Length
                    For iStock = StockIndex To StockIndex + VStockCount - 1
                        If Calculator.GetLinearStockInfo(iStock, StockLength, StockActive, cutId1) Then

                            Dim pageNo As String = Convert.ToString(iStock + 1)
                            Dim tempIt As String = Convert.ToString(iStock)
                            Dim page As PdfPage = document.AddPage
                            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                            Calculator.CreateStockImage(iStock, "test" + tempIt + ".png", 500)
                            Dim image As XImage = XImage.FromFile("test" + tempIt + ".png")
                            Console.WriteLine("picture index: " + tempIt)
                            Dim pos1 As Integer = 0

                            ListBox3.Items.Add(cutId1)
                            gfx.DrawImage(image, 50, 700, 500, 30)
                            gfx.DrawString("Page " + pageNo, font, XBrushes.Black, New XRect(50, 50, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(cutId1, font2, XBrushes.Black, New XRect(50, 70, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

                            Console.WriteLine("Stock={0}:  Length={1}", iStock, StockLength)

                            Dim Stockt As String = Convert.ToString(iStock)
                            Dim xt As String = Convert.ToString(StockLength)
                            'Dim Lengtht As String = Convert.ToString(cutLength(it2))
                            gfx.DrawString("Stock = " + Stockt + "  Length= " + xt + " ", font, XBrushes.Black, New XRect(50, 90 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            pos1 += 20
                            ' Output the information about parts cut from this stock
                            ' First we get quantity of parts cut from the stock
                            partCount = Calculator.GetPartCountOnStock(iStock)
                            ' Iterate by parts and get indices of cut parts
                            For ViPart = 0 To partCount - 1
                                ' Get global part index of ViPart cut from the current stock
                                partIndex = Calculator.GetPartIndexOnStock(iStock, ViPart)
                                ' Get length and location of the part
                                ' X - coordinate on the stock where the part beggins.

                                Calculator.GetResultLinearPart(partIndex, tmp, partLength, VX)
                                ' Output the part information
                                Console.WriteLine("Part= {0}:  X={1}  Length={2}", partIndex, VX, partLength)

                                Dim partt As String = Convert.ToString(partIndex)
                                Dim vxt As String = Convert.ToString(VX)
                                Dim vpartLength As String = Convert.ToString(partLength)
                                gfx.DrawString("Part = " + Stockt + "  x= " + vxt + " Length= " + vpartLength, font, XBrushes.Black, New XRect(50, 90 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                                pos1 += 20

                            Next ViPart
                        End If
                    Next iStock
                End If
            Next iLayout

            document.Save(filename)
            Process.Start(filename)

        Else
            Console.WriteLine("calculate fail")
        End If

    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        Dim iSelector As Integer = ListBox3.SelectedIndex()
        Dim TiSelector As String = Convert.ToString(iSelector)
        PictureBox2.Image = Image.FromFile("test" + TiSelector + ".png")
    End Sub

End Class