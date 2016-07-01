
Imports System.Data.SqlClient
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Public Class CutManagement
    Dim stockID1 As ArrayList = New ArrayList
    Dim stockID2 As ArrayList = New ArrayList
    Dim stockID3 As ArrayList = New ArrayList
    Dim description As ArrayList = New ArrayList
    Dim color As ArrayList = New ArrayList
    Dim size1 As ArrayList = New ArrayList
    Dim internalID As ArrayList = New ArrayList
    Dim Used As ArrayList = New ArrayList

    Dim cutstockID1 As ArrayList = New ArrayList
    Dim cutstockID2 As ArrayList = New ArrayList
    Dim cutstockID3 As ArrayList = New ArrayList
    Dim cutdescription As ArrayList = New ArrayList
    Dim cutcolor As ArrayList = New ArrayList
    Dim cutsize1 As ArrayList = New ArrayList
    Dim cutinternalID As ArrayList = New ArrayList
    Dim cutUsed As ArrayList = New ArrayList
    Dim cutcount As ArrayList = New ArrayList

    Dim opartID As ArrayList = New ArrayList
    Dim odescription As ArrayList = New ArrayList
    Dim ocolor As ArrayList = New ArrayList
    Dim osize As ArrayList = New ArrayList
    Dim ocount As ArrayList = New ArrayList
    Dim ointernalID As ArrayList = New ArrayList
    Dim oshopnumber As ArrayList = New ArrayList
    Dim oitemNumber As ArrayList = New ArrayList
    Dim oitemQuantity As ArrayList = New ArrayList
    Dim osetNumber As ArrayList = New ArrayList
    Dim oUsed As ArrayList = New ArrayList
    Dim oselect As ArrayList = New ArrayList

    Dim uopartID As ArrayList = New ArrayList
    Dim uodescription As ArrayList = New ArrayList
    Dim uocolor As ArrayList = New ArrayList
    Dim uosize As ArrayList = New ArrayList
    Dim uocount As ArrayList = New ArrayList
    Dim uointernalID As ArrayList = New ArrayList
    Dim uoshopnumber As ArrayList = New ArrayList
    Dim uoitemNumber As ArrayList = New ArrayList
    Dim uoitemQuantity As ArrayList = New ArrayList
    Dim uosetNumber As ArrayList = New ArrayList
    Dim uoUsed As ArrayList = New ArrayList

    Private Sub CutManagement_load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = "Easy Cut V1.0"

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con

        '
        ' Reads stock from database 
        '
        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, internalID FROM stockNew"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

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

        End While
        readerObj.Close()

        '
        ' Reads parts from database 
        '
        cmd.CommandText = "SELECT shopNumber FROM parts"
        cmd.ExecuteNonQuery()
        Dim readerObj1 As SqlClient.SqlDataReader = cmd.ExecuteReader

        While readerObj1.Read
            If Not ComboBox6.Items.Contains(readerObj1("shopNumber").ToString) Then
                ComboBox6.Items.Add(readerObj1("shopNumber").ToString)
            End If
        End While
        readerObj1.Close()

    End Sub

    '
    ' Page 1
    '

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
        If ListBox2.SelectedIndex >= 0 Then
            pString = description(Used(ListBox2.SelectedIndex)) + vbCrLf + color(Used(ListBox2.SelectedIndex)) + vbCrLf + size1(Used(ListBox2.SelectedIndex)) + vbCrLf + stockID1(Used(ListBox2.SelectedIndex)) + vbCrLf + stockID2(Used(ListBox2.SelectedIndex)) + vbCrLf + stockID3(Used(ListBox2.SelectedIndex))
            RichTextBox1.Text = pString
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cutdescription.Add(description(Used(ListBox2.SelectedIndex)))
        cutstockID1.Add(stockID1(Used(ListBox2.SelectedIndex)))
        cutstockID2.Add(stockID2(Used(ListBox2.SelectedIndex)))
        cutstockID3.Add(stockID3(Used(ListBox2.SelectedIndex)))
        cutcolor.Add(color(Used(ListBox2.SelectedIndex)))
        cutsize1.Add(TextBox2.Text)
        cutcount.Add(TextBox1.Text)
        cutinternalID.Add(internalID(Used(ListBox2.SelectedIndex)))
        cutUsed.Add(Used(ListBox2.SelectedIndex))

        Dim pString As String = ""
        For it = 0 To cutinternalID.Count - 1
            pString = pString + cutcolor(it) + " " + cutstockID1(it) + "  Size: " + cutsize1(it) + "  Count: " + cutcount(it) + vbCrLf
            RichTextBox2.Text = pString
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim pString As String = ""
        cutdescription.Clear()
        cutstockID1.Clear()
        cutstockID2.Clear()
        cutstockID3.Clear()
        cutcolor.Clear()
        cutsize1.Clear()
        cutcount.Clear()
        cutinternalID.Clear()
        cutUsed.Clear()
        RichTextBox2.Text = pString
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Calculator As CutGLib.CutEngine
        Calculator = New CutGLib.CutEngine
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
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
        Dim imagecount As Integer = 0
        Dim imagecountfile As ArrayList = New ArrayList

        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, count, internalID FROM stockNew"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        '
        ' Reads data from database 
        '

        Dim usedstockID1 As ArrayList = New ArrayList
        Dim usedstockID2 As ArrayList = New ArrayList
        Dim usedstockID3 As ArrayList = New ArrayList
        Dim useddescription As ArrayList = New ArrayList
        Dim usedcolor As ArrayList = New ArrayList
        Dim usedsize1 As ArrayList = New ArrayList
        Dim usedinternalID As ArrayList = New ArrayList
        Dim usedUsed As ArrayList = New ArrayList
        Dim usedcount As ArrayList = New ArrayList

        While readerObj.Read
            For it = 0 To cutUsed.Count - 1
                If (String.Equals(cutinternalID(it), readerObj("internalID").ToString)) Then
                    useddescription.Add(readerObj("description").ToString)
                    usedstockID1.Add(readerObj("stockID1").ToString)
                    usedstockID2.Add(readerObj("stockID2").ToString)
                    usedstockID3.Add(readerObj("stockID3").ToString)
                    usedcolor.Add(readerObj("color").ToString)
                    usedsize1.Add(readerObj("size").ToString)
                    usedinternalID.Add(readerObj("internalID").ToString)

                    Dim temp As Integer = Convert.ToInt32(cutcount(it))
                    Dim temp1 As Integer = Convert.ToInt32(readerObj("count").ToString)
                    If temp < temp1 Then
                        usedcount.Add(temp)
                    Else
                        usedcount.Add(readerObj("count").ToString)
                    End If
                End If
            Next
            If readerObj.IsDBNull(1) Then
                readerObj.Close()
                Exit While
            End If
        End While

        '
        'Initialize PDF
        '

        Dim document As PdfDocument = New PdfDocument
        document.Info.Title = "Cut Instructions"

        Dim font As XFont = New XFont("Verdana", 12, XFontStyle.Bold)
        Dim font2 As XFont = New XFont("Verdana", 14, XFontStyle.Bold)
        Dim font3 As XFont = New XFont("Verdana", 16, XFontStyle.Bold)
        Dim filename As String = "CutInstructions.pdf"
        '
        ' Adds parts to Calculator
        '

        For it = 0 To usedcount.Count - 1
            Dim temp2 As Double = Convert.ToInt32(usedsize1(it))
            Dim temp1 As Integer = Convert.ToDouble(usedcount(it))
            Calculator.AddLinearStock(temp2, temp1, usedinternalID(it))

        Next

        For it = 0 To cutcount.Count - 1
            Dim temp1 As Integer = Convert.ToInt32(cutcount(it))
            Dim temp2 As Double = Convert.ToDouble(cutsize1(it))
            Calculator.AddLinearPart(temp2, temp1, cutinternalID(it2))

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

            '
            'Printing output
            '

            Dim StockIndex, VStockCount, ViPart, iLayout, partCount, partIndex, tmp, iStock As Integer
            Dim partLength, VX, StockLength As Double
            Dim StockActive As Boolean
            Dim cutId1 As String = ""
            iStock = 0
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
                            Console.WriteLine(cutId1 + "------------------------")
                            Dim pageNo As String = Convert.ToString(iStock + 1)
                            Dim tempIt As String = Convert.ToString(iStock)
                            Dim page As PdfPage = document.AddPage
                            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                            Calculator.CreateStockImage(iStock, "test" + tempIt + ".png", 500)
                            Dim image As XImage = XImage.FromFile("test" + tempIt + ".png")
                            imagecount = tempIt
                            imagecountfile.Add(image)
                            Console.WriteLine("picture index: " + tempIt)
                            Dim pos1 As Integer = 0
                            Dim descriptionString As String = ""
                            Dim descriptionString2 As String = ""
                            Dim descriptionString3 As String = ""
                            Dim descriptionString4 As String = ""
                            Dim descriptionString5 As String = ""

                            For it = 0 To cutinternalID.Count - 1
                                If String.Equals(cutId1, cutinternalID(it)) Then
                                    descriptionString = cutdescription(it)
                                    descriptionString2 = "Color: " + cutcolor(it) + "  Size: " + cutsize1(it) + "''  "
                                    descriptionString3 = cutstockID1(it)
                                    descriptionString4 = cutstockID2(it)
                                    descriptionString5 = cutstockID3(it)
                                End If
                            Next

                            'ListBox3.Items.Add(cutId1)
                            gfx.DrawImage(image, 50, 700, 500, 30)
                            gfx.DrawString("Page " + pageNo, font, XBrushes.Black, New XRect(50, 50, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(descriptionString, font3, XBrushes.Black, New XRect(50, 65, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(descriptionString2, font2, XBrushes.Black, New XRect(50, 85, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(descriptionString3, font2, XBrushes.Black, New XRect(50, 100, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(descriptionString4, font2, XBrushes.Black, New XRect(50, 115, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(descriptionString5, font2, XBrushes.Black, New XRect(50, 130, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

                            Console.WriteLine("Stock={0}:  Length={1}", iStock, StockLength)

                            Dim Stockt As String = Convert.ToString(iStock)
                            Dim xt As String = Convert.ToString(StockLength)
                            'Dim Lengtht As String = Convert.ToString(cutLength(it2))
                            gfx.DrawString("Stock = " + Stockt + "  Length= " + xt + " ", font, XBrushes.Black, New XRect(50, 150 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
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
                                gfx.DrawString("Part = " + Stockt + "  x= " + vxt + " Length= " + vpartLength, font, XBrushes.Black, New XRect(50, 150 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                                pos1 += 20

                            Next ViPart
                        End If

                    Next iStock

                End If
            Next iLayout

            document.Save(filename)
            document.Close()
            Process.Start(filename)
            Calculator.Clear()
            For it = 0 To imagecount
                Dim temp As String = Convert.ToString(it)
                imagecountfile(it).Dispose()
                My.Computer.FileSystem.DeleteFile("test" + temp + ".png")
            Next

        Else
            Console.WriteLine("calculate fail")
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox1.Checked = True Then
            For it = 0 To ListBox3.Items.Count - 1
                ListBox3.SetSelected(it, True)
            Next
        Else
            For it = 0 To ListBox3.Items.Count - 1
                ListBox3.SetSelected(it, False)
            Next
        End If
    End Sub

    '
    ' Page 2
    '

    Private Sub ComboBox6_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedValueChanged
        ListBox3.Items.Clear()
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        cmd.CommandText = "SELECT partID, description, color, size, count, internalID, shopNumber, itemNumber, itemQuantity FROM parts"
        cmd.ExecuteNonQuery()
        Dim readerObj1 As SqlClient.SqlDataReader = cmd.ExecuteReader

        Dim it As Integer = 0
        While readerObj1.Read
            Dim tempSize As Double = Convert.ToDouble(readerObj1("Size").ToString)

            ListBox3.Items.Add(readerObj1("description").ToString)


            osetNumber.Add("ItemNumber: " + readerObj1("itemNumber").ToString + "item Quantity: " + readerObj1("itemQuantity").ToString)


            If String.Equals(ComboBox6.SelectedItem, readerObj1("shopNumber").ToString) And Not ComboBox5.Items.Contains(readerObj1("partID").ToString) Then
                ComboBox5.Items.Add(readerObj1("partID").ToString)
            End If
            If String.Equals(ComboBox6.SelectedItem, readerObj1("shopNumber").ToString) And Not ComboBox4.Items.Contains(readerObj1("description").ToString) Then
                ComboBox4.Items.Add(readerObj1("description").ToString)
            End If
            If String.Equals(ComboBox6.SelectedItem, readerObj1("shopNumber").ToString) And Not ComboBox7.Items.Contains(osetNumber(it)) Then
                ComboBox7.Items.Add(osetNumber(it))
            End If

            opartID.Add(readerObj1("partID").ToString)
            odescription.Add(readerObj1("description").ToString)
            ocolor.Add(readerObj1("color").ToString)
            osize.Add(readerObj1("size").ToString)
            ocount.Add(readerObj1("count").ToString)
            ointernalID.Add(readerObj1("internalID").ToString)
            oshopnumber.Add(readerObj1("shopNumber").ToString)
            oitemNumber.Add(readerObj1("itemNumber").ToString)
            oitemQuantity.Add(readerObj1("itemQuantity").ToString)
            oselect.add(it)
            it += 1
        End While
    End Sub

    Private Sub ComboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        ListBox3.Items.Clear()
        For it = 0 To ointernalID.Count - 1
            If String.Equals(ComboBox5.SelectedItem, opartID(it)) Then
                ListBox3.Items.Add(odescription(it))
            End If
        Next

    End Sub

    Private Sub ComboBox4_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedValueChanged
        ListBox3.Items.Clear()
        For it = 0 To ointernalID.Count - 1
            If String.Equals(ComboBox4.SelectedItem, odescription(it)) Then
                ListBox3.Items.Add(odescription(it))
            End If
        Next

    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ListBox3.Items.Clear()
        For it = 0 To ointernalID.Count - 1
            If String.Equals(ComboBox7.SelectedItem, osetNumber(it)) Then
                ListBox3.Items.Add(odescription(it))
            End If
        Next
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        Dim pString As String = ""
        If ListBox3.SelectedIndex >= 0 Then
            pString = odescription(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Color: " + ocolor(oselect(ListBox3.SelectedIndex)) + " Size: " + osize(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Count: " + ocount(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Shop Number: " + oshopnumber(oselect(ListBox3.SelectedIndex))
            RichTextBox4.Text = pString
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListBox3.SelectedIndex >= 0 Then
            oUsed.Add(ListBox3.SelectedIndex)
            ListBox1.Items.Add(ListBox3.SelectedItem)
            uopartID.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uodescription.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uocolor.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uosize.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uocount.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uointernalID.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uoshopnumber.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uoitemNumber.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uoitemQuantity.Add(opartID(oselect(ListBox3.SelectedIndex)))
            uoUsed.Add(oselect(ListBox3.SelectedIndex))

            ListBox3.Items.RemoveAt(ListBox3.SelectedIndex)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            While ListBox3.Items.Count > 0
                ListBox3.SelectedIndex = 0
                oUsed.Add(ListBox3.SelectedIndex)
                ListBox1.Items.Add(ListBox3.SelectedItem)

                uopartID.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uodescription.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uocolor.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uosize.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uocount.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uointernalID.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uoshopnumber.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uoitemNumber.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uoitemQuantity.Add(opartID(oselect(ListBox3.SelectedIndex)))
                uoUsed.Add(oselect(ListBox3.SelectedIndex))

                ListBox3.Items.RemoveAt(ListBox3.SelectedIndex)
            End While
        ElseIf CheckBox1.Checked = False Then
            oUsed.Clear()
            While ListBox1.Items.Count > 0
                ListBox1.SelectedIndex = 0
                ListBox3.Items.Add(ListBox1.SelectedItem)
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End While
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        While ListBox1.Items.Count > 0
            ListBox1.SelectedIndex = 0
            ListBox3.Items.Add(ListBox1.SelectedItem)
            ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        End While

        oUsed.Clear()
        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()
        uoUsed.Clear()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'uopartID.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uodescription.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uocolor.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uosize.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uocount.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uointernalID.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uoshopnumber.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uoitemNumber.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uoitemQuantity.Add(opartID(oselect(ListBox3.SelectedIndex)))
        'uoUsed.Add(oselect(ListBox3.SelectedIndex))

        For it = 0 To uointernalID.Count - 1
            If uodescription(it) = uodescription(it) Then

            End If
        Next


    End Sub
End Class