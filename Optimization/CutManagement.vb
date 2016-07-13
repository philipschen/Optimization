Imports System.Data.SqlClient
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Public Class CutManagement
    ' page 1
    ' all from Newstock table
    Dim stockID1 As ArrayList = New ArrayList
    Dim stockID2 As ArrayList = New ArrayList
    Dim stockID3 As ArrayList = New ArrayList
    Dim description As ArrayList = New ArrayList
    Dim color As ArrayList = New ArrayList
    Dim size1 As ArrayList = New ArrayList
    Dim internalID As ArrayList = New ArrayList
    Dim Used As ArrayList = New ArrayList
    ' cut stock
    Dim cutstockID1 As ArrayList = New ArrayList
    Dim cutstockID2 As ArrayList = New ArrayList
    Dim cutstockID3 As ArrayList = New ArrayList
    Dim cutdescription As ArrayList = New ArrayList
    Dim cutcolor As ArrayList = New ArrayList
    Dim cutsize1 As ArrayList = New ArrayList
    Dim cutinternalID As ArrayList = New ArrayList
    Dim cutUsed As ArrayList = New ArrayList
    Dim cutcount As ArrayList = New ArrayList

    ' page 2
    ' all from newparts table
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
    ' everything in the listbox
    Dim lopartID As ArrayList = New ArrayList
    Dim lodescription As ArrayList = New ArrayList
    Dim locolor As ArrayList = New ArrayList
    Dim losize As ArrayList = New ArrayList
    Dim locount As ArrayList = New ArrayList
    Dim lointernalID As ArrayList = New ArrayList
    Dim loshopnumber As ArrayList = New ArrayList
    Dim loitemNumber As ArrayList = New ArrayList
    Dim loitemQuantity As ArrayList = New ArrayList
    Dim losetNumber As ArrayList = New ArrayList
    Dim loUsed As ArrayList = New ArrayList
    Dim loselect As ArrayList = New ArrayList
    ' calculation parts list
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
        ' removes single cuts
        TabControl1.TabPages.RemoveAt(0)

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

        Dim it As Integer = 0
        Dim it2 As Integer = 0
        Dim itWorks As Boolean = False

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
            itWorks = True
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
            If itWorks Then
                document.Save(filename)
                document.Close()
                Process.Start(filename)
                Calculator.Clear()

                For it = 0 To imagecount
                    Dim temp As String = Convert.ToString(it)
                    imagecountfile(it).Dispose()
                    My.Computer.FileSystem.DeleteFile("test" + temp + ".png")
                Next
            End If
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
        ListBox1.Items.Clear()
        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()


        lopartID.Clear()
        lodescription.Clear()
        locolor.Clear()
        losize.Clear()
        locount.Clear()
        lointernalID.Clear()
        loshopnumber.Clear()
        loitemNumber.Clear()
        loitemQuantity.Clear()
        loselect.Clear()
        ListBox3.Items.Clear()
        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
        con.Open()
        cmd.Connection = con
        cmd.CommandText = "SELECT partID, description, color, size, count, internalID, shopNumber, itemNumber, itemQuantity FROM parts ORDER BY itemNumber, itemQuantity, description"
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
            oselect.Add(it)

            lopartID.Add(opartID(it))
            lodescription.Add(odescription(it))
            locolor.Add(ocolor(it))
            losize.Add(osize(it))
            locount.Add(ocount(it))
            lointernalID.Add(ointernalID(it))
            loshopnumber.Add(oshopnumber(it))
            loitemNumber.Add(oitemNumber(it))
            loitemQuantity.Add(oitemQuantity(it))
            loselect.Add(it)

            it += 1
        End While
    End Sub

    Private Sub ComboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        ListBox1.Items.Clear()
        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()

        lopartID.Clear()
        lodescription.Clear()
        locolor.Clear()
        losize.Clear()
        locount.Clear()
        lointernalID.Clear()
        loshopnumber.Clear()
        loitemNumber.Clear()
        loitemQuantity.Clear()
        loselect.Clear()
        ListBox3.Items.Clear()
        For it = 0 To ointernalID.Count - 1
            If String.Equals(ComboBox5.SelectedItem, opartID(it)) Then
                ListBox3.Items.Add(odescription(it))

                lopartID.Add(opartID(it))
                lodescription.Add(odescription(it))
                locolor.Add(ocolor(it))
                losize.Add(osize(it))
                locount.Add(ocount(it))
                lointernalID.Add(ointernalID(it))
                loshopnumber.Add(oshopnumber(it))
                loitemNumber.Add(oitemNumber(it))
                loitemQuantity.Add(oitemQuantity(it))
                loselect.Add(it)
            End If
        Next

    End Sub

    Private Sub ComboBox4_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedValueChanged
        ListBox1.Items.Clear()
        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()

        lopartID.Clear()
        lodescription.Clear()
        locolor.Clear()
        losize.Clear()
        locount.Clear()
        lointernalID.Clear()
        loshopnumber.Clear()
        loitemNumber.Clear()
        loitemQuantity.Clear()
        loselect.Clear()
        ListBox3.Items.Clear()
        For it = 0 To ointernalID.Count - 1
            If String.Equals(ComboBox4.SelectedItem, odescription(it)) Then
                ListBox3.Items.Add(odescription(it))

                lopartID.Add(opartID(it))
                lodescription.Add(odescription(it))
                locolor.Add(ocolor(it))
                losize.Add(osize(it))
                locount.Add(ocount(it))
                lointernalID.Add(ointernalID(it))
                loshopnumber.Add(oshopnumber(it))
                loitemNumber.Add(oitemNumber(it))
                loitemQuantity.Add(oitemQuantity(it))
                loselect.Add(it)
            End If
        Next

    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ListBox1.Items.Clear()
        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()

        lopartID.Clear()
        lodescription.Clear()
        locolor.Clear()
        losize.Clear()
        locount.Clear()
        lointernalID.Clear()
        loshopnumber.Clear()
        loitemNumber.Clear()
        loitemQuantity.Clear()
        loselect.Clear()
        ListBox3.Items.Clear()
        For it = 0 To ointernalID.Count - 1
            If String.Equals(ComboBox7.SelectedItem, osetNumber(it)) Then
                ListBox3.Items.Add(odescription(it))

                lopartID.Add(opartID(it))
                lodescription.Add(odescription(it))
                locolor.Add(ocolor(it))
                losize.Add(osize(it))
                locount.Add(ocount(it))
                lointernalID.Add(ointernalID(it))
                loshopnumber.Add(oshopnumber(it))
                loitemNumber.Add(oitemNumber(it))
                loitemQuantity.Add(oitemQuantity(it))
                loselect.Add(it)
            End If
        Next
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        'ListBox1.ClearSelected()
        Dim pString As String = ""
        If ListBox3.SelectedIndex >= 0 Then
            pString = odescription(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Color: " + ocolor(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Size: " + osize(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Count: " + ocount(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Shop Number: " + oshopnumber(oselect(ListBox3.SelectedIndex))
            RichTextBox4.Text = pString
        End If
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        'ListBox3.ClearSelected()
        Dim pString As String = ""
        If ListBox1.SelectedIndex >= 0 Then
            pString = uodescription(ListBox1.SelectedIndex) + vbCrLf + "Color: " + uocolor(ListBox1.SelectedIndex) + vbCrLf + " Size: " + uosize(ListBox1.SelectedIndex) + vbCrLf + "Count: " + uocount(ListBox1.SelectedIndex) + vbCrLf + "Shop Number: " + uoshopnumber(ListBox1.SelectedIndex)
            RichTextBox4.Text = pString
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListBox3.SelectedIndex >= 0 Then
            oUsed.Add(ListBox3.SelectedIndex)
            ListBox1.Items.Add(ListBox3.SelectedItem)

            uopartID.Add(lopartID(ListBox3.SelectedIndex))
            uodescription.Add(lodescription(ListBox3.SelectedIndex))
            uocolor.Add(locolor(ListBox3.SelectedIndex))
            uosize.Add(losize(ListBox3.SelectedIndex))
            uocount.Add(locount(ListBox3.SelectedIndex))
            uointernalID.Add(lointernalID(ListBox3.SelectedIndex))
            uoshopnumber.Add(loshopnumber(ListBox3.SelectedIndex))
            uoitemNumber.Add(loitemNumber(ListBox3.SelectedIndex))
            uoitemQuantity.Add(loitemQuantity(ListBox3.SelectedIndex))
            uoUsed.Add(ListBox3.SelectedIndex)

            lopartID.RemoveAt(ListBox3.SelectedIndex)
            lodescription.RemoveAt(ListBox3.SelectedIndex)
            locolor.RemoveAt(ListBox3.SelectedIndex)
            losize.RemoveAt(ListBox3.SelectedIndex)
            locount.RemoveAt(ListBox3.SelectedIndex)
            lointernalID.RemoveAt(ListBox3.SelectedIndex)
            loshopnumber.RemoveAt(ListBox3.SelectedIndex)
            loitemNumber.RemoveAt(ListBox3.SelectedIndex)
            loitemQuantity.RemoveAt(ListBox3.SelectedIndex)
            loselect.RemoveAt(ListBox3.SelectedIndex)
            ListBox3.Items.RemoveAt(ListBox3.SelectedIndex)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            For it = 0 To ListBox3.Items.Count - 1
                ListBox3.SelectedIndex = it
                oUsed.Add(ListBox3.SelectedIndex)
                ListBox1.Items.Add(ListBox3.SelectedItem)

                uopartID.Add(lopartID(ListBox3.SelectedIndex))
                uodescription.Add(lodescription(ListBox3.SelectedIndex))
                uocolor.Add(locolor(ListBox3.SelectedIndex))
                uosize.Add(losize(ListBox3.SelectedIndex))
                uocount.Add(locount(ListBox3.SelectedIndex))
                uointernalID.Add(lointernalID(ListBox3.SelectedIndex))
                uoshopnumber.Add(loshopnumber(ListBox3.SelectedIndex))
                uoitemNumber.Add(loitemNumber(ListBox3.SelectedIndex))
                uoitemQuantity.Add(loitemQuantity(ListBox3.SelectedIndex))
                uoUsed.Add(oUsed(ListBox3.SelectedIndex))
            Next
            ListBox3.Items.Clear()
        ElseIf CheckBox1.Checked = False Then
            oUsed.Clear()
            While ListBox1.Items.Count > 0
                ListBox1.SelectedIndex = 0
                ListBox3.Items.Add(ListBox1.SelectedItem)
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End While
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
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'While ListBox1.Items.Count > 0
        '    ListBox1.SelectedIndex = 0
        '    ListBox3.Items.Add(ListBox1.SelectedItem)
        '    ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
        'End While
        ListBox3.Items.Clear()
        ListBox1.Items.Clear()

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
        Dim partcountIndex As ArrayList = New ArrayList
        Dim partlistd As ArrayList = New ArrayList
        Dim partlistid As ArrayList = New ArrayList
        Dim partlistcount As ArrayList = New ArrayList
        Dim partlistcolor As ArrayList = New ArrayList
        Dim itWorks As Boolean = False
        partcountIndex.Add(0)
        Dim itpart As Integer = 0
        For it = 0 To uopartID.Count - 1
            If Not partlistd.Contains(uodescription(it)) Then
                partlistd.Add(uodescription(it))
                partlistid.Add(uopartID(it))

                Dim tempcount As Integer = 0
                For it1 = 0 To uodescription.Count - 1
                    If String.Equals(partlistd(itpart), uodescription(it1)) Then
                        tempcount += uocount(it1)
                    End If
                Next
                partlistcount.Add(tempcount)
                partlistcolor.Add(uocolor(it))
                itpart += 1
                If it > 0 Then
                    partcountIndex.Add(it)
                End If
            End If
        Next
        '
        'Initialize PDF
        '


        Dim clock As DateTime = New DateTime
        clock = My.Computer.Clock.LocalTime
        Dim imagecountfile As ArrayList = New ArrayList
        Dim picturecount As Integer = 0
        Dim document As PdfDocument = New PdfDocument
        document.Info.Title = "Cut Instructions"


        Dim font As XFont = New XFont("Verdana", 12, XFontStyle.Regular)
        Dim font2 As XFont = New XFont("Verdana", 12, XFontStyle.Regular)
        Dim font3 As XFont = New XFont("Verdana", 14, XFontStyle.Regular)
        Dim filename As String = clock.Year.ToString + clock.Month.ToString + clock.Day.ToString + clock.Minute.ToString + clock.Second.ToString + "CutInstructions.pdf"

        '
        'Loops for individual types
        '

        Dim exdlist As ArrayList = New ArrayList
        Dim excountlist As ArrayList = New ArrayList
        Dim excountlistu As ArrayList = New ArrayList
        Dim excountwork As ArrayList = New ArrayList
        For it = 0 To partlistd.Count - 1

            Label14.Text = "Processing Request..."
            Dim Calculator As CutGLib.CutEngine
            Calculator = New CutGLib.CutEngine
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = "Data Source=TOSHIBA-2015\SQLEXPRESS;Initial Catalog=OptimizationDatabase;Integrated Security=True"
            con.Open()
            cmd.Connection = con

            cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, count, internalID, context1 FROM stockNew"
            cmd.ExecuteNonQuery()
            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

            Dim usedstockID1 As ArrayList = New ArrayList
            Dim usedstockID2 As ArrayList = New ArrayList
            Dim usedstockID3 As ArrayList = New ArrayList
            Dim useddescription As ArrayList = New ArrayList
            Dim usedcolor As ArrayList = New ArrayList
            Dim usedsize1 As ArrayList = New ArrayList
            Dim usedinternalID As ArrayList = New ArrayList
            Dim usedUsed As ArrayList = New ArrayList
            Dim usedcount As ArrayList = New ArrayList
            Dim usedsaw As ArrayList = New ArrayList
            Dim stock_exists As Boolean = False

            Dim xusedsize1 As ArrayList = New ArrayList
            Dim xusedcount As ArrayList = New ArrayList
            Dim xusedinternalID As ArrayList = New ArrayList
            Dim xusedcontext2 As ArrayList = New ArrayList


            '
            ' Finds Corresponding Stocks
            '
            While readerObj.Read
                If (readerObj("description").ToString.Contains(partlistd(it)) Or readerObj("stockID2").ToString.Contains(partlistid(it))) Then
                    usedstockID1.Add(readerObj("stockID1").ToString)
                    usedstockID2.Add(readerObj("stockID2").ToString)
                    usedstockID3.Add(readerObj("stockID3").ToString)
                    useddescription.Add(readerObj("description").ToString)
                    usedcolor.Add(readerObj("color").ToString)
                    usedsize1.Add(readerObj("size").ToString)
                    usedcount.Add(readerObj("count").ToString)
                    usedinternalID.Add(readerObj("internalID").ToString)
                    usedsaw.Add(readerObj("context1").ToString)
                    stock_exists = True
                End If
            End While
            readerObj.Close()


            '
            ' Decides Which Stock to Use Then inserts
            '
            Dim usedstock As Integer = 0
            If usedinternalID.Count > 1 Then
                Dim frm1 As ChooseStock = New ChooseStock()
                frm1.des = useddescription
                frm1.id = usedstockID2
                frm1.color = usedcolor
                frm1.saw = usedsaw
                frm1.part1 = partlistd(it) + "  ID: " + partlistid(it) + "  Color: " + partlistcolor(it)

                If frm1.ShowDialog() = DialogResult.OK Then
                    usedstock = frm1.ListBox1.SelectedIndex
                    usedsaw = frm1.saw
                    cmd.CommandText = "UPDATE stockNew SET context1 = '" + usedsaw(usedstock) + "' WHERE internalID = '" + usedinternalID(usedstock) + "'"
                    cmd.ExecuteNonQuery()

                    cmd.CommandText = "SELECT size, count, internalID, context2 FROM stockUsed"
                    cmd.ExecuteNonQuery()
                    readerObj = cmd.ExecuteReader
                    While readerObj.Read
                        If String.Equals(usedinternalID(usedstock), readerObj("internalID").ToString) Then
                            xusedsize1.Add(readerObj("size").ToString)
                            xusedcount.Add(readerObj("count").ToString)
                            xusedinternalID.Add(readerObj("internalID").ToString)
                            xusedcontext2.Add(readerObj("context2").ToString)

                        End If
                    End While
                    readerObj.Close()
                End If

            End If

            If xusedinternalID.Count > 0 Then
                For it1 = 0 To xusedinternalID.Count - 1
                    Dim temp3 As Double = Convert.ToDouble(xusedsize1(it1))
                    Dim temp4 As Integer = Convert.ToInt32(xusedcount(it1))
                    Calculator.AddLinearStock(temp3, temp4)

                Next
            End If
            If usedinternalID.Count > 0 Then
                Dim temp1 As Double = Convert.ToDouble(usedsize1(usedstock))
                Dim temp2 As Integer = Convert.ToInt32(usedcount(usedstock))
                Calculator.AddLinearStock(temp1, temp2)

            Else

                '
                ' Adds Page if can't find stock
                '

                Dim page As PdfPage = document.AddPage
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                gfx.DrawString("No stock/is not a cut extrusion: " + partlistd(it), font2, XBrushes.Black, New XRect(50, 50, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                Label14.Text = "Calculate Fail, No Stock"
                excountlist.Add(partlistcount(it))
                excountlistu.Add(0)
                excountwork.Add(1)
            End If

            '
            ' Loop for inserting cut parts
            '

            For it1 = 0 To uodescription.Count - 1
                If String.Equals(partlistd(it), uodescription(it1)) Then

                    Dim temp3 As Double = Convert.ToDouble(uosize(it1))
                    Dim temp4 As Integer = Convert.ToInt32(uocount(it1))

                    Calculator.AddLinearPart(temp3, temp4)
                End If
            Next

            '
            'Calculates Cuts 
            '

            Dim result As String = "1"
            If stock_exists Then
                result = Calculator.ExecuteLinear()

            End If
            '
            'Formats Outputs
            '
            If (result = "") Then
                itWorks = True

                '
                'Printing output
                '
                Dim remainder As Double = 0
                Dim StockIndex, VStockCount, ViPart, iLayout, partCount, partIndex, tmp, iStock, sectioncount As Integer
                Dim partLength, VX, StockLength As Double
                Dim StockActive As Boolean
                iStock = 0
                sectioncount = 0
                Dim pos2 As Integer = 0
                Console.WriteLine("Created {0} different layouts", Calculator.LayoutCount)
                ' Iterate by each layout and output information about each layout,
                ' such as number and length of used stocks and part indices cut from the stocks

                Dim itemcount As Integer = 0
                Dim itemcount2 As Integer = 0
                For iLayout = 0 To Calculator.LayoutCount - 1
                    Calculator.GetLayoutInfo(iLayout, StockIndex, VStockCount)
                    ' StockIndex is global index of the first stock used in the layout iLayout
                    ' VStockCount is quantity of stocks of the same length as StockIndex us                                        If VStockCount > 0 Then
                    Dim slength As Double
                    Calculator.GetLinearStockInfo(StockIndex, slength, StockActive)
                    If String.Equals(slength.ToString, usedsize1(usedstock)) Then
                        itemcount += VStockCount
                    End If
                    If Not String.Equals(slength.ToString, usedsize1(usedstock)) Then
                        itemcount2 += VStockCount
                    End If

                    Console.WriteLine("Layout={0}:  Start Stock={1};  Count of Stock={2}", iLayout, StockIndex, VStockCount)

                    '
                    ' Add Page
                    '
                    Dim over20 As Boolean = False
                    Dim page As PdfPage
                    Dim gfx As XGraphics
                    partCount = Calculator.GetPartCountOnStock(iStock)
                    If partCount > 20 Then
                        page = document.AddPage
                        gfx = XGraphics.FromPdfPage(page)
                        over20 = True
                        font = New XFont("Verdana", 11, XFontStyle.Regular)
                        sectioncount = 0

                    ElseIf partCount > 12 Then
                        page = document.AddPage
                        gfx = XGraphics.FromPdfPage(page)
                        pos2 = 0
                        sectioncount = 0
                        font = New XFont("Verdana", 12, XFontStyle.Regular)
                    Else
                        font = New XFont("Verdana", 12, XFontStyle.Regular)
                        If sectioncount = 0 Then
                            page = document.AddPage
                            gfx = XGraphics.FromPdfPage(page)
                            pos2 = 0
                            sectioncount += 1
                        Else
                            pos2 = 350
                            sectioncount = 0
                        End If

                    End If

                    iStock = StockIndex
                    If Calculator.GetLinearStockInfo(StockIndex, StockLength, StockActive) Then
                        partCount = Calculator.GetPartCountOnStock(iStock)


                        Dim pageNo As String = Convert.ToString(iLayout + 1)
                        Dim tempIt As String = Convert.ToString(iStock)
                        ' Dim page As PdfPage = document.AddPage


                        Dim temppic As String = Convert.ToString(picturecount)
                        Calculator.CreateStockImage(iStock, "test" + temppic + ".png", 1000)
                        Dim image As XImage = XImage.FromFile("test" + temppic + ".png")
                        picturecount += 1
                        imagecountfile.Add(image)

                        Dim Stockt As String = Convert.ToString(iStock + 1)
                        Dim xt As String = Convert.ToString(StockLength)
                        Dim temp As String = Convert.ToString(VStockCount)
                        Dim temp1 As String = Convert.ToString(iLayout + 1)
                        Dim pos1 As Integer = 0
                        Dim sizestring As String = xt



                        If Not String.Equals(slength.ToString, usedsize1(usedstock)) Then
                            sizestring = xt + "  Used Part Check Bin"
                        End If

                        Dim descriptionString As String = partlistd(it)
                        Dim descriptionString2 As String = "ID: " + usedstockID2(usedstock)
                        Dim descriptionString3 As String = "ID2: " + usedstockID3(usedstock)
                        Dim descriptionString4 As String = "Stock Size: " + sizestring + "  Color: " + usedcolor(usedstock)
                        Dim descriptionstring5 As String = "Saw number: " + usedsaw(usedstock)
                        Dim descriptionstring6 As String = "# Of Stock Cut: " + temp
                        Dim descriptionstring7 As String = "Layout Number = " + temp1
                        Dim leftanchor As Integer = 50



                        'ListBox3.Items.Add(cutId1)

                        gfx.DrawString("Page " + pageNo, font2, XBrushes.Black, New XRect(leftanchor, 50 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString, font3, XBrushes.Black, New XRect(leftanchor, 65 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString2, font2, XBrushes.Black, New XRect(leftanchor, 80 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString3, font2, XBrushes.Black, New XRect(leftanchor, 95 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString4, font2, XBrushes.Black, New XRect(leftanchor, 110 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionstring5, font3, XBrushes.Black, New XRect(leftanchor + 360, 50 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionstring6, font3, XBrushes.Black, New XRect(leftanchor + 360, 65 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionstring7, font2, XBrushes.Black, New XRect(leftanchor + 360, 80 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

                        Console.WriteLine("Layout={0}:  Length={1}", iStock, StockLength)


                        ' Output the information about parts cut from this stock
                        ' First we get quantity of parts cut from the stock

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
                            Dim vxt As String = Convert.ToString(VX + partLength)
                            Dim vpartLength As String = Convert.ToString(partLength)
                            gfx.DrawString("Cut position= " + vxt + " Length= " + vpartLength, font, XBrushes.Black, New XRect(leftanchor, 140 + pos1 + pos2, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            If over20 Then
                                pos1 += 11
                            Else
                                pos1 += 15
                            End If
                            If VX + partLength > remainder Then
                                remainder = VX + partLength
                            End If
                        Next ViPart

                        gfx.DrawImage(image, 50, 160 + pos1 + pos2, 500, 20)



                        If usedsize1(usedstock) - remainder > 12 Then
                            Dim internalID As ArrayList = New ArrayList
                            Dim rn As New Random
                            Dim it1 As Integer
                            Dim boolinternalID As Boolean = True
                            Dim inputusedID As Integer

                            cmd.CommandText = "SELECT context2  FROM stockUsed"
                            cmd.ExecuteNonQuery()

                            'This will loop through all returned records 
                            readerObj = cmd.ExecuteReader
                            While readerObj.Read
                                Dim temp4 As Integer = Convert.ToInt64(readerObj("context2").ToString)
                                internalID.Add(temp4)
                            End While
                            readerObj.Close()
                            inputusedID = rn.Next(1000000, 9999999)
                            For it1 = 0 To internalID.Count - 1
                                If internalID(it1) = inputusedID Then
                                    it1 = 0
                                    inputusedID = rn.Next(1000000, 9999999)
                                End If
                            Next

                            Dim temp3 = Convert.ToString(inputusedID)
                            Dim temp2 As String = Convert.ToString(usedsize1(usedstock) - remainder)
                            cmd.CommandText = "INSERT INTO stockUsed VALUES('" + usedstockID1(usedstock) + "', '" + usedstockID2(usedstock) + "' , '" + usedstockID3(usedstock) + "', '" + useddescription(usedstock) + "' , '" + usedcolor(usedstock) + "', " + temp2 + ", " + temp + ", " + usedinternalID(usedstock) + " , '' , '" + usedsaw(usedstock) + "' , " + temp3 + ", '')"
                            cmd.ExecuteNonQuery()
                        End If
                    End If

                Next iLayout
                excountlist.Add(itemcount)
                excountlistu.Add(itemcount2)
                excountwork.Add(0)

            ElseIf stock_exists Then

                Console.WriteLine("calculate fail")
                Dim page As PdfPage = document.AddPage
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                gfx.DrawString("Not enough stock: " + partlistd(it), font2, XBrushes.Black, New XRect(50, 50, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                excountlist.Add(usedcount(usedstock))
                excountlistu.Add(0)
                excountwork.Add(2)
            Else

            End If

            Calculator.Clear()

        Next

        If itWorks Then
            Dim page As PdfPage = document.AddPage
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim pos1 As Integer = 0
            gfx.DrawString("Total Part List", font3, XBrushes.Black, New XRect(50, 50, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
            For it = 0 To partlistd.Count - 1
                If excountwork(it) = 0 Then
                    gfx.DrawString((it + 1).ToString + ". New Stock: " + excountlist(it).ToString + " " + "Used Stock: " + excountlistu(it).ToString + " " + partlistd(it), font2, XBrushes.Black, New XRect(50, 70 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    pos1 += 15
                ElseIf excountwork(it) = 1 Then
                    gfx.DrawString((it + 1).ToString + ". Not Cut, Partcount: " + excountlist(it).ToString + "  " + partlistd(it), font2, XBrushes.Black, New XRect(50, 70 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    pos1 += 15
                ElseIf excountwork(it) = 2 Then
                    gfx.DrawString((it + 1).ToString + ". Not Enough Stock, current Stock: " + excountlist(it).ToString + "  " + partlistd(it), font2, XBrushes.Black, New XRect(50, 70 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    pos1 += 15
                End If

            Next

            Label14.Text = "Request Complete"
            document.Save(filename)
            document.Close()
            Process.Start(filename)

            For it = 0 To picturecount - 1
                Dim temp As String = Convert.ToString(it)
                imagecountfile(it).Dispose()
                My.Computer.FileSystem.DeleteFile("test" + temp + ".png")
            Next

        End If
    End Sub
End Class