Imports System.Data.SqlClient
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Public Class CutManagement
    Dim connectionstring As Class1 = New Class1
    ' page 2
    ' lists from newparts table
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
    Dim olistorderline As ArrayList = New ArrayList
    ' lists of everything in the listbox
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
    Dim lolistorderline As ArrayList = New ArrayList
    ' list of parts used in calculation
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
    Dim uolistorderline As ArrayList = New ArrayList

    Private Sub CutManagement_load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.  
        Me.Text = connectionstring.version
        ' Removes single cuts tab
        TabControl1.TabPages.RemoveAt(0)

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

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
    ' Page 2: Cut by Shop Oder
    '

    Private Sub ComboBox6_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedValueChanged
        'parts to be cut
        ListBox1.Items.Clear()
        ListBox3.Items.Clear()

        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()
        uolistorderline.Clear()

        opartID.Clear()
        odescription.Clear()
        ocolor.Clear()
        osize.Clear()
        ocount.Clear()
        ointernalID.Clear()
        oshopnumber.Clear()
        oitemNumber.Clear()
        oitemQuantity.Clear()
        oselect.Clear()
        olistorderline.Clear()
        'parts in list
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
        lolistorderline.Clear()

        ComboBox5.Items.Clear()
        ComboBox4.Items.Clear()
        ComboBox7.Items.Clear()
        osetNumber.Clear()

        Dim con As New SqlConnection
        Dim cmd As New SqlCommand
        con.ConnectionString = connectionstring.connect1
        con.Open()
        cmd.Connection = con

        cmd.CommandText = "SELECT partID, description, color, size, count, internalID, shopNumber, itemNumber, itemQuantity, context1 FROM parts ORDER BY context1, itemNumber"
        cmd.ExecuteNonQuery()
        Dim readerObj1 As SqlClient.SqlDataReader = cmd.ExecuteReader

        'imports shop order parts
        Dim it As Integer = 0
        While readerObj1.Read
            If String.Equals(ComboBox6.SelectedItem, readerObj1("shopNumber").ToString) Then
                Dim tempSize As Double = Convert.ToDouble(readerObj1("Size").ToString)
                osetNumber.Add("Quote:" + readerObj1("context1").ToString + " Line :" + readerObj1("itemNumber").ToString + " Count :" + readerObj1("itemQuantity").ToString)

                ListBox3.Items.Add(readerObj1("description").ToString)

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
                oitemQuantity.Add(readerObj1("itemQuantity"))
                oselect.Add(it)
                olistorderline.Add(readerObj1("context1").ToString)


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
                lolistorderline.Add(olistorderline(it))

                it += 1
            End If
        End While
    End Sub

    Private Sub ComboBox5_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedValueChanged
        ListBox1.Items.Clear()
        ListBox3.Items.Clear()

        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()
        uolistorderline.Clear()

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
        lolistorderline.Clear()

        'Adds parts by partID
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
                lolistorderline.Add(olistorderline(it))
            End If
        Next

    End Sub

    Private Sub ComboBox4_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedValueChanged
        ListBox1.Items.Clear()
        ListBox3.Items.Clear()

        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()
        uolistorderline.Clear()

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
        lolistorderline.Clear()

        ' Adds parts by Description
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
                lolistorderline.Add(olistorderline(it))
            End If
        Next

    End Sub
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ListBox1.Items.Clear()
        ListBox3.Items.Clear()

        uopartID.Clear()
        uodescription.Clear()
        uocolor.Clear()
        uosize.Clear()
        uocount.Clear()
        uointernalID.Clear()
        uoshopnumber.Clear()
        uoitemNumber.Clear()
        uoitemQuantity.Clear()
        uolistorderline.Clear()

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
        lolistorderline.Clear()

        ' Adds parts by shopitem number
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
                lolistorderline.Add(olistorderline(it))
            End If
        Next
    End Sub

    Private Sub ListBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox3.SelectedIndexChanged
        Dim pString As String = ""

        ' Sets item full description
        If ListBox3.SelectedIndex >= 0 Then
            pString = odescription(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Color: " + ocolor(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Size: " + osize(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Count: " + ocount(oselect(ListBox3.SelectedIndex)) + vbCrLf + "Shop Number: " + oshopnumber(oselect(ListBox3.SelectedIndex))
            RichTextBox4.Text = pString
        End If
    End Sub
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim pString As String = ""
        ' Sets item full description
        If ListBox1.SelectedIndex >= 0 Then
            pString = uodescription(ListBox1.SelectedIndex) + vbCrLf + "Color: " + uocolor(ListBox1.SelectedIndex) + vbCrLf + "Size: " + uosize(ListBox1.SelectedIndex) + vbCrLf + "Count: " + uocount(ListBox1.SelectedIndex) + vbCrLf + "Shop Number: " + uoshopnumber(ListBox1.SelectedIndex)
            RichTextBox4.Text = pString
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListBox3.SelectedIndex >= 0 Then
            oUsed.Add(ListBox3.SelectedIndex)
            ListBox1.Items.Add(ListBox3.SelectedItem)

            'Clears the list of parts to cuts and repopulates list of cut parts
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
            uolistorderline.Add(lolistorderline(ListBox3.SelectedIndex))

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
            lolistorderline.RemoveAt(ListBox3.SelectedIndex)

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
                uolistorderline.Add(lolistorderline(ListBox3.SelectedIndex))
            Next

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
            lolistorderline.Clear()

            ListBox3.Items.Clear()
        ElseIf CheckBox1.Checked = False Then
            oUsed.Clear()
            While ListBox1.Items.Count > 0
                ListBox1.SelectedIndex = 0
                ListBox3.Items.Add(ListBox1.SelectedItem)
                ListBox1.Items.RemoveAt(ListBox1.SelectedIndex)
            End While
            For it = 0 To uopartID.Count - 1
                lopartID.Add(uopartID(it))
                lodescription.Add(uodescription(it))
                locolor.Add(uocolor(it))
                losize.Add(uosize(it))
                locount.Add(uocount(it))
                lointernalID.Add(uointernalID(it))
                loshopnumber.Add(uoshopnumber(it))
                loitemNumber.Add(uoitemNumber(it))
                loitemQuantity.Add(uoitemQuantity(it))
                loselect.Add(it)
                lolistorderline.Add(uolistorderline(it))
            Next
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
            uolistorderline.Clear()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Clears All Cuts
        ListBox3.Items.Clear()
        ListBox1.Items.Clear()

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
        lolistorderline.Clear()

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
        uolistorderline.Clear()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim uocountdownrep As ArrayList = New ArrayList
        For it = 0 To uocount.Count - 1
            uocountdownrep.Add(uocount(it) / uoitemQuantity(it))
        Next
        Dim uocountdown As ArrayList = New ArrayList
        For it = 0 To uoitemQuantity.Count - 1
            uocountdown.Add(1)
        Next
        '
        ' lists of nonduplicate, combined parts
        '
        Dim partcountIndex As ArrayList = New ArrayList
        Dim partlistd As ArrayList = New ArrayList
        Dim partlistid As ArrayList = New ArrayList
        Dim partlistcount As ArrayList = New ArrayList
        Dim partlistcolor As ArrayList = New ArrayList
        Dim partsinternalID As ArrayList = New ArrayList
        Dim itWorks As Boolean = False
        partcountIndex.Add(0)
        Dim itpart As Integer = 0
        '
        ' Removes duplicate parts
        '
        For it = 0 To uopartID.Count - 1
            partsinternalID.Add(uointernalID(it))
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


        Dim font As XFont = New XFont("Verdana", 11, XFontStyle.Regular)
        Dim font2 As XFont = New XFont("Verdana", 12, XFontStyle.Regular)
        Dim font3 As XFont = New XFont("Verdana", 13, XFontStyle.Regular)
        Dim filename As String = clock.Year.ToString + clock.Month.ToString + clock.Day.ToString + clock.Minute.ToString + clock.Second.ToString + "CutInstructions.pdf"

        '
        ' Lists for database updates
        '
        Dim newstockinternalID As ArrayList = New ArrayList
        Dim usedstockinternalID As ArrayList = New ArrayList
        Dim usedstockinternalamount As ArrayList = New ArrayList
        Dim usedstockinternalIDcreated As ArrayList = New ArrayList

        '
        ' Lists for pdf output
        '
        Dim excountlist As ArrayList = New ArrayList
        Dim excountlistu As ArrayList = New ArrayList
        Dim excountwork As ArrayList = New ArrayList
        Dim excountused As ArrayList = New ArrayList

        '
        ' List for which used parts to remove
        '
        Dim usedcontext2 As ArrayList = New ArrayList
        Dim usedcontextamount As ArrayList = New ArrayList


        ' List of new stock that is used
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
        Dim usedminsize As ArrayList = New ArrayList
        Dim stock_exists As ArrayList = New ArrayList


        ' List of used stock that is used
        Dim xusedsize1 As ArrayList = New ArrayList
        Dim xusedcount As ArrayList = New ArrayList
        Dim xusedinternalID As ArrayList = New ArrayList
        Dim xusedcontext2 As ArrayList = New ArrayList
        '
        ' Populates list of corresponding new and used stocks
        '
        For it = 0 To partlistd.Count - 1
            ' List of possible new stocks
            Dim pusedstockID1 As ArrayList = New ArrayList
            Dim pusedstockID2 As ArrayList = New ArrayList
            Dim pusedstockID3 As ArrayList = New ArrayList
            Dim puseddescription As ArrayList = New ArrayList
            Dim pusedcolor As ArrayList = New ArrayList
            Dim pusedsize1 As ArrayList = New ArrayList
            Dim pusedinternalID As ArrayList = New ArrayList
            Dim pusedUsed As ArrayList = New ArrayList
            Dim pusedcount As ArrayList = New ArrayList
            Dim pusedsaw As ArrayList = New ArrayList
            Dim pusedminsize As ArrayList = New ArrayList

            stock_exists.Add(False)
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = connectionstring.connect1
            con.Open()
            cmd.Connection = con
            '
            ' Finds Corresponding new Stocks
            '
            cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size, count, internalID, context1, context2 FROM stockNew"
            cmd.ExecuteNonQuery()
            Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
            While readerObj.Read
                If (readerObj("description").ToString.Contains(partlistd(it)) Or readerObj("stockID2").ToString.Contains(partlistid(it))) Then
                    pusedstockID1.Add(readerObj("stockID1").ToString)
                    pusedstockID2.Add(readerObj("stockID2").ToString)
                    pusedstockID3.Add(readerObj("stockID3").ToString)
                    puseddescription.Add(readerObj("description").ToString)
                    pusedcolor.Add(readerObj("color").ToString)
                    pusedsize1.Add(readerObj("size").ToString)
                    pusedcount.Add(readerObj("count").ToString)
                    pusedinternalID.Add(readerObj("internalID").ToString)
                    pusedsaw.Add(readerObj("context1").ToString)
                    pusedminsize.Add(readerObj("context2").ToString)

                End If
            End While
            readerObj.Close()

            '
            ' Decides Which new Stock to Use and inserts saw if applicable
            '
            Dim pusedstock As Integer = 0
            If pusedinternalID.Count > 1 Then
                Dim frm1 As ChooseStock = New ChooseStock()
                frm1.des = puseddescription
                frm1.id = pusedstockID2
                frm1.color = pusedcolor
                frm1.saw = pusedsaw
                frm1.part1 = partlistd(it) + "  ID: " + partlistid(it) + "  Color: " + partlistcolor(it)

                If frm1.ShowDialog() = DialogResult.OK Then
                    pusedstock = frm1.ListBox1.SelectedIndex
                    pusedsaw = frm1.saw
                    cmd.CommandText = "UPDATE stockNew SET context1 = '" + pusedsaw(pusedstock) + "' WHERE internalID = '" + pusedinternalID(pusedstock) + "'"
                    cmd.ExecuteNonQuery()
                    '
                    ' Adds chosen stock to used new stock list
                    '
                    usedstockID1.Add(pusedstockID1(pusedstock))
                    usedstockID2.Add(pusedstockID2(pusedstock))
                    usedstockID3.Add(pusedstockID3(pusedstock))
                    useddescription.Add(puseddescription(pusedstock))
                    usedcolor.Add(pusedcolor(pusedstock))
                    usedsize1.Add(pusedsize1(pusedstock))
                    usedcount.Add(pusedcount(pusedstock))
                    usedinternalID.Add(pusedinternalID(pusedstock))
                    usedsaw.Add(pusedsaw(pusedstock))
                    usedminsize.Add(pusedminsize(pusedstock))
                    stock_exists(it) = True
                    '
                    ' Selects which used stock to use
                    '
                    cmd.CommandText = "SELECT size, count, internalID, context2 FROM stockUsed"
                    cmd.ExecuteNonQuery()
                    readerObj = cmd.ExecuteReader
                    While readerObj.Read
                        If String.Equals(pusedinternalID(pusedstock), readerObj("internalID").ToString) Then
                            xusedsize1.Add(readerObj("size").ToString)
                            xusedcount.Add(readerObj("count").ToString)
                            xusedinternalID.Add(readerObj("internalID").ToString)
                            xusedcontext2.Add(readerObj("context2").ToString)

                        End If
                    End While
                    readerObj.Close()
                Else
                    usedstockID1.Add("")
                    usedstockID2.Add("")
                    usedstockID3.Add("")
                    useddescription.Add("")
                    usedcolor.Add("")
                    usedsize1.Add("")
                    usedcount.Add("")
                    usedinternalID.Add("")
                    usedsaw.Add("")
                    usedminsize.Add("")
                End If
            Else
                usedstockID1.Add("")
                usedstockID2.Add("")
                usedstockID3.Add("")
                useddescription.Add("")
                usedcolor.Add("")
                usedsize1.Add("")
                usedcount.Add("")
                usedinternalID.Add("")
                usedsaw.Add("")
                usedminsize.Add("")
            End If
        Next

        '
        ' Checks if any parts are omited
        '
        Dim frm2 As OmitParts = New OmitParts()
        frm2.des = partlistd
        If frm2.ShowDialog() = DialogResult.OK Then
            For it = 0 To frm2.omitted.Count - 1
                Dim intloc As Integer = partlistd.IndexOf(frm2.omitted(it))

                partlistd.Remove(frm2.omitted(it))

                For it1 = 0 To xusedinternalID.Count - 1
                    If String.Equals(xusedinternalID(it1), usedinternalID(intloc)) Then
                        xusedsize1.RemoveAt(intloc)
                        xusedcount.RemoveAt(intloc)
                        xusedinternalID.RemoveAt(intloc)
                        xusedcontext2.RemoveAt(intloc)
                    End If
                Next

                usedstockID1.RemoveAt(intloc)
                usedstockID2.RemoveAt(intloc)
                usedstockID3.RemoveAt(intloc)
                useddescription.RemoveAt(intloc)
                usedcolor.RemoveAt(intloc)
                usedsize1.RemoveAt(intloc)
                usedcount.RemoveAt(intloc)
                usedinternalID.RemoveAt(intloc)
                usedsaw.RemoveAt(intloc)
                usedminsize.RemoveAt(intloc)
                stock_exists.RemoveAt(intloc)
            Next
        End If
        '
        ' Sort by saw
        '
        Dim switchpos As ArrayList = New ArrayList
        Dim alreadyin As ArrayList = New ArrayList
        Dim sorted As SortedList(Of String, Integer) = New SortedList(Of String, Integer)
        For it = 0 To usedsaw.Count - 1
            sorted.Add(usedsaw(it) + it.ToString, it)
        Next
        ' Loop over pairs in the collection.
        For Each pair As KeyValuePair(Of String, Integer) In sorted
            switchpos.Add(pair.Value)
        Next

        Dim tusedstockid1 As ArrayList = New ArrayList
        Dim tusedstockid2 As ArrayList = New ArrayList
        Dim tusedstockid3 As ArrayList = New ArrayList
        Dim tuseddescription As ArrayList = New ArrayList
        Dim tusedcolor As ArrayList = New ArrayList
        Dim tusedsize1 As ArrayList = New ArrayList
        Dim tusedcount As ArrayList = New ArrayList
        Dim tusedinternalID As ArrayList = New ArrayList
        Dim tusedsaw As ArrayList = New ArrayList
        Dim tusedminsize As ArrayList = New ArrayList
        Dim tstock_exists As ArrayList = New ArrayList

        For it = 0 To usedstockID1.Count - 1
            tusedstockid1.Add(usedstockID1(it))
            tusedstockid2.Add(usedstockID2(it))
            tusedstockid3.Add(usedstockID3(it))
            tuseddescription.Add(useddescription(it))
            tusedcolor.Add(usedcolor(it))
            tusedsize1.Add(usedsize1(it))
            tusedcount.Add(usedcount(it))
            tusedinternalID.Add(usedinternalID(it))
            tusedsaw.Add(usedsaw(it))
            tusedminsize.Add(usedminsize(it))
            tstock_exists.Add(stock_exists(it))
        Next

        usedstockID1.Clear()
        usedstockID2.Clear()
        usedstockID3.Clear()
        useddescription.Clear()
        usedcolor.Clear()
        usedsize1.Clear()
        usedcount.Clear()
        usedinternalID.Clear()
        usedsaw.Clear()
        usedminsize.Clear()
        stock_exists.Clear()

        For it = 0 To switchpos.Count - 1
            usedstockID1.Add(tusedstockid1(switchpos(it)))
            usedstockID2.Add(tusedstockid2(switchpos(it)))
            usedstockID3.Add(tusedstockid3(switchpos(it)))
            useddescription.Add(tuseddescription(switchpos(it)))
            usedcolor.Add(tusedcolor(switchpos(it)))
            usedsize1.Add(tusedsize1(switchpos(it)))
            usedcount.Add(tusedcount(switchpos(it)))
            usedinternalID.Add(tusedinternalID(switchpos(it)))
            usedsaw.Add(tusedsaw(switchpos(it)))
            usedminsize.Add(tusedminsize(switchpos(it)))
            stock_exists.Add(tstock_exists(switchpos(it)))
        Next
        '
        ' Loops for individual parts
        '
        Label14.Text = "Calculating Cuts..."
        For it = 0 To partlistd.Count - 1
            Label14.Text = "Calculating Cuts..."
            Dim Calculator As CutGLib.CutEngine
            Calculator = New CutGLib.CutEngine
            'Calculator.SetComputerLicenseKey("1234")
            Dim con As New SqlConnection
            Dim cmd As New SqlCommand
            con.ConnectionString = connectionstring.connect1
            con.Open()
            cmd.Connection = con
            Dim readerObj As SqlClient.SqlDataReader
            '
            ' Adds Stocks
            '
            If xusedinternalID.Count > 0 Then
                For it1 = 0 To xusedinternalID.Count - 1
                    If String.Equals(xusedinternalID(it1), usedinternalID(it)) Then
                        Dim temp3 As Double = Convert.ToDouble(xusedsize1(it1))
                        Dim temp4 As Integer = Convert.ToInt32(xusedcount(it1))
                        Calculator.AddLinearStock(temp3, temp4)
                    End If

                Next
            End If
            If stock_exists(it) Then
                newstockinternalID.Add(usedinternalID(it))
                Dim temp1 As Double = Convert.ToDouble(usedsize1(it))
                Dim temp2 As Integer = Convert.ToInt32(usedcount(it))
                Calculator.AddLinearStock(temp1, temp2)
            Else
                '
                ' Adds Page if can't find stock
                '
                newstockinternalID.Add(0)
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
            ' Calculates Cuts 
            '

            Dim result As String = "1"
            If stock_exists(it) Then
                result = Calculator.ExecuteLinear()

            End If
            '
            ' Formats Outputs
            '
            If (result = "") Then
                itWorks = True
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
                    '
                    ' Gets count of stock used
                    '
                    Dim remainder As Double = 0
                    Calculator.GetLayoutInfo(iLayout, StockIndex, VStockCount)

                    Dim slength As Double
                    Calculator.GetLinearStockInfo(StockIndex, slength, StockActive)
                    If String.Equals(slength.ToString, usedsize1(it)) Then
                        itemcount += VStockCount
                    End If
                    If Not String.Equals(slength.ToString, usedsize1(it)) Then
                        itemcount2 += VStockCount
                    End If

                    'Console.WriteLine("Layout={0}:  Start Stock={1};  Count of Stock={2}", iLayout, StockIndex, VStockCount)

                    '
                    ' Determines if 1 or 2 layouts per page
                    '
                    Dim over20 As Boolean = False
                    Dim page As PdfPage
                    Dim gfx As XGraphics
                    partCount = Calculator.GetPartCountOnStock(iStock)
                    If partCount > 18 Then
                        page = document.AddPage
                        gfx = XGraphics.FromPdfPage(page)
                        over20 = True
                        font = New XFont("Verdana", 10, XFontStyle.Regular)
                        sectioncount = 0

                    ElseIf partCount > 8 Then
                        page = document.AddPage
                        gfx = XGraphics.FromPdfPage(page)
                        pos2 = 0
                        sectioncount = 0
                        font = New XFont("Verdana", 11, XFontStyle.Regular)
                    Else
                        font = New XFont("Verdana", 11, XFontStyle.Regular)
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
                    '
                    ' Prints each Layout
                    '
                    iStock = StockIndex
                    If Calculator.GetLinearStockInfo(StockIndex, StockLength, StockActive) Then
                        partCount = Calculator.GetPartCountOnStock(iStock)

                        Dim pageNo As String = Convert.ToString(iLayout + 1)

                        Calculator.CreateStockImage(iStock, "test" + picturecount.ToString + ".png", 1000)
                        Dim image As XImage = XImage.FromFile("test" + picturecount.ToString + ".png")
                        picturecount += 1
                        imagecountfile.Add(image)

                        Dim Stockt As String = Convert.ToString(iStock + 1)
                        Dim xt As String = Convert.ToString(StockLength)
                        Dim temp As String = Convert.ToString(VStockCount)
                        Dim temp1 As String = Convert.ToString(iLayout + 1)
                        Dim pos1 As Integer = 0
                        Dim sizestring As String = xt


                        If Not String.Equals(slength.ToString, usedsize1(it)) Then
                            sizestring = xt + "  Used Part Check Bin"
                        End If

                        Dim descriptionString As String = partlistd(it)
                        Dim descriptionString2 As String = "ID: " + usedstockID2(it)
                        Dim descriptionString3 As String = "ID2: " + usedstockID3(it)
                        Dim descriptionString4 As String = "Stock Size: " + sizestring + "  Color: " + usedcolor(it)
                        Dim descriptionstring5 As String = "Saw Used: " + usedsaw(it)
                        Dim descriptionstring6 As String = "# Of Stock Cut: " + temp
                        Dim descriptionstring7 As String = "Layout Number = " + temp1
                        Dim leftanchor As Integer = 40

                        gfx.DrawString("Page " + pageNo, font2, XBrushes.Black, New XRect(leftanchor, 50 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString, font3, XBrushes.Black, New XRect(leftanchor, 65 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString2, font2, XBrushes.Black, New XRect(leftanchor, 80 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString3, font2, XBrushes.Black, New XRect(leftanchor, 95 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionString4, font2, XBrushes.Black, New XRect(leftanchor, 110 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionstring5, font2, XBrushes.Black, New XRect(leftanchor + 355, 65 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionstring6, font2, XBrushes.Black, New XRect(leftanchor + 355, 80 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(descriptionstring7, font2, XBrushes.Black, New XRect(leftanchor + 355, 95 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)

                        Console.WriteLine("Layout={0}:  Length={1}", iStock, StockLength)
                        Dim lengthlist As ArrayList = New ArrayList
                        ' Iterate by parts and get indices of cut parts
                        For ViPart = 0 To partCount - 1
                            ' Get global part index of ViPart cut from the current stock
                            partIndex = Calculator.GetPartIndexOnStock(iStock, ViPart)
                            ' Get length and location of the part
                            ' X - coordinate on the stock where the part beggins.

                            Calculator.GetResultLinearPart(partIndex, tmp, partLength, VX)
                            ' Output the part information
                            Console.WriteLine("Part= {0}:  X={1}  Length={2}", partIndex, VX, partLength)

                            '
                            ' Prints parts and positions
                            '
                            gfx.DrawString("Cut position= " + VX.ToString + " Length= " + partLength.ToString, font, XBrushes.Black, New XRect(leftanchor, 135 + pos1 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            '
                            ' Prints corresponding Quote window line
                            '
                            Dim correspondingnum As String = ""
                            Dim tempstring As ArrayList = New ArrayList
                            Dim layoutit As Integer = VStockCount
                            Dim remove1 As Integer = 0
                            Dim remove2 As Integer = 0


                            For it1 = 0 To uosize.Count - 1
                                If layoutit > 0 Then
                                    If partLength = uosize(it1) AndAlso String.Equals(partlistid(it), uopartID(it1)) Then
                                        For it2 = uocountdown(it1) To uoitemQuantity(it1)
                                            If layoutit > 0 Then
                                                Dim remove3 As Integer = 0

                                                For it3 = 1 To uocountdownrep(it1)
                                                    If layoutit > 0 Then
                                                        tempstring.Add(uolistorderline(it1) + "-" + uoitemNumber(it1).ToString + "-" + it2.ToString)
                                                        layoutit -= 1
                                                        remove2 += 1
                                                        remove3 += 1
                                                    End If
                                                    If layoutit = 0 Then
                                                    End If
                                                Next
                                                'uocountdownrep(it1) -= remove3
                                                'remove3 = 0
                                            End If
                                        Next
                                        uocountdown(it1) += remove2
                                        remove2 = 0
                                    End If

                                End If
                            Next


                            Dim count1 As Integer = 1
                            Dim line1 As Boolean = True
                            Dim line2 As Boolean = False
                            For it1 = 0 To tempstring.Count - 1
                                If count1 < 3 Then
                                    correspondingnum = correspondingnum + tempstring(it1) + " "
                                    count1 += 1
                                Else
                                    correspondingnum = ""
                                    correspondingnum = tempstring(it1) + " "
                                    count1 = 0
                                    line2 = True
                                End If

                                If line1 AndAlso (count1 = 3 Or it1 = tempstring.Count - 1) Then
                                    gfx.DrawString("Orders: " + correspondingnum, font, XBrushes.Black, New XRect(leftanchor + 260, 135 + pos1 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                                    If over20 Then
                                        pos1 += 10
                                    Else
                                        pos1 += 14
                                    End If
                                    line1 = False
                                ElseIf line2 AndAlso (count1 = 4 Or it1 = tempstring.Count - 1) Then
                                    gfx.DrawString("  " + correspondingnum, font, XBrushes.Black, New XRect(leftanchor, 135 + pos1 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                                    If over20 Then
                                        pos1 += 10
                                    Else
                                        pos1 += 14
                                    End If
                                End If
                            Next

                            If VX + partLength > remainder Then
                                remainder = VX + partLength
                            End If
                        Next ViPart

                        If usedsize1(it) - remainder > usedminsize(it) Then
                            gfx.DrawString("Save Remainder Piece ", font, XBrushes.Black, New XRect(leftanchor, 135 + pos1 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                            If over20 Then
                                pos1 += 10
                            Else
                                pos1 += 14
                            End If
                        End If

                        gfx.DrawImage(image, leftanchor, 145 + pos1 + pos2 - 10, 520, 20)
                        '
                        ' Prints out which quote and line number the part corresponds to
                        '
                        'For it2 = 0 To lengthlist.Count - 1
                        '    Dim correspondingnum As String = "Orders: "
                        '    Dim tempstring As ArrayList = New ArrayList
                        '    For it1 = 0 To uosize.Count - 1
                        '        If lengthlist(it2) = uosize(it1) AndAlso Not tempstring.Contains(uolistorderline(it1) + "-" + uoitemNumber(it1).ToString + " ") Then
                        '            For it3 = 0 To uocount(it1) / uoitemQuantity(it1)
                        '                tempstring.Add(uolistorderline(it1) + "-" + uoitemNumber(it1).ToString + "  " + uocountdown(it1))
                        '                uocountdown(it1) -= 1
                        '            Next
                        '        End If
                        '    Next
                        '    Dim count1 As Integer = 1
                        '    Dim line1 As Boolean = True
                        '    Dim line2 As Boolean = False
                        '    For it1 = 0 To tempstring.Count - 1
                        '        If count1 < 4 Then
                        '            correspondingnum = correspondingnum + tempstring(it1)
                        '            count1 += 1
                        '        Else
                        '            correspondingnum = ""
                        '            correspondingnum = tempstring(it1)
                        '            count1 = 0
                        '            line2 = True
                        '        End If

                        '        If line1 AndAlso (count1 = 4 Or it1 = tempstring.Count - 1) Then
                        '            gfx.DrawString("Size: " + lengthlist(it2).ToString + " " + correspondingnum, font, XBrushes.Black, New XRect(leftanchor, 170 + pos1 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        '            If over20 Then
                        '                pos1 += 10
                        '            Else
                        '                pos1 += 14
                        '            End If
                        '            line1 = False
                        '        ElseIf line2 AndAlso (count1 = 4 Or it1 = tempstring.Count - 1) Then
                        '            gfx.DrawString("  " + correspondingnum, font, XBrushes.Black, New XRect(leftanchor, 170 + pos1 + pos2 - 10, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        '            If over20 Then
                        '                pos1 += 10
                        '            Else
                        '                pos1 += 14
                        '            End If
                        '        End If
                        '    Next
                        'Next

                        '
                        ' Saves list of used stock that was actually used
                        '
                        If StockLength <> usedsize1(it) Then
                            cmd.CommandText = "SELECT stockID2, size, count, internalID, context2  FROM stockUsed"
                            cmd.ExecuteNonQuery()


                            readerObj = cmd.ExecuteReader
                            While readerObj.Read
                                If String.Equals(usedinternalID(it), readerObj("internalID").ToString) And StockLength = readerObj("size") Then
                                    usedcontext2.Add(readerObj("context2").ToString)
                                    usedcontextamount.Add(temp)
                                End If
                            End While
                            readerObj.Close()
                        End If

                        '
                        ' Saves remainder parts to table
                        '
                        If usedsize1(it) - remainder > usedminsize(it) Then
                            Dim internalID As ArrayList = New ArrayList
                            Dim rn As New Random
                            Dim it1 As Integer
                            Dim boolinternalID As Boolean = True
                            Dim inputusedID As Integer
                            Dim stockid2 As ArrayList = New ArrayList
                            Dim size As ArrayList = New ArrayList
                            Dim count As ArrayList = New ArrayList
                            Dim boolupdatetable As Boolean = False

                            cmd.CommandText = "SELECT stockID2, size, count, context2  FROM stockUsed"
                            cmd.ExecuteNonQuery()

                            readerObj = cmd.ExecuteReader
                            While readerObj.Read
                                stockid2.Add(readerObj("stockID2").ToString)
                                size.Add(readerObj("size").ToString)
                                count.Add(readerObj("count"))
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

                            Dim temp3 As String = Convert.ToString(inputusedID)
                            Dim temp2 As String = Convert.ToString(usedsize1(it) - remainder)
                            If stockid2.Count > 1 Then
                                For it2 = 0 To stockid2.Count - 1
                                    If String.Equals(stockid2(it2), usedstockID2(it)) AndAlso String.Equals(size(it2), temp2) Then
                                        count(it2) += Convert.ToInt64(temp)
                                        cmd.CommandText = "UPDATE stockUsed SET count = " + count(it2).ToString + " WHERE context2 = " + internalID(it2).ToString
                                        cmd.ExecuteNonQuery()
                                        usedstockinternalID.Add(internalID(it2).ToString)
                                        usedstockinternalamount.Add(count(it2).ToString)
                                        boolupdatetable = True
                                    End If
                                Next
                            End If

                            If Not boolupdatetable Then
                                cmd.CommandText = "INSERT INTO stockUsed VALUES('" + usedstockID1(it) + "', '" + usedstockID2(it) + "' , '" + usedstockID3(it) + "', '" + useddescription(it) + "' , '" + usedcolor(it) + "', " + temp2 + ", " + temp + ", " + usedinternalID(it) + " , '' , '" + usedsaw(it) + "' , " + temp3 + ", '')"
                                cmd.ExecuteNonQuery()
                                usedstockinternalIDcreated.Add(inputusedID)
                            End If
                        End If
                    End If

                Next iLayout
                excountlist.Add(itemcount)
                excountlistu.Add(itemcount2)
                excountwork.Add(0)

            ElseIf stock_exists(it) Then
                itWorks = True
                Console.WriteLine("calculate fail")
                Dim page As PdfPage = document.AddPage
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
                gfx.DrawString("Not enough stock: " + partlistd(it), font2, XBrushes.Black, New XRect(50, 50, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                excountlist.Add(usedcount(it))
                excountlistu.Add(0)
                excountwork.Add(2)
            End If
            Console.WriteLine(result)
            Calculator.Clear()
        Next

        '
        ' Adds stock List in ouput
        '
        If itWorks Then
            Dim page As PdfPage = document.AddPage
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim pos1 As Integer = 0
            gfx.DrawString("Total Part List", font3, XBrushes.Black, New XRect(50, 50, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
            For it = 0 To partlistd.Count - 1
                If excountwork(it) = 0 Then
                    gfx.DrawString((it + 1).ToString + ". New: " + excountlist(it).ToString + " " + "Used: " + excountlistu(it).ToString + " " + partlistd(it) + " ID:" + partlistid(it), font, XBrushes.Black, New XRect(40, 70 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    pos1 += 15
                ElseIf excountwork(it) = 1 Then
                    gfx.DrawString((it + 1).ToString + ". Not Cut, Partcount: " + excountlist(it).ToString + "  " + partlistd(it) + " ID:" + partlistid(it), font, XBrushes.Black, New XRect(40, 70 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    pos1 += 15
                ElseIf excountwork(it) = 2 Then
                    gfx.DrawString((it + 1).ToString + ". Not Enough, current Stock: " + excountlist(it).ToString + "  " + partlistd(it) + " ID:" + partlistid(it), font, XBrushes.Black, New XRect(40, 70 + pos1, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    pos1 += 15
                End If

            Next

            '
            ' System Cleanup
            '
            Label14.Text = "Request Complete"
            document.Save(filename)
            document.Close()
            Process.Start(filename)

            For it = 0 To picturecount - 1
                imagecountfile(it).Dispose()
                My.Computer.FileSystem.DeleteFile("test" + it.ToString + ".png")
            Next

            '
            ' Removes used parts and stocks from tables
            '
            Dim result1 As DialogResult = MessageBox.Show("Confirm Cuts and Update Database", "Confirm Cuts", MessageBoxButtons.YesNo)
            If result1 = DialogResult.Yes Then
                Dim con As New SqlConnection
                Dim cmd As New SqlCommand
                con.ConnectionString = connectionstring.connect1
                con.Open()
                cmd.Connection = con

                For it = 0 To partsinternalID.Count - 1
                    cmd.CommandText = "DELETE FROM parts WHERE internalID = " + partsinternalID(it).ToString
                    cmd.ExecuteNonQuery()
                Next
                For it = 0 To newstockinternalID.Count - 1
                    If excountwork(it) = 0 Then
                        cmd.CommandText = "SELECT count, internalID FROM stockNew"
                        cmd.ExecuteNonQuery()
                        Dim count As Integer
                        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
                        While readerObj.Read
                            If String.Equals(readerObj("internalID").ToString, newstockinternalID(it)) Then
                                count = readerObj("count")
                            End If
                        End While
                        readerObj.Close()
                        count = count - excountlist(it)

                        cmd.CommandText = "Update stockNew SET count = " + count.ToString + " WHERE internalID = " + newstockinternalID(it).ToString
                        cmd.ExecuteNonQuery()
                    End If
                Next
                For it = 0 To usedcontext2.Count - 1
                    If excountwork(it) = 0 Then
                        cmd.CommandText = "SELECT count, context2 FROM stockUsed"
                        cmd.ExecuteNonQuery()
                        Dim count As Integer
                        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
                        While readerObj.Read
                            If String.Equals(readerObj("context2").ToString, usedcontext2(it)) Then
                                count = readerObj("count")
                            End If
                        End While
                        readerObj.Close()
                        count = count - Convert.ToInt64(usedcontextamount(it))
                        If count > 0 Then
                            cmd.CommandText = "UPDATE stockUsed SET count = " + count.ToString + " WHERE context2 = " + usedcontext2(it)
                            cmd.ExecuteNonQuery()
                        Else
                            cmd.CommandText = "DELETE FROM stockUsed WHERE context2 = " + usedcontext2(it)
                            cmd.ExecuteNonQuery()
                        End If
                    End If
                Next
                Me.Close()
            Else
                '
                ' If you do not want to remove parts and stocks from tables.
                '
                Dim con As New SqlConnection
                Dim cmd As New SqlCommand
                con.ConnectionString = connectionstring.connect1
                con.Open()
                cmd.Connection = con
                For it = 0 To usedstockinternalIDcreated.Count - 1
                    cmd.CommandText = "DELETE FROM stockUsed WHERE context2 = " + usedstockinternalIDcreated(it).ToString
                    cmd.ExecuteNonQuery()
                Next

                cmd.CommandText = "SELECT count, context2 FROM stockUsed"
                cmd.ExecuteNonQuery()
                Dim count As Integer

                For it = 0 To usedstockinternalID.Count - 1
                    Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader
                    While readerObj.Read
                        If String.Equals(readerObj("context2").ToString, usedcontext2(it)) Then
                            count = readerObj("count")
                        End If
                    End While
                    readerObj.Close()
                    count -= usedcontextamount(it)
                    cmd.CommandText = "UPDATE stockUsed SET count = " + count.ToString + " WHERE context2 = " + usedstockinternalID(it).ToString
                    cmd.ExecuteNonQuery()
                Next
                Me.Close()
            End If

        End If
    End Sub
End Class