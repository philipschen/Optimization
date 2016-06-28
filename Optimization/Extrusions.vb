Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

Public Class Extrusions

    Dim stockID1 As ArrayList = New ArrayList
    Dim stockID2 As ArrayList = New ArrayList
    Dim stockID3 As ArrayList = New ArrayList
    Dim description As ArrayList = New ArrayList
    Dim color As ArrayList = New ArrayList
    Dim size As ArrayList = New ArrayList
    Dim internalID As ArrayList = New ArrayList
    Dim Used As ArrayList = New ArrayList

    Private Sub Extrusions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        cmd.CommandText = "SELECT stockID1, stockID2, stockID3, description, color, size FROM stockNew"
        cmd.ExecuteNonQuery()
        Dim readerObj As SqlClient.SqlDataReader = cmd.ExecuteReader

        '
        ' Reads data from database 
        '

        While readerObj.Read
            Dim tempSize As Integer = Convert.ToInt32(readerObj("Size").ToString)

            ComboBox1.Items.Add(readerObj("stockID1").ToString)
            ComboBox2.Items.Add(readerObj("stockID2").ToString)
            ComboBox3.Items.Add(readerObj("stockID3").ToString)
            description.Add(readerObj("description").ToString)

            stockID1.Add(readerObj("stockID1").ToString)
            stockID2.Add(readerObj("stockID2").ToString)
            stockID3.Add(readerObj("stockID3").ToString)
            color.Add(readerObj("color").ToString)
            size.Add(readerObj("size").ToString)
            ' internalID.Add(readerObj("internalID").ToString)


            If readerObj.IsDBNull(1) Then
                readerObj.Close()
                Exit While
            End If
        End While

    End Sub

    Private Sub comboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox1.SelectedText, stockID1(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            End If
        Next

    End Sub

    Private Sub comboBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedValueChanged

        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox2.SelectedText, stockID2(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)

            End If
        Next

    End Sub

    Private Sub comboBox3_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedValueChanged

        For it1 = 0 To description.Count - 1

            If String.Equals(ComboBox3.SelectedText, stockID3(it1)) And Not Used.Contains(it1) Then
                ListBox2.Items.Add(description(it1) + " " + color(it1) + " ")
                Used.Add(it1)
            End If
        Next

    End Sub

    Private Sub ListBox2_SelectedValueChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedValueChanged
        Dim pString As String = ""
        If ListBox2.SelectedIndex = 0 Then
            pString = "Cannot find piece, Click Input to input new stock"
        End If
        If ListBox2.SelectedIndex > 0 Then
            pString = description(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + color(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + size(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + stockID1(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + stockID2(Used(ListBox2.SelectedIndex - 1)) + vbCrLf + stockID3(Used(ListBox2.SelectedIndex - 1))

        End If
        RichTextBox1.Text = pString
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If ListBox2.SelectedItem = "New Part" Then
            Dim frmNewExtr As NewExtrusions = New NewExtrusions()
            frmNewExtr.ShowDialog()

        ElseIf ListBox2.SelectedIndex >= 0 Then

        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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
        ' Create an array to set multiple values at once.
        Dim students(5, 2) As String
        students(0, 0) = "Zara"
        students(0, 1) = "Ali"
        students(1, 0) = "Nuha"
        students(1, 1) = "Ali"
        students(2, 0) = "Arilia"
        students(2, 1) = "RamKumar"
        students(3, 0) = "Rita"
        students(3, 1) = "Jones"
        students(4, 0) = "Umme"
        students(4, 1) = "Ayman"
        ' Fill A2:B6 with an array of values (First and Last Names).
        shXL.Range("A2", "B6").Value = students
        ' Fill C2:C6 with a relative formula (=A2 & " " & B2).
        raXL = shXL.Range("C2", "C6")
        raXL.Formula = "=A2 & "" "" & B2"
        ' Fill D2:D6 values.
        With shXL
            .Cells(2, 4).Value = "Biology"
            .Cells(3, 4).Value = "Mathmematics"
            .Cells(4, 4).Value = "Physics"
            .Cells(5, 4).Value = "Mathmematics"
            .Cells(6, 4).Value = "Arabic"
        End With
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


End Class